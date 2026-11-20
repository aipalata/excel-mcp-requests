import { randomUUID } from "node:crypto";
import { NextRequest } from "next/server";
import { z } from "zod";
import { rememberDecision, runAgent } from "@/lib/agent";
import { mcpConfig, openAIConfig, workbookName, WorkbookError, WorkspaceError } from "@/lib/config";
import type { Outcome, WireEvent } from "@/lib/contracts";
import { withMcp } from "@/lib/mcp";
import { applyStatus, consumePlan, serialized, snapshot, withTurn } from "@/lib/operations";
import { tracer, type Trace } from "@/lib/trace";
import { readTickets } from "@/lib/workbook";

export const runtime = "nodejs";
export const maxDuration = 60;

const uuid = z.uuid();
const actionSchema = z.discriminatedUnion("kind", [
  z.strictObject({ kind: z.literal("refresh") }),
  z.strictObject({ kind: z.literal("chat"), message: z.string().trim().min(1).max(1600), requestId: uuid }),
  z.strictObject({ kind: z.literal("confirm"), planId: uuid }),
  z.strictObject({ kind: z.literal("cancel"), planId: uuid }),
]);
type Action = z.infer<typeof actionSchema>;

async function boundedBody(request: NextRequest) {
  if (!request.body || !request.headers.get("content-type")?.startsWith("application/json")) return;
  const reader = request.body.getReader();
  const decoder = new TextDecoder();
  let size = 0;
  let body = "";
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      size += value.length;
      if (size > 12_288) {
        await reader.cancel();
        return;
      }
      body += decoder.decode(value, { stream: true });
    }
    return actionSchema.safeParse(JSON.parse(body + decoder.decode()));
  } catch {
    return;
  } finally {
    reader.releaseLock();
  }
}

async function perform(action: Action, sessionId: string, trace: Trace): Promise<Outcome> {
  const signal = AbortSignal.timeout(48_000);
  trace.mark("request", {
    refresh: "Запрошено обновление таблицы",
    chat: "Получена команда пользователя",
    confirm: "Пользователь нажал «Подтвердить изменение»",
    cancel: "Пользователь нажал «Отменить»",
  }[action.kind]);

  if (action.kind === "refresh") {
    return serialized(() => withMcp(mcpConfig(), trace, async mcp =>
      snapshot(await readTickets(mcp)), signal), signal);
  }

  if (action.kind === "chat" && !openAIConfig().apiKey) {
    throw new WorkspaceError("OpenAI не настроен. Добавьте API-ключ в .env.local и перезапустите приложение.");
  }

  return withTurn(sessionId, async () => {
    if (action.kind === "chat") {
      return withMcp(mcpConfig(), trace, mcp => runAgent(mcp, trace, {
        sessionId, requestId: action.requestId, message: action.message, signal,
      }), signal);
    }

    const plan = consumePlan(action.planId, sessionId);
    trace.mark("confirm", action.kind === "cancel" ? "План отменён" : "План подтверждён сервером",
      { ticket: plan.ticket.id, status: plan.status });
    if (action.kind === "cancel") {
      const message = `Изменение заявки №${plan.ticket.id} отменено.`;
      rememberDecision(sessionId, message);
      return { message };
    }

    const result = await serialized(() => withMcp(mcpConfig(), trace, mcp => applyStatus(mcp, plan), signal), signal);
    if (result.message) rememberDecision(sessionId, result.message);
    return result;
  });
}

export async function POST(request: NextRequest) {
  const cloud = Boolean(process.env.VERCEL);
  const host = request.headers.get("host") || "";
  const origin = request.headers.get("origin");
  const localHost = /^(?:localhost|127\.0\.0\.1|\[::1\])(?::\d{1,5})?$/.test(host);
  const deploymentHost = process.env.VERCEL_URL;
  const protectedHost = deploymentHost && deploymentHost !== process.env.VERCEL_PROJECT_PRODUCTION_URL && host === deploymentHost;
  if (cloud && (!protectedHost || process.env.EXCEL_CLOUD_READ_ONLY !== "1")) {
    return Response.json({ error: "Книга доступна на защищённом адресе развёртывания Vercel." }, { status: 403 });
  }
  if (cloud && !process.env.MCP_HTTP_TOKEN?.trim()) {
    return Response.json({ error: "Облачная книга ожидает настройки доступа." }, { status: 503 });
  }
  if ((!cloud && !localHost) || (origin && origin !== `${cloud ? "https" : "http"}://${host}`) ||
      request.headers.get("sec-fetch-site") === "cross-site") {
    return Response.json({ error: "Запрос с другого источника отклонён." }, { status: 403 });
  }
  const payload = await boundedBody(request);
  if (!payload?.success) return Response.json({ error: "Проверьте тип и длину запроса." }, { status: 400 });
  if (cloud && payload.data.kind !== "refresh") {
    return Response.json({ error: "На Vercel доступно только чтение книги." }, { status: 403 });
  }

  const existing = uuid.safeParse(request.cookies.get("requests_session")?.value);
  const sessionId = existing.success ? existing.data : randomUUID();
  const encoder = new TextEncoder();
  const operationId = randomUUID();
  let closed = false;
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      const send = (event: WireEvent) => {
        if (!closed) controller.enqueue(encoder.encode(JSON.stringify(event) + "\n"));
      };
      const trace = tracer(operationId, event => send({ type: "trace", event }));
      void (async () => {
        try {
          send({
            type: "config", openAI: !cloud && Boolean(openAIConfig().apiKey), file: workbookName(),
            mode: process.env.MCP_MODE || "stdio", readOnly: cloud,
          });
          const result = await perform(payload.data, sessionId, trace);
          send({ type: "result", result });
        } catch (error) {
          const message = error instanceof WorkspaceError ? error.message : "Операция прервана. Проверьте трассировку и обновите таблицу.";
          if (payload.data.kind === "refresh") {
            send({ type: "result", result: { workbookError: message } });
          } else {
            trace.start("response", "Операция остановлена").error(message);
            send({ type: "result", result: { error: message, workbookError: error instanceof WorkbookError ? message : undefined } });
          }
        } finally {
          if (!closed) controller.close();
        }
      })();
    },
    cancel() { closed = true; },
  });

  const headers = new Headers({
    "Content-Type": "application/x-ndjson; charset=utf-8",
    "Cache-Control": "no-store, no-transform",
    "X-Content-Type-Options": "nosniff",
  });
  if (!existing.success) {
    const secure = request.nextUrl.protocol === "https:" ? "; Secure" : "";
    headers.set("Set-Cookie", `requests_session=${sessionId}; Path=/; HttpOnly; SameSite=Strict; Max-Age=86400${secure}`);
  }
  return new Response(stream, { headers });
}
