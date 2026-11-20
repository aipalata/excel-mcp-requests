import { z } from "zod";
import { WorkspaceError } from "./config";
import { statuses, type Plan } from "./contracts";

const planSchema = z.strictObject({
  id: z.uuid(),
  ticket: z.strictObject({ id: z.number().int(), client: z.string(), problem: z.string(), status: z.enum(statuses) }),
  status: z.enum(statuses),
  expiresAt: z.number().int(),
});

export function cloudWorkbookWritable() {
  return Boolean(process.env.VERCEL && process.env.MCP_MODE === "http" &&
    process.env.EXCEL_CLOUD_READ_ONLY === "0" && process.env.WORKBOOK_WRITE_TOKEN?.trim());
}

async function callWorkbookState<T>(name: string, payload: Record<string, unknown>, responseSchema: z.ZodType<T>) {
  const project = process.env.WORKBOOK_SUPABASE_PROJECT;
  const key = process.env.WORKBOOK_SUPABASE_PUBLISHABLE_KEY;
  const token = process.env.WORKBOOK_WRITE_TOKEN;
  if (!project || !/^[a-z0-9]{20}$/.test(project) || !key || !token) {
    throw new WorkspaceError("Облачное хранилище заявок ожидает настройки.");
  }
  try {
    const response = await fetch(`https://${project}.supabase.co/rest/v1/rpc/${name}`, {
      method: "POST", cache: "no-store", redirect: "error", signal: AbortSignal.timeout(5_000),
      headers: { apikey: key, "x-workbook-write": token, "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!response.ok) throw new Error("Supabase RPC failed");
    return responseSchema.parse(await response.json());
  } catch {
    throw new WorkspaceError("Не удалось обратиться к облачному хранилищу заявок. Повторите запрос.");
  }
}

export function saveCloudPlan(session: string, plan: Plan) {
  return callWorkbookState("requests_status_plan", { session, plan: plan.id, payload: plan }, planSchema);
}

export async function takeCloudPlan(session: string, plan: string) {
  const result = await callWorkbookState("requests_status_plan", { session, plan }, planSchema.nullable());
  if (!result || result.expiresAt <= Date.now()) {
    throw new WorkspaceError("План недоступен. Запросите изменение статуса ещё раз.");
  }
  return result;
}

export function cloudChatAllowed() {
  return callWorkbookState("requests_chat_allowed", {}, z.boolean());
}

export const cloudWriteReceipt = z.object({
  at: z.string().regex(/^[AD][0-9]{1,4}$/), ticketId: z.number().int().positive().nullable(), replayed: z.boolean(),
});

export function getCloudWriteReceipt(operation: string, requestHash: string) {
  return callWorkbookState("requests_workbook_state", { operation, request_hash: requestHash },
    z.union([cloudWriteReceipt, z.object({ pending: z.literal(true) }), z.object({ conflict: z.literal(true) })]));
}
