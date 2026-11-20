import type { Stage, TraceEvent } from "./contracts";

function clean(value: unknown, depth = 0): unknown {
  if (depth > 18) return "[…]";
  if (typeof value === "string") {
    let safe = value;
    for (const secret of [process.env.OPENAI_API_KEY, process.env.MCP_HTTP_TOKEN].map(value => value?.trim())) {
      if (secret) safe = safe.replaceAll(secret, "[скрыто]");
    }
    return safe
      .replace(/\bsk-[A-Za-z0-9_-]{12,}\b/g, "[скрыто]")
      .replace(/\bBearer\s+[^\s"',]+/gi, "Bearer [скрыто]")
      .replace(/https?:\/\/[^\s"'<>]+/gi, "[URL скрыт]")
      .replace(/\/(?:Users|home|private|var|tmp|srv|mnt|data)\/[^\n"']+/g, "[локальный путь]")
      .slice(0, 1800);
  }
  if (Array.isArray(value)) return value.map(item => clean(item, depth + 1));
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [
      key, /authorization|token|api[_-]?key|password|secret/i.test(key) ? "[скрыто]" : clean(item, depth + 1),
    ]));
  }
  return value;
}

export function tracer(operationId: string, emit: (event: TraceEvent) => void) {
  let sequence = 0;

  function start(stage: Stage, title: string, details: { tool?: string; args?: unknown } = {}) {
    const started = performance.now();
    const event: TraceEvent = {
      operationId, sequence: ++sequence, stage, title, at: new Date().toISOString(),
      state: "running", tool: details.tool, args: clean(details.args),
    };
    emit(event);
    return {
      ok(result?: unknown) {
        emit({ ...event, state: "ok", result: clean(result), durationMs: Math.round(performance.now() - started) });
      },
      error(message: string) {
        emit({ ...event, state: "error", error: String(clean(message)), durationMs: Math.round(performance.now() - started) });
      },
    };
  }

  return {
    start,
    mark(stage: Stage, title: string, result?: unknown) {
      start(stage, title).ok(result);
    },
  };
}

export type Trace = ReturnType<typeof tracer>;
