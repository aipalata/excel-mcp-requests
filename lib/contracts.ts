export const statuses = ["Новый", "В работе", "Готово"] as const;
export type Status = (typeof statuses)[number];

export type Ticket = {
  id: number;
  client: string;
  problem: string;
  status: Status;
};

export type Plan = {
  id: string;
  ticket: Ticket;
  status: Status;
  expiresAt: number;
};

export type Stage =
  | "request" | "initialize" | "tools/list" | "model.request"
  | "model.choice" | "tools/call" | "model.result" | "confirm" | "response";

export type TraceEvent = {
  operationId: string;
  sequence: number;
  stage: Stage;
  title: string;
  at: string;
  state: "running" | "ok" | "error";
  tool?: string;
  args?: unknown;
  result?: unknown;
  error?: string;
  durationMs?: number;
};

export type Outcome = {
  message?: string;
  rows?: Ticket[];
  readAt?: string;
  matches?: Ticket[];
  plan?: Plan;
  error?: string;
  workbookError?: string;
};

export type WireEvent =
  | { type: "trace"; event: TraceEvent }
  | { type: "result"; result: Outcome }
  | { type: "config"; openAI: boolean; file: string; mode: string };

export type WorkspaceAction =
  | { kind: "refresh" }
  | { kind: "chat"; message: string; requestId: string }
  | { kind: "confirm" | "cancel"; planId: string };
