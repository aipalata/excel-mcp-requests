import { createHash, randomUUID } from "node:crypto";
import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { z } from "zod";
import { statuses, type Outcome, type Plan, type Status, type Ticket } from "./contracts";
import { WorkbookError, WorkspaceError } from "./config";
import type { ExcelMcp } from "./mcp";
import { publicTicket, readTickets, sameTicket, writeCells, type LocatedTicket } from "./workbook";

const literalText = (max: number) => z.string().trim().min(1).max(max).refine(value =>
  !/^[=+\-@＝＋－＠]|[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(value.normalize("NFKC")) &&
  !/^\d{4}-\d{2}-\d{2}(?:$|T)/.test(value));

export const newTicketSchema = z.strictObject({ client: literalText(100), problem: literalText(450) });
export const lookupSchema = z.strictObject({
  id: z.number().int().positive().max(999_999_999).nullable(),
  client: z.string().trim().min(1).max(100).nullable(),
  status: z.enum(statuses).nullable(),
});
export const statusPlanSchema = lookupSchema.extend({ status: z.enum(statuses) })
  .refine(values => Number(values.id !== null) + Number(values.client !== null) === 1);

type SessionPlan = Plan & { sessionId: string };
const lockKey = "__excelRequestsState" as const;
const globalState = globalThis as typeof globalThis & {
  [lockKey]?: { queue: Promise<unknown>; plans: Map<string, SessionPlan>; turns: Set<string> };
};
const state = globalState[lockKey] ??= { queue: Promise.resolve(), plans: new Map(), turns: new Set() };

export function serialized<T>(task: () => Promise<T>, signal?: AbortSignal): Promise<T> {
  const result = state.queue.then(() => {
    signal?.throwIfAborted();
    return task();
  });
  state.queue = result.catch(() => undefined);
  return result;
}

export async function withTurn<T>(sessionId: string, task: () => Promise<T>) {
  if (state.turns.has(sessionId)) throw new WorkspaceError("Предыдущая команда ещё выполняется.");
  state.turns.add(sessionId);
  try {
    return await task();
  } finally {
    state.turns.delete(sessionId);
  }
}

export function snapshot(rows: LocatedTicket[]) {
  return { rows: rows.map(publicTicket), readAt: new Date().toISOString() };
}

const normalName = (name: string) => name.toLocaleLowerCase("ru").replaceAll("ё", "е").trim().replace(/\s+/g, " ");

export function selectTickets(rows: LocatedTicket[], query: z.infer<typeof lookupSchema>) {
  return rows.filter(ticket =>
    (query.id === null || ticket.id === query.id) &&
    (query.client === null || ` ${normalName(ticket.client)} `.includes(` ${normalName(query.client)} `)) &&
    (query.status === null || ticket.status === query.status));
}

const journalRecord = z.strictObject({
  row: z.number().int().min(2).max(1001),
  id: z.number().int().positive().max(999_999_999),
  client: z.string(),
  problem: z.string(),
  committed: z.boolean(),
});

async function journalEntry(directory: string, workbook: string, sessionId: string, requestId: string) {
  const name = createHash("sha256").update(JSON.stringify([workbook, sessionId, requestId])).digest("hex");
  const path = join(directory, `${name}.json`);
  let record: z.infer<typeof journalRecord> | undefined;
  try {
    record = journalRecord.parse(JSON.parse(await readFile(path, "utf8")));
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
  }

  return {
    record,
    async write(next: z.infer<typeof journalRecord>) {
      await mkdir(directory, { recursive: true, mode: 0o700 });
      const temporary = `${path}.${randomUUID()}.tmp`;
      await writeFile(temporary, JSON.stringify(next), { flag: "wx", mode: 0o600 });
      await rename(temporary, path);
    },
  };
}

export async function addTicket(
  mcp: ExcelMcp, input: z.infer<typeof newTicketSchema>, context: { sessionId: string; requestId: string; journalDir: string },
) {
  const fields = newTicketSchema.parse(input);
  const rows = await readTickets(mcp, "Читаем заявки и определяем свободный ID");
  const journal = await journalEntry(context.journalDir, mcp.file, context.sessionId, context.requestId);

  let record = journal.record;
  if (record) {
    if (record.client !== fields.client || record.problem !== fields.problem) {
      throw new WorkspaceError("Эта операция уже использована с другими данными.");
    }
    const existing = rows.find(ticket => ticket.id === record?.id);
    if (existing?.client === record.client && existing.problem === record.problem) {
      if (!record.committed) await journal.write({ ...record, committed: true });
      return { ...snapshot(rows), created: publicTicket(existing), replayed: true };
    }
    if (record.committed || existing || rows.some(ticket => ticket.row === record?.row)) {
      throw new WorkspaceError("Состояние Excel изменилось после прошлой попытки. Проверьте таблицу перед новой отправкой.");
    }
  } else {
    const id = Math.max(100, ...rows.map(row => row.id)) + 1;
    const row = Math.max(1, ...rows.map(row => row.row)) + 1;
    if (row > 1001 || id > 999_999_999) throw new WorkspaceError("Достигнут лимит заявок или идентификаторов.");
    record = { ...fields, id, row, committed: false };
    await journal.write(record);
  }

  const { id, row, client, problem } = record;
  const created: Ticket = { id, client, problem, status: "Новый" };
  await writeCells(mcp, `A${row}`, [[id, client, problem, created.status]], `Добавляем заявку №${id}`);
  const verified = await readTickets(mcp, `Перечитываем заявку №${id}`);
  if (!verified.some(ticket => ticket.row === row && sameTicket(ticket, created))) {
    throw new WorkbookError("Контрольное чтение не подтвердило добавление. Обновите таблицу.");
  }
  await journal.write({ ...record, committed: true });
  return { ...snapshot(verified), created, replayed: false };
}

export async function prepareStatus(
  mcp: ExcelMcp, sessionId: string, input: z.infer<typeof statusPlanSchema>,
): Promise<Outcome> {
  const query = statusPlanSchema.parse(input);
  const rows = await readTickets(mcp, "Ищем заявку перед изменением статуса");
  const target = selectTickets(rows, { ...query, status: null });
  const base = snapshot(rows);
  if (target.length === 0) {
    return { ...base, message: query.id ? `Заявка №${query.id} не найдена.` : "По этому имени заявок нет." };
  }
  if (target.length > 1) {
    return { ...base, matches: target.map(publicTicket), message: "Найдено несколько заявок. Уточните номер заявки." };
  }

  const ticket = publicTicket(target[0]);
  if (ticket.status === query.status) return { ...base, message: `У заявки №${ticket.id} уже статус «${ticket.status}».` };
  const plan: Plan = { id: randomUUID(), ticket, status: query.status, expiresAt: Date.now() + 300_000 };
  for (const [id, old] of state.plans) {
    if (old.sessionId === sessionId || old.expiresAt <= Date.now()) state.plans.delete(id);
  }
  state.plans.set(plan.id, { ...plan, sessionId });
  return { ...base, plan, message: `Подтвердите изменение статуса заявки №${ticket.id}.` };
}

export function consumePlan(planId: string, sessionId: string): Plan {
  const plan = state.plans.get(planId);
  if (!plan || plan.sessionId !== sessionId || plan.expiresAt <= Date.now()) {
    throw new WorkspaceError("План недоступен. Запросите изменение статуса ещё раз.");
  }
  state.plans.delete(planId);
  return plan;
}

export async function applyStatus(mcp: ExcelMcp, plan: Plan): Promise<Outcome> {
  if (plan.expiresAt <= Date.now()) throw new WorkspaceError("Срок подтверждения истёк.");
  const before = await readTickets(mcp, "Сверяем заявку с подтверждённым планом");
  const current = before.find(ticket => ticket.id === plan.ticket.id);
  if (!current || !sameTicket(current, plan.ticket)) {
    return { ...snapshot(before), error: "Заявка изменилась после согласования. Запросите новый план." };
  }
  if (plan.expiresAt <= Date.now()) throw new WorkspaceError("Срок подтверждения истёк.");
  await writeCells(mcp, `D${current.row}`, [[plan.status]], `Меняем статус заявки №${current.id}`);
  const after = await readTickets(mcp, "Проверяем статус и остальные поля");
  const expected = before.map(ticket => ticket.id === current.id ? { ...ticket, status: plan.status } : ticket);
  if (after.length !== expected.length || after.some((ticket, index) =>
    ticket.row !== expected[index].row || !sameTicket(ticket, expected[index]))) {
    throw new WorkbookError("Контрольное чтение выявило расхождение. Обновите таблицу.");
  }
  return { ...snapshot(after), message: `Статус заявки №${current.id} изменён на «${plan.status}».` };
}

export function defaultJournalDir() {
  return join(process.cwd(), ".runtime", "operations");
}
