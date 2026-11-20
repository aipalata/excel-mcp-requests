import { z } from "zod";
import { statuses, type Ticket } from "./contracts";
import { WorkbookError } from "./config";
import type { ExcelMcp } from "./mcp";

export const sheet = "Заявки";
export const headings = ["ID", "Клиент", "Проблема", "Статус"];
export const seed = [
  headings,
  [101, "Иван Петров", "Ноутбук не включается", "Новый"],
  [102, "Мария Соколова", "Принтер не печатает", "В работе"],
  [103, "Олег Иванов", "Мерцает монитор", "Готово"],
];

const limit = 1000;
const endCell = /^(?:[A-D][1-9]\d*:)?[A-D]([1-9]\d*)$/;
const readResult = z.object({ range: z.string(), values: z.array(z.array(z.unknown())), next_range: z.string().nullable().optional() });
const rawTicket = z.tuple([
  z.union([z.number(), z.string().regex(/^\d+$/)]).pipe(z.coerce.number<string | number>().int().positive().max(999_999_999)),
  z.string().min(1),
  z.string().min(1),
  z.enum(statuses),
]);

export type LocatedTicket = Ticket & { row: number };

export async function usedLastRow(mcp: ExcelMcp) {
  const result = z.object({ used_range: z.string() }).safeParse(await mcp.call("describe_sheet", { sheet }));
  const last = result.success ? Number(endCell.exec(result.data.used_range)?.[1]) : NaN;
  if (!Number.isInteger(last) || last > limit + 1) {
    throw new WorkbookError("Лист должен содержать столбцы A:D и до 1000 заявок.");
  }
  return last;
}

export async function readCells(mcp: ExcelMcp, range: string, title?: string) {
  const result = readResult.safeParse(await mcp.call("read_range", { sheet, range, max_cells: 4004 }, title));
  if (!result.success || result.data.range !== range || result.data.next_range) {
    throw new WorkbookError("MCP вернул неполный диапазон. Обновите таблицу.");
  }
  return result.data.values;
}

export async function readTickets(mcp: ExcelMcp, title?: string): Promise<LocatedTicket[]> {
  const last = await usedLastRow(mcp);
  const rows = await readCells(mcp, `A1:D${last}`, title);
  if (!rows[0] || headings.some((heading, column) => rows[0][column] !== heading)) {
    throw new WorkbookError("Ожидаются заголовки ID, Клиент, Проблема, Статус на листе «Заявки».");
  }

  const tickets: LocatedTicket[] = [];
  const ids = new Set<number>();
  for (const [index, values] of rows.slice(1).entries()) {
    if (values.every(value => value === null || value === "")) continue;
    const parsed = rawTicket.safeParse(values);
    if (!parsed.success || ids.has(parsed.data[0])) {
      throw new WorkbookError(`В строке ${index + 2} неверные данные или повторяющийся ID.`);
    }
    const [id, client, problem, status] = parsed.data;
    ids.add(id);
    tickets.push({ row: index + 2, id, client, problem, status });
  }
  return tickets;
}

export function publicTicket({ id, client, problem, status }: Ticket): Ticket {
  return { id, client, problem, status };
}

export function sameTicket(a: Ticket, b: Ticket) {
  return a.id === b.id && a.client === b.client && a.problem === b.problem && a.status === b.status;
}

export async function writeCells(
  mcp: ExcelMcp, at: string, rows: (string | number | null)[][], title?: string, operationId?: string,
) {
  return mcp.call("write_range", { sheet, at, rows }, title, operationId);
}

export async function formatDemoWorkbook(mcp: ExcelMcp) {
  await mcp.call("format_range", {
    sheet, range: "A1:D1", style: { bold: true, font_color: "FFFFFF", fill_color: "08796D" },
  });
  await mcp.call("set_sheet_layout", {
    sheet, layout: { column_widths_chars: { A: 9, B: 24, C: 43, D: 15 }, freeze_panes: "A2", tab_color: "08796D" },
  });
}

export async function restoreDemoRows(mcp: ExcelMcp, reset: boolean) {
  if (reset) await readTickets(mcp);
  const count = reset ? await usedLastRow(mcp) : seed.length;
  const blanks = Array.from({ length: Math.max(0, count - seed.length) }, () => [null, null, null, null]);
  await writeCells(mcp, "A1", [...seed, ...blanks]);

  const verified = await readTickets(mcp, "Проверяем исходные заявки");
  const actual = verified.map(ticket => [ticket.id, ticket.client, ticket.problem, ticket.status]);
  if (JSON.stringify(actual) !== JSON.stringify(seed.slice(1))) {
    throw new WorkbookError("Контрольное чтение не подтвердило исходные данные.");
  }
  return verified;
}
