import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { mkdir, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { mcpConfig } from "../lib/config";
import type { TraceEvent } from "../lib/contracts";
import { withMcp } from "../lib/mcp";
import { addTicket, applyStatus, consumePlan, newTicketSchema, prepareStatus, selectTickets, serialized } from "../lib/operations";
import { tracer } from "../lib/trace";
import { publicTicket, readTickets, restoreDemoRows, seed, sheet, writeCells } from "../lib/workbook";

test("Integration with Excel MCP Server", async t => {
  const directory = await mkdtemp(join(tmpdir(), "requests-mcp-"));
  const workbookDirectory = join(directory, "workbooks");
  await mkdir(workbookDirectory);
  t.after(() => rm(directory, { recursive: true, force: true }));

  const local = mcpConfig();
  assert.equal(local.mode, "stdio");
  const config = { ...local, directory: workbookDirectory };
  const trace: TraceEvent[] = [];
  const logger = tracer(randomUUID(), event => trace.push(event));
  const sessionId = randomUUID();
  const addContext = { sessionId, requestId: randomUUID(), journalDir: join(directory, "journal") };
  const anna = { client: "Анна Смирнова", problem: "Ноутбук сильно тормозит" };
  const change = { id: null, client: "Анна Смирнова", status: "В работе" } as const;

  await withMcp(config, logger, async mcp => {
    await t.test("initialize, tools/list, create and verify the three original rows", async () => {
      assert.equal(mcp.catalog.size, 26);
      for (const name of ["create_workbook", "describe_sheet", "read_range", "write_range"]) assert.ok(mcp.catalog.has(name));
      await mcp.call("create_workbook", { sheets: [sheet], overwrite: false });
      await writeCells(mcp, "A1", seed);
      const rows = await readTickets(mcp);
      assert.deepEqual(rows.map(ticket => ticket.id), [101, 102, 103]);
      assert.equal(rows[1].client, "Мария Соколова");
    });

    await t.test("append persists #104 and concurrent retries use the same row", async () => {
      const [first, second] = await Promise.all([
        serialized(() => addTicket(mcp, anna, addContext)),
        serialized(() => addTicket(mcp, anna, addContext)),
      ]);
      assert.deepEqual(first.created, { id: 104, ...anna, status: "Новый" });
      assert.equal(first.replayed, false);
      assert.equal(second.replayed, true);
      assert.equal(second.rows.length, 4);
    });

    await t.test("plan does not write, cannot be consumed in another session, and cancellation is single-use", async () => {
      const before = await readTickets(mcp);
      const { plan } = await prepareStatus(mcp, sessionId, change);
      assert.ok(plan);
      assert.equal(plan.ticket.status, "Новый");
      assert.deepEqual(await readTickets(mcp), before);
      assert.throws(() => consumePlan(plan.id, randomUUID()), /План недоступен/);
      consumePlan(plan.id, sessionId);
      assert.throws(() => consumePlan(plan.id, sessionId), /План недоступен/);
      assert.deepEqual(await readTickets(mcp), before);
    });

    await t.test("confirmed plan writes only D5 and verifies all other data", async () => {
      const before = await readTickets(mcp);
      const { plan } = await prepareStatus(mcp, sessionId, change);
      assert.ok(plan);
      const result = await applyStatus(mcp, consumePlan(plan.id, sessionId));
      const expected = before.map(ticket => publicTicket(ticket.id === 104 ? { ...ticket, status: "В работе" } : ticket));
      assert.deepEqual(result.rows, expected);
      assert.deepEqual(trace.findLast(event => event.tool === "write_range")?.args, {
        at: "D5", rows: [["В работе"]], sheet, path: "requests.xlsx",
      });
    });

    await t.test("reads two in-progress tickets and rejects #999 without writing", async () => {
      const before = await readTickets(mcp);
      const matches = selectTickets(before, { id: null, client: null, status: "В работе" });
      assert.deepEqual(matches.map(ticket => ticket.id), [102, 104]);
      const start = trace.length;
      const missing = await prepareStatus(mcp, sessionId, { id: 999, client: null, status: "Готово" });
      assert.match(missing.message || "", /999 не найдена/);
      assert.equal(missing.plan, undefined);
      assert.ok(!trace.slice(start).some(event => event.tool === "write_range"));
      assert.deepEqual(await readTickets(mcp), before);
    });

    await t.test("missing problem and formula-like text are rejected before MCP write", async () => {
      const start = trace.length;
      assert.ok(!newTicketSchema.safeParse({ client: "Дмитрий" }).success);
      await assert.rejects(() => addTicket(mcp, { client: "Дмитрий", problem: "" }, {
        ...addContext, requestId: randomUUID(),
      }));
      assert.ok(!newTicketSchema.safeParse({ client: "Дмитрий", problem: "  ＝1+2" }).success);
      assert.ok(!trace.slice(start).some(event => event.tool === "write_range"));
    });

    await t.test("rejects a stale plan after a real external change via MCP", async () => {
      const { plan } = await prepareStatus(mcp, sessionId, { id: 104, client: null, status: "Готово" });
      assert.ok(plan);
      await writeCells(mcp, "C5", [["Изменённое описание"]]);
      const start = trace.length;
      const stale = await applyStatus(mcp, consumePlan(plan.id, sessionId));
      assert.match(stale.error || "", /изменилась/);
      assert.equal(stale.rows?.find(ticket => ticket.id === 104)?.status, "В работе");
      assert.ok(!trace.slice(start).some(event => event.tool === "write_range"));
    });

    await t.test("two clients with matching names require a ticket ID", async () => {
      const other = await addTicket(mcp, { ...anna, problem: "Не заряжается" }, { ...addContext, requestId: randomUUID() });
      assert.equal(other.created.id, 105);
      const ambiguous = await prepareStatus(mcp, sessionId, { ...change, status: "Готово" });
      assert.deepEqual(ambiguous.matches?.map(ticket => ticket.id), [104, 105]);
      assert.equal(ambiguous.plan, undefined);
    });
  });

  await t.test("reopen the file and verify data persisted", async () => {
    await withMcp(config, logger, async mcp => {
      const rows = await readTickets(mcp);
      assert.equal(rows.find(ticket => ticket.id === 104)?.status, "В работе");
      assert.equal(rows.find(ticket => ticket.id === 105)?.client, "Анна Смирнова");
    });
  });

  await t.test("demo reset restores the three original tickets through MCP", async () => {
    await withMcp(config, logger, async mcp => {
      const reset = await restoreDemoRows(mcp, true);
      assert.deepEqual(reset.map(ticket => ticket.id), [101, 102, 103]);
      assert.deepEqual((await readTickets(mcp)).map(ticket => ticket.id), [101, 102, 103]);
    });
  });

  await t.test("a missing MCP command reports a connection error without data", async () => {
    await assert.rejects(() => withMcp({ ...config, command: join(directory, "missing-server") }, logger, async () => {
      assert.fail("Missing MCP server must never connect");
    }), /Excel MCP недоступен/);
  });

  await t.test("trace stores actual arguments, results, state and sequence", () => {
    const read = trace.find(event => event.stage === "tools/call" && event.tool === "read_range" && event.state === "ok");
    assert.ok(read);
    assert.deepEqual(read.args, { path: "requests.xlsx", sheet, range: "A1:D4", max_cells: 4004 });
    assert.equal((read.result as { values: unknown[] }).values.length, 4);
    assert.ok(read.operationId && read.at && read.sequence > 0);
    assert.ok(typeof read.durationMs === "number");
  });
});
