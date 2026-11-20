import { mcpConfig, WorkspaceError } from "../lib/config";
import { withMcp } from "../lib/mcp";
import { tracer } from "../lib/trace";
import { formatDemoWorkbook, readCells, readTickets, restoreDemoRows, sheet, usedLastRow } from "../lib/workbook";

async function main() {
  const action = process.argv[2];
  if (!["tools", "init", "reset"].includes(action)) throw new Error("Используйте tools, init или reset.");
  if (action === "reset" && !process.argv.includes("--yes")) {
    throw new Error("Сброс удаляет добавленные строки. Остановите приложение и запустите npm run excel:reset -- --yes.");
  }

  const config = mcpConfig();
  if (action !== "tools" && (config.mode !== "stdio" || config.file !== "requests.xlsx")) {
    throw new Error("Инициализация и сброс предназначены для локальной демонстрационной книги requests.xlsx.");
  }

  await withMcp(config, tracer(crypto.randomUUID(), event => {
    if (event.state !== "running") console.log(event.stage, event.tool || "", event.state, event.error || "");
  }), async mcp => {
    if (action === "tools") {
      console.log(JSON.stringify([...mcp.catalog.values()].map(({ name, description, inputSchema }) => ({
        name, description, inputSchema,
      })), null, 2));
      return;
    }

    let bookExists = true;
    try {
      await mcp.call("describe_workbook", {});
    } catch (error) {
      if (!(error instanceof WorkspaceError) || action === "reset") throw error;
      await mcp.call("create_workbook", { sheets: [sheet], overwrite: false });
      bookExists = false;
    }

    if (action === "init" && bookExists) {
      const last = await usedLastRow(mcp);
      const first = await readCells(mcp, "A1:D1");
      if (last !== 1 || first.flat().some(cell => cell !== null && cell !== "")) {
        console.log(JSON.stringify(await readTickets(mcp), null, 2));
        return;
      }
    }

    const verified = await restoreDemoRows(mcp, action === "reset");
    if (!bookExists) await formatDemoWorkbook(mcp);
    console.log(JSON.stringify(verified, null, 2));
  });
}

main().catch(error => {
  console.error(error instanceof Error ? error.message : "Команда не завершилась.");
  process.exitCode = 1;
});
