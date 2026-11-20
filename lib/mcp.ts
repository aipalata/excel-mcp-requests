import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { CallToolResultSchema, type Tool } from "@modelcontextprotocol/sdk/types.js";
import { WorkspaceError, type McpConfig } from "./config";
import type { Trace } from "./trace";

const allowed = ["describe_workbook", "describe_sheet", "read_range", "write_range", "create_workbook", "format_range", "set_sheet_layout"] as const;
type ToolName = (typeof allowed)[number];

const titles: Record<ToolName, string> = {
  describe_workbook: "Получаем сведения о книге",
  describe_sheet: "Проверяем границы листа",
  read_range: "Читаем ячейки Excel",
  write_range: "Записываем значения в Excel",
  create_workbook: "Создаём книгу Excel",
  format_range: "Оформляем диапазон",
  set_sheet_layout: "Настраиваем лист",
};

function createTransport(config: McpConfig) {
  if (config.mode === "http") {
    return new StreamableHTTPClientTransport(config.url, {
      requestInit: { headers: config.token ? { Authorization: `Bearer ${config.token}` } : {} },
    });
  }
  return new StdioClientTransport({
    command: config.command,
    args: ["stdio", "--allow-dir", config.directory, "--tools", "workbooks,data,formatting,sheets",
      "--max-file-mb", "8", "--max-cells", "4004"],
    stderr: "ignore",
  });
}

export async function connectMcp(config: McpConfig, trace: Trace, signal?: AbortSignal) {
  const client = new Client({ name: "excel-requests", version: "1.0.0" });
  const transport = createTransport(config);
  const options = { timeout: 8_000, signal };

  async function close() {
    if (transport instanceof StreamableHTTPClientTransport && transport.sessionId) {
      let timer: ReturnType<typeof setTimeout> | undefined;
      try {
        await Promise.race([
          transport.terminateSession(),
          new Promise<void>(resolve => { timer = setTimeout(resolve, 1_000); }),
        ]);
      } catch {
        // Some Streamable HTTP servers don't support explicit session termination.
      } finally {
        clearTimeout(timer);
      }
    }
    await client.close();
  }

  const init = trace.start("initialize", "Подключаемся к Excel MCP", { args: { transport: config.mode } });

  try {
    await client.connect(transport, options);
    init.ok(client.getServerVersion());
  } catch {
    init.error("Соединение с MCP не установлено.");
    await close().catch(() => undefined);
    throw new WorkspaceError("Excel MCP недоступен. Проверьте установку сервера и настройки транспорта.");
  }

  const catalog = new Map<string, Tool>();
  const listing = trace.start("tools/list", "Получаем каталог инструментов");
  try {
    let cursor: string | undefined;
    let pages = 0;
    do {
      const page = await client.listTools(cursor ? { cursor } : undefined, options);
      for (const tool of page.tools) catalog.set(tool.name, tool);
      cursor = page.nextCursor;
      pages++;
    } while (cursor && pages < 10);

    if (cursor || !catalog.has("read_range")) throw new Error("Incomplete catalog");
    listing.ok({
      available: [...catalog.keys()],
      allowed: allowed.filter(name => catalog.has(name)),
    });
  } catch {
    listing.error("Каталог MCP недоступен или несовместим.");
    await close().catch(() => undefined);
    throw new WorkspaceError("MCP подключён, но не вернул совместимый каталог инструментов.");
  }

  return {
    file: config.file,
    catalog,
    close,
    async call(tool: ToolName, args: Record<string, unknown>, title = titles[tool]): Promise<unknown> {
      if (!allowed.includes(tool) || !catalog.has(tool)) {
        throw new WorkspaceError(`Excel MCP не предоставил инструмент ${tool}.`);
      }
      const payload = { ...args, path: config.file };
      const step = trace.start("tools/call", title, { tool, args: payload });
      try {
        const response = CallToolResultSchema.parse(
          await client.callTool({ name: tool, arguments: payload }, CallToolResultSchema, options),
        );
        if (response.isError) {
          step.error(response.content?.map(item => item.type === "text" ? item.text : "").join(" ") ||
            "Сервер отклонил вызов.");
          throw new WorkspaceError(`Excel MCP отклонил вызов ${tool}. Подробности в трассировке.`);
        }
        const content = response.content?.find(item => item.type === "text");
        const result = response.structuredContent ?? (content?.type === "text" ? JSON.parse(content.text) : undefined);
        step.ok(result);
        return result;
      } catch (error) {
        if (error instanceof WorkspaceError) throw error;
        step.error(error instanceof Error ? error.message : "Ответ MCP не получен.");
        throw new WorkspaceError(`Вызов ${tool} не завершился. Проверьте трассировку и обновите таблицу.`);
      }
    },
  };
}

export type ExcelMcp = Awaited<ReturnType<typeof connectMcp>>;

export async function withMcp<T>(config: McpConfig, trace: Trace, task: (mcp: ExcelMcp) => Promise<T>, signal?: AbortSignal) {
  const mcp = await connectMcp(config, trace, signal);
  try {
    return await task(mcp);
  } finally {
    await mcp.close().catch(() => undefined);
  }
}
