import { join } from "node:path";

export class WorkspaceError extends Error {}
export class WorkbookError extends WorkspaceError {}

export function workbookName() {
  const file = process.env.EXCEL_FILE_PATH?.trim() || "requests.xlsx";
  if (!/^[\p{L}\p{N} _-]+\.xlsx$/u.test(file)) {
    throw new WorkspaceError("Укажите в EXCEL_FILE_PATH имя файла .xlsx без пути.");
  }
  return file;
}

export function mcpConfig() {
  const mode = process.env.MCP_MODE?.trim() || "stdio";
  const file = workbookName();

  if (mode === "stdio") {
    if (process.env.VERCEL) throw new WorkspaceError("На Vercel требуется транспорт MCP по HTTPS.");
    return {
      mode, file,
      directory: join(process.cwd(), "data"),
      command: process.env.MCP_STDIO_COMMAND || join(process.cwd(), ".venv", "bin", "excel-mcp-server"),
    } as const;
  }
  if (mode === "http") {
    const token = process.env.MCP_HTTP_TOKEN?.trim();
    let url: URL;
    try {
      url = new URL(process.env.MCP_HTTP_URL || "");
    } catch {
      throw new WorkspaceError("Укажите корректный MCP_HTTP_URL.");
    }
    const loopback = ["localhost", "127.0.0.1", "[::1]"].includes(url.hostname);
    if ((url.protocol !== "https:" && !(loopback && url.protocol === "http:")) ||
        url.username || url.password || url.search || url.hash || (!loopback && !token)) {
      throw new WorkspaceError("Внешний MCP-сервер требует HTTPS и отдельный Bearer-токен.");
    }
    return { mode, file, url, token } as const;
  }
  throw new WorkspaceError("Укажите MCP_MODE=stdio или MCP_MODE=http.");
}

export type McpConfig = ReturnType<typeof mcpConfig>;

export function openAIConfig() {
  return {
    apiKey: process.env.OPENAI_API_KEY?.trim() || "",
    model: process.env.OPENAI_MODEL?.trim() || "gpt-5-mini",
  };
}
