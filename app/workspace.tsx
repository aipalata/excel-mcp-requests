"use client";

import {
  ArrowRight, ArrowUp, Check, ChevronRight, CircleAlert, FileSpreadsheet, FolderOpen,
  ListFilter, LoaderCircle, MessageSquare, Plus, RefreshCw, ShieldCheck, Sparkles, SquarePen,
} from "lucide-react";
import { useCallback, useEffect, useRef, useState, type FormEvent } from "react";
import { statuses, type Outcome, type Plan, type Status, type Ticket, type TraceEvent, type WireEvent, type WorkspaceAction } from "@/lib/contracts";
import { Hero } from "./hero";
import { TracePanel } from "./trace-panel";

const examples = [
  { Icon: Plus, tone: "teal", name: "Добавить заявку", hint: "Новая запись в Excel", prompt: "Добавь заявку от Анны Смирновой: ноутбук сильно тормозит" },
  { Icon: SquarePen, tone: "violet", name: "Изменить статус", hint: "С подтверждением", prompt: "Измени статус заявки Анны Смирновой на «В работе»" },
  { Icon: ListFilter, tone: "blue", name: "Найти заявки", hint: "Список по статусу", prompt: "Покажи все заявки со статусом «В работе»" },
] as const;

const statusTone: Record<Status, string> = { "Новый": "blue", "В работе": "amber", "Готово": "teal" };
type Frame = { id: string; role: "user" | "assistant" | "error"; message: string; matches?: Ticket[] };

function Badge({ status }: { status: Status }) {
  return <span className={`status-badge ${statusTone[status]}`}><span />{status}</span>;
}

function PlanCard({ plan, disabled, decide }: { plan: Plan; disabled: boolean; decide: (kind: "confirm" | "cancel") => void }) {
  return (
    <div className="plan">
      <div className="plan-heading"><ShieldCheck size={18} /><strong>Заявка №{plan.ticket.id} · {plan.ticket.client}</strong></div>
      <div className="status-change"><Badge status={plan.ticket.status} /><ArrowRight size={15} /><Badge status={plan.status} /></div>
      <small>Подтвердите до {new Date(plan.expiresAt).toLocaleTimeString("ru-RU", { hour: "2-digit", minute: "2-digit" })}.</small>
      <div className="plan-actions">
        <button type="button" className="button primary" disabled={disabled} onClick={() => decide("confirm")}>Подтвердить изменение</button>
        <button type="button" className="button secondary" disabled={disabled} onClick={() => decide("cancel")}>Отменить</button>
      </div>
    </div>
  );
}

async function readStream(action: WorkspaceAction, receive: (event: WireEvent) => void) {
  const response = await fetch("/api/workspace", {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(action),
  });
  if (!response.ok || !response.body) {
    let error = "Сервер не принял запрос.";
    try { error = (await response.json()).error || error; } catch {}
    throw new Error(error);
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let remainder = "";
  let finished = false;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      const lines = (remainder + decoder.decode(value, { stream: true })).split("\n");
      remainder = lines.pop() || "";
      for (const line of lines.filter(Boolean)) {
        const event = JSON.parse(line) as WireEvent;
        receive(event);
        if (event.type === "result") finished = true;
      }
    }
  } finally {
    reader.releaseLock();
  }
  if (!finished) throw new Error("Соединение прервалось. Повторите запрос или обновите таблицу.");
}

const prettyTime = (time: string) => new Date(time).toLocaleTimeString("ru-RU");
const ticketCount = (count: number) => {
  const form = new Intl.PluralRules("ru").select(count);
  return `${count} ${form === "one" ? "заявка" : form === "few" ? "заявки" : "заявок"}`;
};

export function Workspace() {
  const [file, setFile] = useState("requests.xlsx");
  const [transport, setTransport] = useState("stdio");
  const [openAI, setOpenAI] = useState<boolean | null>(null);
  const [readOnly, setReadOnly] = useState(false);
  const [mcp, setMcp] = useState<boolean | null>(null);
  const [rows, setRows] = useState<Ticket[] | null>(null);
  const [readAt, setReadAt] = useState("");
  const [workbookError, setWorkbookError] = useState("");
  const [messages, setMessages] = useState<Frame[]>([]);
  const [events, setEvents] = useState<TraceEvent[]>([]);
  const [filter, setFilter] = useState<"Все" | Status>("Все");
  const [matchIds, setMatchIds] = useState<Set<number> | null>(null);
  const [plan, setPlan] = useState<Plan | null>(null);
  const [busy, setBusy] = useState<WorkspaceAction["kind"] | null>(null);
  const [draft, setDraft] = useState("");
  const running = useRef(false);
  const initial = useRef(false);
  const retry = useRef<{ message: string; id: string } | null>(null);
  const input = useRef<HTMLTextAreaElement>(null);
  const transcript = useRef<HTMLDivElement>(null);

  const run = useCallback(async (action: WorkspaceAction) => {
    if (running.current) return;
    running.current = true;
    setBusy(action.kind);
    let readFailed = false;
    let sawConfig = false;

    function finish(result: Outcome) {
      if (result.rows) {
        setRows(result.rows);
        setReadAt(result.readAt || "");
        setWorkbookError("");
      } else if (result.workbookError || readFailed) {
        setRows(null);
        setWorkbookError(result.workbookError || "Контрольное чтение Excel не завершилось.");
      }

      if (action.kind === "chat" && (result.message || result.rows)) retry.current = null;
      if (result.matches && action.kind === "chat" && !result.plan) {
        setMatchIds(new Set(result.matches.map(ticket => ticket.id)));
        setFilter("Все");
      } else if (result.rows) {
        setMatchIds(null);
      }
      if (result.plan) setPlan(result.plan);
      if (action.kind === "confirm" || action.kind === "cancel") setPlan(null);
      if (result.message) {
        const frame: Frame = { id: crypto.randomUUID(), role: "assistant", message: result.message, matches: result.matches };
        setMessages(items => [...items, frame]);
      }
      if (result.error) {
        const frame: Frame = { id: crypto.randomUUID(), role: "error", message: result.error! };
        setMessages(items => [...items, frame]);
        if (action.kind === "chat" && !result.message) setDraft(action.message);
      }
    }

    try {
      await readStream(action, event => {
        if (event.type === "config") {
          sawConfig = true;
          setOpenAI(event.openAI);
          setReadOnly(Boolean(event.readOnly));
          setFile(event.file);
          setTransport(event.mode);
        } else if (event.type === "trace") {
          const next = event.event;
          if (next.stage === "initialize" && next.state !== "running") setMcp(next.state === "ok");
          if (next.stage === "tools/list" && next.state === "error") setMcp(false);
          if (next.state === "error" && (next.stage === "initialize" || next.stage === "tools/list" ||
             next.tool === "read_range" || next.tool === "describe_sheet" || next.tool === "write_range")) readFailed = true;
          setEvents(current => {
            const previous = current.findIndex(item => item.operationId === next.operationId && item.sequence === next.sequence);
            if (previous < 0) return [...current, next].slice(-180);
            return current.map((item, index) => index === previous ? next : item);
          });
        } else {
          finish(event.result);
        }
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Не удалось получить ответ приложения.";
      if (action.kind === "refresh" && !sawConfig) setMcp(false);
      finish(action.kind === "refresh" ? { workbookError: message } : { error: message });
    } finally {
      running.current = false;
      setBusy(null);
    }
  }, []);

  useEffect(() => {
    if (initial.current) return;
    initial.current = true;
    void run({ kind: "refresh" });
  }, [run]);

  useEffect(() => {
    if (transcript.current) transcript.current.scrollTop = transcript.current.scrollHeight;
  }, [messages, plan, busy]);

  function submit(event: FormEvent) {
    event.preventDefault();
    const text = draft.trim();
    if (!text || running.current || !openAI || plan) return;
    const id = retry.current?.message === text ? retry.current.id : crypto.randomUUID();
    retry.current = { message: text, id };
    setMessages(items => [...items, { id: crypto.randomUUID(), role: "user", message: text }]);
    setDraft("");
    void run({ kind: "chat", message: text, requestId: id });
  }

  const visible = rows?.filter(ticket => (filter === "Все" || ticket.status === filter) && (!matchIds || matchIds.has(ticket.id))) ?? [];
  const connection = mcp === null ? "Подключение к MCP" : mcp ? "MCP подключён" : "MCP недоступен";

  return (
    <>
      <header className="site-header">
        <div className="site-header-inner shell">
          <a href="#" className="brand"><span className="brand-mark"><Sparkles size={19} strokeWidth={1.7} /></span><span>AI-агент <strong>«Заявки»</strong></span></a>
          <nav aria-label="Навигация"><a href="#workspace">Рабочая область</a><a href="#under-the-hood">За кадром</a></nav>
          <div className={`health ${mcp ? "healthy" : mcp === false ? "unavailable" : ""}`} role="status"><span className="dot" />{connection}</div>
        </div>
      </header>

      <main className="shell">
        <Hero />
        <section id="workspace" className="workspace" aria-label="Рабочая область">
          <aside className="panel sidebar" aria-labelledby="commands-title">
            <div className="panel-header"><div><h2 id="commands-title">Команды</h2><p>Примеры для чата</p></div></div>
            <div className="example-list">
              {examples.map(({ name, prompt, hint, tone, Icon }) => (
                <button key={name} type="button" className="example" onClick={() => { setDraft(prompt); input.current?.focus(); }}>
                  <span className={`role-icon ${tone}`}><Icon size={17} /></span>
                  <span><strong>{name}</strong><small>{hint}</small></span>
                  <ChevronRight size={14} className="example-arrow" />
                </button>
              ))}
            </div>
            <div className="connections">
              <div><span className={`signal ${openAI ? "good" : ""}`} /> <span>OpenAI</span><span>{readOnly ? "Недоступен" : openAI === null ? "Проверка" : openAI ? "Ключ задан" : "Не настроен"}</span></div>
              <div><span className={`signal ${mcp ? "good" : ""}`} /> <span>MCP</span><code>{transport === "http" ? "HTTP" : "stdio"}</code></div>
              <div><FileSpreadsheet size={13} className="excel-icon" /> <span>Excel</span><code title={file}>{file}</code></div>
            </div>
          </aside>

          <article className="panel chat" aria-labelledby="chat-title">
            <div className="panel-header chat-header">
              <div className="panel-identity">
                <span className="role-icon teal"><MessageSquare size={18} /></span>
                <div><h2 id="chat-title">Чат с агентом</h2><p>Команды на обычном языке</p></div>
              </div>
              <span className="chip"><Sparkles size={12} /> OpenAI</span>
            </div>

            <div className="transcript" ref={transcript} role="log" aria-label="История диалога" aria-live="polite">
              {!messages.length && (
                <div className="greeting">
                  <span className="role-icon teal"><Sparkles size={17} /></span>
                  <div><strong>Чем помочь с заявками?</strong><p>Выберите команду слева или напишите свою.</p></div>
                </div>
              )}
              {messages.map(frame => (
                <div key={frame.id} className={`message ${frame.role}`}>
                  {frame.role === "error" && <CircleAlert size={15} />}
                  <div>{frame.message}
                    {frame.matches && !!frame.matches.length && <ul>
                      {frame.matches.slice(0, 6).map(ticket => <li key={ticket.id}>№{ticket.id} · {ticket.client}</li>)}
                      {frame.matches.length > 6 && <li>Ещё {frame.matches.length - 6} в таблице</li>}
                    </ul>}
                  </div>
                </div>
              ))}
              {plan && <PlanCard plan={plan} disabled={Boolean(busy)} decide={kind => void run({ kind, planId: plan.id })} />}
              {(busy === "chat" || busy === "confirm") && <div className="working"><LoaderCircle className="spinner" size={15} /> {busy === "chat" ? "Агент обрабатывает команду…" : "Excel сохраняет изменение…"}</div>}
            </div>

            {readOnly ? <div className="key-state"><CircleAlert size={15} /><span><strong>Книга доступна для чтения.</strong> Чат в облаке пока отключён.</span></div>
              : openAI === false && <div className="key-state"><CircleAlert size={15} /><span><strong>OpenAI не настроен.</strong> Для отправки команды нужен API-ключ.</span></div>}
            <form className="composer" onSubmit={submit}>
              <label htmlFor="command">Сообщение агенту</label>
              <textarea id="command" ref={input} placeholder="Напишите команду…" maxLength={1600} rows={2}
                value={draft} onChange={event => setDraft(event.target.value)} onKeyDown={event => {
                  if (event.key === "Enter" && !event.shiftKey && !event.nativeEvent.isComposing) {
                    event.preventDefault();
                    event.currentTarget.form?.requestSubmit();
                  }
                }} />
              <button className="send-button" type="submit" title="Отправить команду" aria-label="Отправить команду"
                disabled={!openAI || Boolean(busy) || !draft.trim() || Boolean(plan)}><ArrowUp size={19} /></button>
            </form>
          </article>

          <article className="panel workbook" aria-labelledby="workbook-title">
            <div className="panel-header workbook-header">
              <div className="panel-identity">
                <span className="role-icon amber"><FileSpreadsheet size={19} /></span>
                <div><h2 id="workbook-title">Заявки в Excel</h2><p><code>{file}</code><span className="heading-delimiter">·</span> Заявки</p></div>
              </div>
              <button type="button" className="button secondary refresh-button" title="Обновить данные"
                aria-label="Обновить данные" disabled={Boolean(busy)} onClick={() => void run({ kind: "refresh" })}>
                <RefreshCw size={15} className={busy === "refresh" ? "spinner" : ""} /> <span>Обновить</span>
              </button>
            </div>

            <div className="filters" aria-label="Фильтр по статусу">
              {(["Все", ...statuses] as const).map(status => (
                <button type="button" key={status} aria-pressed={filter === status && !matchIds}
                  onClick={() => { setFilter(status); setMatchIds(null); }}>{status}</button>
              ))}
              {matchIds && <span className="result-filter">Результат агента</span>}
            </div>

            {rows && !workbookError ? (
              <div className="table-scroll" tabIndex={0} aria-label="Таблица заявок">
                <table>
                  <thead><tr><th scope="col">ID</th><th scope="col">Клиент</th><th scope="col">Проблема</th><th scope="col">Статус</th></tr></thead>
                  <tbody>
                    {visible.map(ticket => <tr key={ticket.id}>
                      <td className="ticket-id">{ticket.id}</td><td className="client-name">{ticket.client}</td>
                      <td className="problem">{ticket.problem}</td><td><Badge status={ticket.status} /></td>
                    </tr>)}
                    {!visible.length && <tr><td className="empty-row" colSpan={4}>Подходящих заявок нет.</td></tr>}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className={`workbook-empty ${workbookError ? "workbook-error" : ""}`} role="status">
                {workbookError ? <CircleAlert size={23} /> : busy === "refresh" ? <LoaderCircle className="spinner" size={23} /> : <FolderOpen size={23} />}
                <p>{workbookError || (busy === "refresh" ? "Читаем книгу через MCP…" : "Ожидаем данные Excel")}</p>
              </div>
            )}

            <div className="table-footer" aria-live="polite">
              {rows ? <><span><Check size={13} /> {ticketCount(visible.length)}{visible.length !== rows.length && ` из ${rows.length}`}</span><time dateTime={readAt} suppressHydrationWarning>Обновлено в {prettyTime(readAt)}</time></>
                : <span><span className="signal" /> {workbookError ? "Данные недоступны" : "Ожидаем чтение Excel"}</span>}
            </div>
          </article>
        </section>

        <TracePanel events={events} />
      </main>

      <footer className="site-footer shell"><span><Sparkles size={14} /> AI-агент «Заявки»</span><span>OpenAI · Model Context Protocol · Excel</span></footer>
    </>
  );
}
