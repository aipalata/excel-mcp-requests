"use client";

import { ArrowDown, ArrowRight, Check, ChevronDown, ChevronLeft, ChevronRight, CircleAlert,
  Clock3, Code2, FileSpreadsheet, LayoutPanelTop, Network, Pause, Play, Plug, RotateCcw,
  Sparkles, UserRound } from "lucide-react";
import { useEffect, useState } from "react";
import type { Stage, TraceEvent } from "@/lib/contracts";

const actors = [
  { label: "Пользователь", tone: "blue", Icon: UserRound },
  { label: "Приложение", tone: "neutral", Icon: LayoutPanelTop },
  { label: "OpenAI", tone: "teal", Icon: Sparkles },
  { label: "MCP-клиент", tone: "violet", Icon: Plug },
  { label: "Excel MCP", tone: "violet", Icon: Network },
  { label: "Книга Excel", tone: "amber", Icon: FileSpreadsheet },
] as const;

const compactLabels = ["Вы", "Сайт", "OpenAI", "Клиент", "Сервер", "Excel"];

const explanation = [
  { from: 3, to: 4, caption: "initialize", title: "MCP-клиент подключается", text: "Приложение запускает готовый сервер и согласует с ним версию протокола." },
  { from: 3, to: 4, caption: "tools/list", title: "Клиент запрашивает каталог", text: "Сервер возвращает имена, описания и схемы параметров. Приложение отбирает разрешённые операции." },
  { from: 0, to: 1, caption: "Сообщение", title: "Пользователь задаёт вопрос", text: "Например, просит показать заявки со статусом «В работе»." },
  { from: 1, to: 2, caption: "responses.create", title: "Приложение обращается к модели", text: "В OpenAI уходит запрос с инструкциями, короткой историей и схемами доступных функций." },
  { from: 2, to: 1, caption: "function_call", title: "Модель просит вызвать инструмент", text: "Она выбирает функцию и передаёт аргументы. Приложение проверяет допустимость операции." },
  { from: 3, to: 4, caption: "tools/call", title: "MCP-клиент выполняет вызов", text: "Приложение передаёт проверенные параметры готовому серверу. Для смены статуса запись ждёт подтверждения." },
  { from: 4, to: 5, caption: ".xlsx", title: "Сервер обращается к Excel", text: "Сервер читает диапазон или записывает значения. После записи приложение выполняет контрольное чтение." },
  { from: 4, to: 3, caption: "Результат MCP", title: "Клиент получает результат", text: "В журнал попадает фактический ответ инструмента, его статус и время выполнения." },
  { from: 1, to: 2, caption: "function_call_output", title: "Приложение передаёт результат модели", text: "Модель получает данные и решает, нужен ли ещё один вызов." },
  { from: 1, to: 0, caption: "Ответ", title: "Пользователь видит результат", text: "В чате появляется ответ, а таблица показывает новый снимок Excel, если за этот запрос книга была прочитана." },
] as const;

const liveLinks: Record<Stage, [number, number, string]> = {
  request: [0, 1, "Запрос"], initialize: [3, 4, "initialize"], "tools/list": [3, 4, "tools/list"],
  "model.request": [1, 2, "responses.create"], "model.choice": [2, 1, "function_call"],
  "tools/call": [3, 4, "tools/call"], "model.result": [1, 2, "function_call_output"],
  confirm: [0, 1, "Решение"], response: [1, 0, "Ответ"],
};

type Arrow = { from: number; to: number; caption: string; color: string };

function Diagram({ arrow, title, caption, state }: { arrow: Arrow; title: string; caption?: string; state?: TraceEvent["state"] }) {
  const left = Math.min(arrow.from, arrow.to) * 16.6667 + 8.3333;
  const right = Math.max(arrow.from, arrow.to) * 16.6667 + 8.3333;
  const reversed = arrow.from > arrow.to;

  return (
    <div className="diagram">
      <div className="participants">
        {actors.map(({ label, tone, Icon }, index) => (
          <div key={label} className={`participant ${index === arrow.from || index === arrow.to ? "is-active" : ""}`}>
            <span className={`role-icon ${tone}`}><Icon size={18} strokeWidth={1.7} /></span>
            <span className="participant-label">{label}</span>
            <span className="participant-compact" aria-label={label}>{compactLabels[index]}</span>
          </div>
        ))}
      </div>
      <div className="lifelines" aria-hidden="true">
        {actors.map(actor => <span key={actor.label} />)}
        <div className={`flow-arrow ${reversed ? "is-reverse" : ""}`}
          key={`${arrow.from}-${arrow.to}-${arrow.caption}-${title}`}
          style={{ left: `${left}%`, width: `${right - left}%`, color: arrow.color }}>
          <code>{arrow.caption}</code>
        </div>
      </div>
      <div className="diagram-text">
        {state === "error" ? <CircleAlert size={17} className="red-ink" /> : <ArrowRight size={17} />}
        <div><strong>{title}</strong>{caption && <p>{caption}</p>}</div>
      </div>
    </div>
  );
}

function TraceDetails({ event, selected, onSelect }: { event: TraceEvent; selected: boolean; onSelect: () => void }) {
  const StatusIcon = event.state === "running" ? Clock3 : event.state === "ok" ? Check : CircleAlert;
  return (
    <details className={`trace-item ${selected ? "selected" : ""}`}>
      <summary onClick={onSelect}>
        <span className={`event-state ${event.state}`} title={{running: "Выполняется", ok: "Завершено", error: "Ошибка"}[event.state]}>
          <StatusIcon size={13} />
        </span>
        <span className="trace-label"><code>{event.tool || event.stage}</code><small>{event.title}</small></span>
        <span className="trace-time">{event.durationMs === undefined ? "…" : `${event.durationMs} мс`}</span>
        <ChevronDown className="summary-chevron" size={14} />
      </summary>
      <div className="trace-payload">
        <small>{new Date(event.at).toLocaleTimeString("ru-RU")} · шаг {event.sequence} · {event.operationId}</small>
        {event.args !== undefined && <><b>Параметры</b><pre>{JSON.stringify(event.args, null, 2)}</pre></>}
        {event.result !== undefined && <><b>Результат</b><pre>{JSON.stringify(event.result, null, 2)}</pre></>}
        {event.error && <><b>Ошибка</b><pre>{event.error}</pre></>}
      </div>
    </details>
  );
}

export function TracePanel({ events }: { events: TraceEvent[] }) {
  const [mode, setMode] = useState<"guide" | "live">("guide");
  const [step, setStep] = useState(0);
  const [playing, setPlaying] = useState(false);
  const [selected, setSelected] = useState<string | null>(null);
  const [group, setGroup] = useState<string | null>(null);
  const [technical, setTechnical] = useState(false);
  const operations = [...new Set(events.map(event => event.operationId))].reverse();
  const operation = group && operations.includes(group) ? group : operations[0];
  const active = events.filter(event => event.operationId === operation);
  const mcpEvents = active.filter(event => ["initialize", "tools/list", "tools/call"].includes(event.stage));
  const eventId = (event: TraceEvent) => `${event.operationId}:${event.sequence}`;
  const current = active.find(event => eventId(event) === selected) ?? active.at(-1);
  const guide = explanation[step];

  useEffect(() => {
    if (!playing || mode !== "guide") return;
    const timer = window.setTimeout(() => {
      if (step >= explanation.length - 1) setPlaying(false);
      else setStep(step + 1);
    }, 2100);
    return () => window.clearTimeout(timer);
  }, [mode, playing, step]);

  let [from, to, caption] = current ? liveLinks[current.stage] : [3, 4, ""];
  if (current?.stage === "tools/call" && current.state !== "running") [from, to] = [to, from];
  const tones = { blue: "#2779b8", neutral: "#657075", teal: "#08796d", violet: "#8558b4", amber: "#a86b28" };
  const arrow = mode === "guide"
    ? { ...guide, color: tones[actors[guide.to].tone] }
    : { from, to, caption: current?.stage === "tools/call" ? current.tool || caption : caption,
        color: current?.state === "error" ? "#b54048" : tones[actors[to].tone] };

  const resultLabel = current?.error || (current?.durationMs === undefined
    ? "Выполняется" : `Завершено за ${current.durationMs} мс`);
  const changeStep = (value: number) => { setPlaying(false); setStep(value); };
  return (
    <section id="under-the-hood" className="trace-section" aria-labelledby="trace-heading">
      <div className="section-heading">
        <div className="section-title"><span className="role-icon violet"><Network size={20} /></span><h2 id="trace-heading">Что происходит за кадром</h2></div>
        <a href="#workspace" className="text-link">К рабочей области <ArrowDown size={15} className="upside-down" /></a>
      </div>

      <div className="trace-grid">
        <article className="panel flow-card">
          <div className="panel-header">
            <div><h3>Как AI-агент работает с Excel</h3>
              <p>{mode === "guide" ? "Учебная схема · шаг за шагом" : "События выбранной операции"}</p></div>
            <div className="segmented" aria-label="Режим схемы">
              <button type="button" aria-pressed={mode === "guide"} onClick={() => setMode("guide")}>Схема</button>
              <button type="button" aria-pressed={mode === "live"} onClick={() => { setMode("live"); setPlaying(false); setSelected(null); }}>Выполнение</button>
            </div>
          </div>
          {mode === "guide" || current
            ? <Diagram arrow={arrow} title={mode === "guide" ? guide.title : current!.title}
                caption={mode === "guide" ? guide.text : resultLabel} state={mode === "live" ? current?.state : undefined} />
            : <div className="trace-empty">Обновите таблицу или отправьте команду, чтобы увидеть события.</div>}
          {mode === "guide" ? (
            <div className="flow-controls">
              <button className="icon-button" onClick={() => changeStep(0)} disabled={step === 0 && !playing} title="Сначала" aria-label="Сначала"><RotateCcw size={16} /></button>
              <button className="icon-button" onClick={() => changeStep(step - 1)} disabled={step === 0} title="Предыдущий шаг" aria-label="Предыдущий шаг"><ChevronLeft size={19} /></button>
              <button className="icon-button" onClick={() => changeStep(step + 1)} disabled={step === explanation.length - 1} title="Следующий шаг" aria-label="Следующий шаг"><ChevronRight size={19} /></button>
              <button className="button secondary play-control" onClick={() => {
                if (!playing && step === explanation.length - 1) setStep(0);
                setPlaying(!playing);
              }}>{playing ? <Pause size={14} /> : <Play size={14} />} {playing ? "Пауза" : "Воспроизвести"}</button>
              <span className="step-count">Шаг {step + 1} из {explanation.length}</span>
            </div>
          ) : (
            <div className="flow-controls"><span className="step-count">{active.length ? `${active.length} событий` : "Событий пока нет"}</span>
              {current && <button className="button secondary" onClick={() => setSelected(null)}>К последнему шагу</button>}
            </div>
          )}
        </article>

        <article className="panel journal">
          <div className="panel-header">
            <div><h3>Журнал MCP</h3><p>Фактические вызовы сервера</p></div>
            <button className="icon-button json-toggle" aria-label="Показать весь ход выполнения" title="Показать весь ход выполнения"
              aria-pressed={technical} onClick={() => setTechnical(!technical)}><Code2 size={17} /></button>
          </div>
          {operations.length > 0 && (
            <label className="operation-picker"><span>Операция</span>
              <select value={operation} onChange={event => { setGroup(event.target.value); setSelected(null); }}>
                {operations.map((id, index) => {
                  const first = events.find(event => event.operationId === id)!;
                  return <option key={id} value={id}>{index === 0 ? "Последняя · " : ""}{new Date(first.at).toLocaleTimeString("ru-RU")} · {first.title}</option>;
                })}
              </select>
            </label>
          )}
          <div className="trace-list">
            {(technical ? active : mcpEvents).map(event => (
              <TraceDetails key={eventId(event)} event={event} selected={selected === eventId(event)}
                onSelect={() => { setSelected(eventId(event)); setMode("live"); setPlaying(false); }} />
            ))}
            {!(technical ? active : mcpEvents).length && (
              <p className="trace-empty">{active.length ? "В этой операции MCP не вызывался." : "Вызовы появятся после подключения к Excel."}</p>
            )}
          </div>
        </article>
      </div>
    </section>
  );
}
