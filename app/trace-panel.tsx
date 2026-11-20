"use client";

import { ArrowDown, ArrowRight, Check, ChevronDown, ChevronLeft, ChevronRight, CircleAlert,
  Clock3, Code2, Database, FileSpreadsheet, LayoutPanelTop, Network, Pause, Play, Plug, RotateCcw,
  Sparkles, UserRound, type LucideIcon } from "lucide-react";
import { useEffect, useState } from "react";
import type { Stage, TraceEvent } from "@/lib/contracts";

const tones = { blue: "#2779b8", neutral: "#657075", teal: "#08796d", violet: "#8558b4", amber: "#a86b28" };
type Participant = { label: string; compact: string; tone: keyof typeof tones; Icon: LucideIcon };
export type GuideEnvironment = "local" | "cloud";

const localActors: readonly Participant[] = [
  { label: "Пользователь", compact: "Вы", tone: "blue", Icon: UserRound },
  { label: "Приложение", compact: "Сайт", tone: "neutral", Icon: LayoutPanelTop },
  { label: "OpenAI", compact: "OpenAI", tone: "teal", Icon: Sparkles },
  { label: "MCP-клиент", compact: "Клиент", tone: "violet", Icon: Plug },
  { label: "Excel MCP", compact: "Сервер", tone: "violet", Icon: Network },
  { label: "Книга Excel", compact: "Excel", tone: "amber", Icon: FileSpreadsheet },
];
const hostedActors: readonly Participant[] = [
  localActors[0],
  { ...localActors[1], label: "Vercel", compact: "Vercel" },
  ...localActors.slice(2, 5),
  { ...localActors[5], label: "Копия Excel" },
];
const cloudActors: readonly Participant[] = [
  ...hostedActors,
  { label: "Supabase", compact: "Облако", tone: "teal", Icon: Database },
];

function statusGuide(cloud: boolean) {
  return [
    { from: 0, to: 1, caption: "Сообщение", title: "Пользователь просит сменить статус", text: "Например, «Переведи заявку №101 в работу»." },
    { from: 3, to: 4, caption: "initialize", title: "MCP-клиент подключается", text: cloud
      ? "На Vercel клиент подключается по HTTPS к Python-функции, которая запускает готовый Excel MCP."
      : "Приложение запускает готовый Excel MCP на компьютере и подключается к нему через stdio." },
    { from: 3, to: 4, caption: "tools/list", title: "Клиент запрашивает каталог", text: "Сервер возвращает инструменты и схемы параметров. Приложение отбирает разрешённые операции." },
    { from: 1, to: 2, caption: "responses.create", title: "Приложение обращается к OpenAI", text: "Модель получает сообщение, инструкции и схемы функций." },
    { from: 2, to: 1, caption: "function_call", title: "Модель предлагает подготовить изменение", text: "OpenAI выбирает функцию подготовки статуса и передаёт номер заявки и новый статус." },
    { from: 3, to: 4, caption: "tools/call", title: "Приложение ищет заявку через MCP", text: "Excel MCP получает вызовы describe_sheet и read_range, чтобы определить границы листа и прочитать строки." },
    ...(cloud ? [{ from: 6, to: 1, caption: "Storage", title: "Функция скачивает текущую книгу", text: "Она узнаёт в PostgreSQL, какая версия актуальна, и скачивает её из приватного Storage во временную папку." }] : []),
    { from: 4, to: 5, caption: "read_range", title: "Excel MCP читает ячейки", text: cloud
      ? "Готовый сервер читает временную копию .xlsx. Приложение получает текущие поля заявки."
      : "Готовый сервер читает локальный файл .xlsx. Приложение получает текущие поля заявки." },
    { from: 1, to: 2, caption: "function_call_output", title: "Приложение сохраняет план", text: cloud
      ? "План хранится в PostgreSQL до пяти минут и привязан к сеансу. Модель получает результат подготовки."
      : "План хранится в памяти приложения до пяти минут. Модель получает результат подготовки." },
    { from: 1, to: 0, caption: "План", title: "Пользователь видит карточку подтверждения", text: "Карточка показывает текущий и предлагаемый статусы. На этом шаге Excel ещё хранит прежний статус." },
    { from: 0, to: 1, caption: "Подтвердить", title: "Нажатие запускает отдельный запрос", text: "Приложение получает сохранённый план и через MCP сверяет с ним текущие поля заявки." },
    { from: 4, to: 5, caption: "write_range", title: "Excel MCP меняет ячейку статуса", text: cloud
      ? "Готовый сервер записывает новый статус во временную копию Excel."
      : "Готовый сервер записывает новый статус в локальный файл Excel." },
    ...(cloud ? [
      { from: 1, to: 6, caption: "Storage", title: "Vercel загружает новую версию .xlsx", text: "Supabase Storage хранит обновлённый файл отдельно от предыдущей версии." },
      { from: 1, to: 6, caption: "SQL", title: "PostgreSQL подтверждает версию", text: "Функция Vercel передаёт новую версию в PostgreSQL. База делает её текущей, если исходная версия всё ещё актуальна, и запоминает результат операции." },
    ] : []),
    { from: 3, to: 4, caption: "read_range", title: "Приложение проверяет запись", text: "Контрольное чтение сверяет новый статус и остальные поля заявки." },
    { from: 1, to: 0, caption: "Ответ", title: "Пользователь видит сохранённый статус", text: "В чате появляется подтверждение, а таблица показывает обновлённые данные Excel." },
  ];
}

const guides = { local: statusGuide(false), cloud: statusGuide(true) };

const liveLinks: Record<Stage, [number, number, string]> = {
  request: [0, 1, "Запрос"], initialize: [3, 4, "initialize"], "tools/list": [3, 4, "tools/list"],
  "model.request": [1, 2, "responses.create"], "model.choice": [2, 1, "function_call"],
  "tools/call": [3, 4, "tools/call"], "model.result": [1, 2, "function_call_output"],
  confirm: [0, 1, "Решение"], response: [1, 0, "Ответ"],
};

type Arrow = { from: number; to: number; caption: string; color: string };

function Diagram({ actors, arrow, title, caption, state }: {
  actors: readonly Participant[]; arrow: Arrow; title: string; caption?: string; state?: TraceEvent["state"];
}) {
  const lane = 100 / actors.length;
  const left = (Math.min(arrow.from, arrow.to) + .5) * lane;
  const right = (Math.max(arrow.from, arrow.to) + .5) * lane;
  const columns = { gridTemplateColumns: `repeat(${actors.length}, minmax(0, 1fr))` };
  const reversed = arrow.from > arrow.to;

  return (
    <div className="diagram">
      <div className="participants" style={columns}>
        {actors.map(({ label, compact, tone, Icon }, index) => (
          <div key={label} className={`participant ${index === arrow.from || index === arrow.to ? "is-active" : ""}`}>
            <span className={`role-icon ${tone}`}><Icon size={18} strokeWidth={1.7} /></span>
            <span className="participant-label">{label}</span>
            <span className="participant-compact" aria-label={label}>{compact}</span>
          </div>
        ))}
      </div>
      <div className="lifelines" style={columns} aria-hidden="true">
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

export function TracePanel({ events, defaultGuide }: { events: TraceEvent[]; defaultGuide: GuideEnvironment }) {
  const [mode, setMode] = useState<"guide" | "live">("guide");
  const [environment, setEnvironment] = useState(defaultGuide);
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
  const explanation = guides[environment];
  const guide = explanation[step];
  const actors = mode === "guide" ? (environment === "cloud" ? cloudActors : localActors)
    : (defaultGuide === "cloud" ? hostedActors : localActors);

  useEffect(() => {
    if (!playing || mode !== "guide") return;
    const timer = window.setTimeout(() => {
      if (step >= explanation.length - 1) setPlaying(false);
      else setStep(step + 1);
    }, 3400);
    return () => window.clearTimeout(timer);
  }, [mode, playing, step, explanation.length]);

  let [from, to, caption] = current ? liveLinks[current.stage] : [3, 4, ""];
  if (current?.stage === "tools/call" && current.state !== "running") [from, to] = [to, from];
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
              <p>{mode === "guide" ? "Учебный сценарий · смена статуса" : "События выбранной операции"}</p></div>
            <div className="segmented" aria-label="Режим схемы">
              <button type="button" aria-pressed={mode === "guide"} onClick={() => setMode("guide")}>Схема</button>
              <button type="button" aria-pressed={mode === "live"} onClick={() => { setMode("live"); setPlaying(false); setSelected(null); }}>Выполнение</button>
            </div>
          </div>
          {mode === "guide" && (
            <div className="segmented guide-environments" role="group" aria-label="Среда учебной схемы">
              <button type="button" aria-pressed={environment === "local"}
                onClick={() => { setEnvironment("local"); changeStep(0); }}>Локально</button>
              <button type="button" aria-pressed={environment === "cloud"}
                onClick={() => { setEnvironment("cloud"); changeStep(0); }}>Vercel + Supabase</button>
            </div>
          )}
          {mode === "guide" || current
            ? <Diagram actors={actors} arrow={arrow} title={mode === "guide" ? guide.title : current!.title}
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
