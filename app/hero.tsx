import { ArrowDownRight } from "lucide-react";

export function Hero() {
  return (
    <section className="hero" aria-labelledby="hero-title">
      <h1 id="hero-title">AI-агент <span>«Заявки»</span></h1>
      <p>Управление заявками через OpenAI и Excel MCP</p>
      <a className="hero-link" href="#under-the-hood"><ArrowDownRight size={14} /> Как проходит запрос</a>
    </section>
  );
}
