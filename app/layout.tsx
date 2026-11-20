import type { Metadata } from "next";
import "@fontsource-variable/golos-text";
import "@fontsource-variable/roboto-condensed";
import "./globals.css";

export const metadata: Metadata = {
  title: "AI-агент «Заявки»",
  description: "Управление заявками через OpenAI и Excel MCP",
  icons: { icon: "/icon.svg" },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="ru"><body>{children}</body></html>;
}
