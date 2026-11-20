import { Workspace } from "./workspace";

export default function Home() {
  return <Workspace defaultGuide={process.env.VERCEL ? "cloud" : "local"} />;
}
