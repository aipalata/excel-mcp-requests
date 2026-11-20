import type { NextConfig } from "next";

const config: NextConfig = {
  poweredByHeader: false,
  serverExternalPackages: ["@modelcontextprotocol/sdk", "openai"],
};

export default config;
