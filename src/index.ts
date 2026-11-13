#!/usr/bin/env node
/**
 * Entry point. Transport is chosen by MCP_TRANSPORT (or --http):
 *   stdio (default) - local client spawns this process (Claude Code/Desktop on this machine)
 *   http            - Streamable HTTP server, expose it with Tailscale serve/funnel
 */
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { close } from "./browser.js";
import { startHttp } from "./http.js";
import { createServer } from "./tools.js";

const mode = process.argv.includes("--http") ? "http" : (process.env.MCP_TRANSPORT ?? "stdio");

const shutdown = async () => {
  await close().catch(() => undefined);
  process.exit(0);
};
process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);

if (mode === "http") {
  await startHttp();
} else {
  await createServer().connect(new StdioServerTransport());
}
