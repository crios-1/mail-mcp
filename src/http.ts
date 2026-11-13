/**
 * Streamable HTTP transport (stateless) with a shared-secret token.
 *
 * Meant to sit behind `tailscale serve` (tailnet only) or `tailscale funnel`
 * (public, needed for claude.ai connectors). Binds to localhost by default so
 * the only way in is through Tailscale.
 *
 * Auth: either header `Authorization: Bearer <token>` or the token as a path
 * segment, `/mcp/<token>`, for clients that cannot send custom headers
 * (claude.ai custom connectors).
 */
import { createServer as createHttpServer, type IncomingMessage, type ServerResponse } from "node:http";
import { timingSafeEqual } from "node:crypto";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { handleSetup } from "./setup.js";
import { createServer } from "./tools.js";

const MAX_BODY_BYTES = 1_000_000;

function tokenMatches(given: string | undefined, expected: string): boolean {
  if (!given) return false;
  const a = Buffer.from(given);
  const b = Buffer.from(expected);
  return a.length === b.length && timingSafeEqual(a, b);
}

function readJson(req: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    let size = 0;
    req.on("data", (c: Buffer) => {
      size += c.length;
      if (size > MAX_BODY_BYTES) {
        reject(new Error("Request body too large"));
        req.destroy();
      } else chunks.push(c);
    });
    req.on("end", () => {
      if (!size) return resolve(undefined);
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString("utf8")));
      } catch {
        reject(new Error("Invalid JSON"));
      }
    });
    req.on("error", reject);
  });
}

function sendJson(res: ServerResponse, status: number, body: unknown, headers: Record<string, string> = {}): void {
  res.writeHead(status, { "Content-Type": "application/json", ...headers }).end(JSON.stringify(body));
}

const rpcError = (message: string) => ({ jsonrpc: "2.0", error: { code: -32000, message }, id: null });

export async function startHttp(): Promise<void> {
  const token = process.env.MCP_AUTH_TOKEN;
  if (!token || token.length < 24) {
    throw new Error("MCP_AUTH_TOKEN must be set (min 24 chars) in HTTP mode. Generate one: openssl rand -hex 32");
  }
  const host = process.env.MCP_HOST ?? "127.0.0.1";
  const port = Number(process.env.MCP_PORT ?? 8787);

  const http = createHttpServer(async (req, res) => {
    const url = new URL(req.url ?? "/", "http://localhost");
    const parts = url.pathname.split("/").filter(Boolean);

    if (url.pathname === "/healthz") return sendJson(res, 200, { ok: true });

    // Setup page: /setup/<token>/... (token in path so a plain browser can open it).
    if (parts[0] === "setup") {
      if (!tokenMatches(parts[1], token)) return sendJson(res, 404, { error: "Not found" });
      if (parts.length === 2 && !url.pathname.endsWith("/")) {
        res.writeHead(302, { Location: `${url.pathname}/` }).end(); // relative URLs in the page need the slash
        return;
      }
      return handleSetup(req, res, parts.slice(2), token);
    }
    if (parts[0] !== "mcp" || parts.length > 2) return sendJson(res, 404, { error: "Not found" });

    const bearer = /^Bearer\s+(.+)$/i.exec(req.headers.authorization ?? "")?.[1];
    if (!tokenMatches(bearer, token) && !tokenMatches(parts[1], token)) {
      return sendJson(res, 401, rpcError("Unauthorized"), { "WWW-Authenticate": "Bearer" });
    }

    // Stateless mode: no server-initiated streams or session resumption.
    if (req.method !== "POST") {
      return sendJson(res, 405, rpcError("Method not allowed"), { Allow: "POST" });
    }

    let body: unknown;
    try {
      body = await readJson(req);
    } catch (e) {
      return sendJson(res, 400, rpcError((e as Error).message));
    }

    // One server+transport per request (SDK's stateless pattern). The browser,
    // tab pool and job queue are module-level, so they are shared across requests.
    const server = createServer();
    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
    res.on("close", () => {
      void transport.close();
      void server.close();
    });
    try {
      await server.connect(transport);
      await transport.handleRequest(req, res, body);
    } catch (e) {
      console.error("MCP request failed:", e);
      if (!res.headersSent) sendJson(res, 500, rpcError("Internal server error"));
    }
  });

  await new Promise<void>((resolve) => http.listen(port, host, resolve));
  console.error(`outlook-browser MCP listening on http://${host}:${port}/mcp`);
  console.error(`Setup page: http://${host}:${port}/setup/<MCP_AUTH_TOKEN>/`);
}
