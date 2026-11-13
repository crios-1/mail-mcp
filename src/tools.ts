/** Outlook web mail actions as MCP tools. Shared by the stdio and HTTP transports. */
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import PQueue from "p-queue";
import type { Page } from "playwright";
import { z } from "zod";
import { acquirePage, POOL_SIZE, releasePage } from "./browser.js";
import { createDraft, deleteMail, listMails, readMail, searchMails } from "./outlook.js";

const text = (t: string) => ({ content: [{ type: "text" as const, text: t }] });
const folder = z.string().default("inbox").describe("inbox | drafts | sent | deleted | exact folder name");

/**
 * Load handling: at most POOL_SIZE tool calls run at once (one browser tab
 * each); the rest wait in the queue. Each job has a timeout, and when too
 * many jobs are waiting new calls are refused instead of piling up.
 * The queue is module-level, so it is shared by every MCP session/client.
 */
const MAX_WAITING = Number(process.env.OUTLOOK_MAX_QUEUE ?? 50);
const JOB_TIMEOUT_MS = Number(process.env.OUTLOOK_JOB_TIMEOUT_MS ?? 120_000);
const queue = new PQueue({ concurrency: POOL_SIZE });

async function run<T>(job: (page: Page) => Promise<T>): Promise<T> {
  if (queue.size >= MAX_WAITING) {
    throw new Error(`Server busy: ${queue.size} calls waiting (max ${MAX_WAITING}). Retry later.`);
  }
  return queue.add(
    async () => {
      const page = await acquirePage();
      let broken = false;
      let timer: NodeJS.Timeout | undefined;
      try {
        return await Promise.race([
          job(page),
          new Promise<never>((_, reject) => {
            timer = setTimeout(() => reject(new Error(`Timed out after ${JOB_TIMEOUT_MS} ms`)), JOB_TIMEOUT_MS);
          }),
        ]);
      } catch (e) {
        broken = true; // tab may be in an unknown state: discard it
        console.error(`Tool call failed: ${(e as Error).message.split("\n")[0]}`);
        throw e;
      } finally {
        clearTimeout(timer);
        releasePage(page, broken);
      }
    },
    { throwOnTimeout: true },
  ) as Promise<T>;
}

/** Lists one inbox mail through the normal queue; throws when not logged in. */
export async function testConnection(): Promise<number> {
  return (await run((p) => listMails(p, "inbox", 1))).length;
}

/** Builds a fresh MCP server with all tools registered (cheap; the browser is shared). */
export function createServer(): McpServer {
  const server = new McpServer({ name: "outlook-browser", version: "0.2.0" });

  server.tool(
    "list_mails",
    "List the newest mails in a folder. Returns index (use it in other tools) and a summary.",
    { folder, limit: z.number().int().min(1).max(100).default(20) },
    async ({ folder, limit }) => text(JSON.stringify(await run((p) => listMails(p, folder, limit)), null, 2)),
  );

  server.tool(
    "read_mail",
    "Open a mail by index and return its content. Index comes from list_mails (pass the same folder) or from search_mails (pass the same query).",
    {
      folder,
      index: z.number().int().min(0),
      query: z.string().optional().describe("The search_mails query the index belongs to; omit for list_mails indexes"),
    },
    async ({ folder, index, query }) => text(await run((p) => readMail(p, folder, index, query))),
  );

  server.tool(
    "search_mails",
    "Search all mail folders with Outlook search. Query syntax as in the Outlook search box: words, from:name, to:name, subject:word, hasattachments:yes, received:>=2026-01-01. Returns index (use with read_mail plus the same query) and a summary.",
    { query: z.string().min(1), limit: z.number().int().min(1).max(100).default(20) },
    async ({ query, limit }) => text(JSON.stringify(await run((p) => searchMails(p, query, limit)), null, 2)),
  );

  server.tool(
    "delete_mail",
    "Move a mail to Deleted Items (recoverable). Call list_mails first: indexes shift after deletes. Don't run deletes in parallel on the same folder.",
    { folder, index: z.number().int().min(0) },
    async ({ folder, index }) => text(`Deleted: ${await run((p) => deleteMail(p, folder, index))}`),
  );

  server.tool(
    "create_draft",
    "Create a draft (concept) mail. It is saved in Drafts, never sent.",
    {
      to: z.array(z.string().email()).min(1),
      cc: z.array(z.string().email()).default([]),
      subject: z.string().min(1).describe("Required: used to verify the draft was saved"),
      body: z.string(),
    },
    async ({ to, cc, subject, body }) => {
      await run((p) => createDraft(p, to, subject, body, cc));
      return text(`Draft saved: "${subject}"`);
    },
  );

  server.tool(
    "queue_status",
    "Show load: running and waiting calls.",
    {},
    async () => text(JSON.stringify({ running: queue.pending, waiting: queue.size, concurrency: POOL_SIZE, maxWaiting: MAX_WAITING })),
  );

  return server;
}
