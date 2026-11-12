#!/usr/bin/env node
/** MCP server (stdio) exposing Outlook web mail actions as tools. */
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import PQueue from "p-queue";
import type { Page } from "playwright";
import { z } from "zod";
import { acquirePage, close, POOL_SIZE, releasePage } from "./browser.js";
import { createDraft, deleteMail, listMails, readMail } from "./outlook.js";

const server = new McpServer({ name: "outlook-browser", version: "0.1.0" });

const text = (t: string) => ({ content: [{ type: "text" as const, text: t }] });
const folder = z.string().default("inbox").describe("inbox | drafts | sent | deleted | exact folder name");

/**
 * Load handling: at most POOL_SIZE tool calls run at once (one browser tab
 * each); the rest wait in the queue. Each job has a timeout, and when too
 * many jobs are waiting new calls are refused instead of piling up.
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
        throw e;
      } finally {
        clearTimeout(timer);
        releasePage(page, broken);
      }
    },
    { throwOnTimeout: true },
  ) as Promise<T>;
}

server.tool(
  "list_mails",
  "List the newest mails in a folder. Returns index (use it in other tools) and a summary.",
  { folder, limit: z.number().int().min(1).max(100).default(20) },
  async ({ folder, limit }) => text(JSON.stringify(await run((p) => listMails(p, folder, limit)), null, 2)),
);

server.tool(
  "read_mail",
  "Open a mail by index (from list_mails) and return its content.",
  { folder, index: z.number().int().min(0) },
  async ({ folder, index }) => text(await run((p) => readMail(p, folder, index))),
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
    subject: z.string(),
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

process.on("SIGINT", async () => { await close(); process.exit(0); });
await server.connect(new StdioServerTransport());
