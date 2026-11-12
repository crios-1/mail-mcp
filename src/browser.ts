import { chromium, type BrowserContext, type Page } from "playwright";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));

/** Browser: "chrome" (default), "msedge", or "chromium" (Playwright's bundled build). */
export const CHANNEL = process.env.OUTLOOK_BROWSER_CHANNEL ?? "chrome";

/** Persistent profile per browser: keeps the Office 365 session cookies between runs. */
export const PROFILE_DIR =
  process.env.OUTLOOK_PROFILE_DIR ?? path.resolve(here, "..", ".browser-profile", CHANNEL);

export const OWA_URL = process.env.OUTLOOK_URL ?? "https://outlook.office.com/mail/";

/** Max parallel tabs (= max parallel tool calls). */
export const POOL_SIZE = Math.max(1, Number(process.env.OUTLOOK_CONCURRENCY ?? 3));

export async function launch(headless: boolean): Promise<BrowserContext> {
  return chromium.launchPersistentContext(PROFILE_DIR, {
    headless,
    channel: CHANNEL === "chromium" ? undefined : CHANNEL,
    viewport: { width: 1400, height: 900 },
    locale: "en-US",
    // Corporate network blocks QUIC (HTTP/3 over UDP) to outlook.office.com -> force TCP.
    args: ["--disable-quic"],
  });
}

let contextPromise: Promise<BrowserContext> | undefined;
const idle: Page[] = [];
let created = 0;

function getContext(): Promise<BrowserContext> {
  contextPromise ??= launch(process.env.OUTLOOK_HEADLESS !== "false").catch((e) => {
    contextPromise = undefined;
    throw e;
  });
  return contextPromise;
}

async function openOutlookPage(): Promise<Page> {
  const ctx = await getContext();
  const page = created === 0 && ctx.pages()[0] ? ctx.pages()[0] : await ctx.newPage();
  created++;
  try {
    await page.goto(OWA_URL, { waitUntil: "commit", timeout: 120_000 });
    await page.waitForLoadState("domcontentloaded", { timeout: 120_000 }).catch(() => undefined);
    if (/login\.microsoftonline\.com|login\.live\.com/.test(page.url())) {
      throw new Error("Not logged in. Run `npm run login` once to sign in interactively.");
    }
    await page.getByRole("listbox").first().waitFor({ timeout: 60_000 });
    return page;
  } catch (e) {
    created--;
    await page.close().catch(() => undefined);
    throw e;
  }
}

/**
 * Borrow a logged-in Outlook tab. The p-queue in index.ts limits concurrency
 * to POOL_SIZE, so at most POOL_SIZE tabs exist.
 */
export async function acquirePage(): Promise<Page> {
  while (idle.length) {
    const p = idle.pop()!;
    if (!p.isClosed()) return p;
    created--;
  }
  return openOutlookPage();
}

export function releasePage(page: Page, broken = false): void {
  if (broken || page.isClosed()) {
    created--;
    void page.close().catch(() => undefined);
  } else {
    idle.push(page);
  }
}

export async function close(): Promise<void> {
  const ctx = await contextPromise?.catch(() => undefined);
  await ctx?.close();
  contextPromise = undefined;
  idle.length = 0;
  created = 0;
}
