import { chromium, type BrowserContext, type Page } from "playwright";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));

/**
 * Browser: "chrome", "msedge", or "chromium" (Playwright's bundled build).
 * Default: chromium on Linux (no Chrome needed on a headless server), chrome elsewhere.
 */
export const CHANNEL =
  process.env.OUTLOOK_BROWSER_CHANNEL ?? (process.platform === "linux" ? "chromium" : "chrome");

/** Persistent profile per browser: keeps the Office 365 session cookies between runs. */
export const PROFILE_DIR =
  process.env.OUTLOOK_PROFILE_DIR ?? path.resolve(here, "..", ".browser-profile", CHANNEL);

export const OWA_URL = process.env.OUTLOOK_URL ?? "https://outlook.cloud.microsoft/mail/";

/** Max parallel tabs (= max parallel tool calls). */
export const POOL_SIZE = Math.max(1, Number(process.env.OUTLOOK_CONCURRENCY ?? 3));

/**
 * Headless browsers send "HeadlessChrome" in their user agent, and Outlook serves
 * them a blank page. Report a normal Chrome UA for the same major version instead.
 */
async function userAgent(): Promise<string> {
  if (process.env.OUTLOOK_USER_AGENT) return process.env.OUTLOOK_USER_AGENT;
  const b = await chromium.launch({ headless: true, channel: CHANNEL });
  const major = b.version().split(".")[0];
  await b.close();
  const os =
    process.platform === "win32" ? "Windows NT 10.0; Win64; x64"
    : process.platform === "darwin" ? "Macintosh; Intel Mac OS X 10_15_7"
    : "X11; Linux x86_64";
  return `Mozilla/5.0 (${os}) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/${major}.0.0.0 Safari/537.36`;
}

export async function launch(headless: boolean): Promise<BrowserContext> {
  return chromium.launchPersistentContext(PROFILE_DIR, {
    headless,
    // "chromium" = full Chromium in new headless mode (not the stripped headless shell).
    channel: CHANNEL,
    userAgent: headless ? await userAgent() : undefined,
    viewport: { width: 1400, height: 900 },
    locale: "en-US",
    // Corporate network blocks QUIC (HTTP/3 over UDP) to outlook.office.com -> force TCP.
    // Linux: store cookies with a fixed key instead of the desktop keyring, so a session
    // made by the (headed) login is readable by the headless server under systemd.
    args: ["--disable-quic", "--password-store=basic"],
  });
}

/**
 * Some Outlook front-end IPs time out now and then (DNS rotates between them),
 * so retry navigation a few times with a shorter timeout per attempt.
 */
export async function gotoOutlook(page: Page, attempts = 4): Promise<void> {
  for (let i = 1; ; i++) {
    try {
      await page.goto(OWA_URL, { waitUntil: "commit", timeout: 30_000 });
      return;
    } catch (e) {
      if (i >= attempts) throw e;
      console.error(`Outlook navigation attempt ${i} failed: ${(e as Error).message.split("\n")[0]}; retrying`);
    }
  }
}

class NotLoggedIn extends Error {}

let contextPromise: Promise<BrowserContext> | undefined;
const idle: Page[] = [];
let created = 0;

export function getContext(): Promise<BrowserContext> {
  contextPromise ??= launch(process.env.OUTLOOK_HEADLESS !== "false").catch((e) => {
    contextPromise = undefined;
    throw e;
  });
  return contextPromise;
}

async function openOutlookPage(): Promise<Page> {
  const ctx = await getContext();
  // Always a new tab: the startup tab may be in use by the /setup login page.
  const page = await ctx.newPage();
  created++;
  try {
    // The network to Outlook is flaky: retry the whole "load Outlook and wait for
    // the mail list" step, not only the navigation.
    for (let attempt = 1; ; attempt++) {
      try {
        await gotoOutlook(page);
        await page.waitForLoadState("domcontentloaded", { timeout: 60_000 }).catch(() => undefined);
        if (/login\.microsoftonline\.com|login\.live\.com|login\.microsoft\.com/.test(page.url())) {
          throw new NotLoggedIn("Not logged in. Sign in again via the setup page (/setup/<token>/).");
        }
        await page.getByRole("listbox").first().waitFor({ timeout: 45_000 });
        return page;
      } catch (e) {
        if (e instanceof NotLoggedIn || attempt >= 3) throw e;
        console.error(`Opening Outlook tab, attempt ${attempt} failed: ${(e as Error).message.split("\n")[0]}; retrying`);
      }
    }
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
