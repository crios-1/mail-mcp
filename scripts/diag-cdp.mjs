// Diagnostic: launch plain chrome.exe (no Playwright flags) and attach via CDP.
import { chromium } from "playwright";
import { spawn } from "node:child_process";
import path from "node:path";
import os from "node:os";

const exe = "C:/Program Files/Google/Chrome/Application/chrome.exe";
const dir = path.join(os.tmpdir(), "owa-cdp-diag");
const proc = spawn(exe, ["--remote-debugging-port=9333", `--user-data-dir=${dir}`, "--headless=new", "about:blank"]);
await new Promise((r) => setTimeout(r, 3000));
const b = await chromium.connectOverCDP("http://127.0.0.1:9333");
const p = b.contexts()[0].pages()[0] ?? (await b.contexts()[0].newPage());
for (const url of ["https://outlook.office.com/mail/", "https://outlook.office365.com/mail/"]) {
  try {
    const r = await p.goto(url, { waitUntil: "commit", timeout: 60000 });
    console.log("OK  ", url, r?.status(), "->", p.url().slice(0, 80));
  } catch (e) { console.log("FAIL", url, e.message.split("\n")[0]); }
}
await b.close();
proc.kill();
