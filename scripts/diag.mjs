// Diagnostic: can a Playwright browser reach Outlook? Usage: node scripts/diag.mjs [msedge|chromium]
import { chromium } from "playwright";
const channel = process.argv[2] === "chromium" ? undefined : (process.argv[2] ?? "chrome");
const b = await chromium.launch({ headless: true, channel, args: ["--disable-quic"] });
const p = await b.newPage();
for (const url of ["https://www.microsoft.com", "https://login.microsoftonline.com", "https://outlook.office.com/mail/"]) {
  try {
    const r = await p.goto(url, { waitUntil: "commit", timeout: 60000 });
    console.log("OK  ", url, r?.status(), "->", p.url().slice(0, 80));
  } catch (e) { console.log("FAIL", url, e.message.split("\n")[0]); }
}
await b.close();
