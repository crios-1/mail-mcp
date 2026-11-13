/**
 * One-time interactive login. Opens a visible browser; you sign in yourself
 * (password + MFA). The session is stored in the persistent profile so the
 * MCP server can later run headless.
 */
import { launch, OWA_URL, PROFILE_DIR } from "./browser.js";

const ctx = await launch(false);
const page = ctx.pages()[0] ?? (await ctx.newPage());
// Don't wait for "load": the Microsoft/SSO redirect chain can take long and you sign in meanwhile.
// The network to outlook.office.com is flaky here: retry a few times.
for (let attempt = 1; ; attempt++) {
  try {
    await page.goto(OWA_URL, { waitUntil: "commit", timeout: 120_000 });
    break;
  } catch (e) {
    if (attempt >= 5) throw e;
    console.log(`Attempt ${attempt} failed (${(e as Error).message.split("\n")[0]}), retrying...`);
  }
}
console.log(`Sign in in the browser window. Profile: ${PROFILE_DIR}`);
console.log("Waiting until the inbox is visible (max 5 min)...");

await page.waitForURL(/outlook\.(office|office365|live)\.com\/mail|outlook\.cloud\.microsoft\/mail/, { timeout: 300_000 });
await page.getByRole("listbox").first().waitFor({ timeout: 300_000 });
console.log("Logged in. Session saved. You can close this now.");
await ctx.close();
