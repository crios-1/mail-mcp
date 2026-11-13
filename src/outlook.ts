/**
 * Outlook on the web automation. Every function works on its own tab (Page)
 * from the pool, so calls can run in parallel. OWA has no stable ids, so we
 * rely on ARIA roles/labels; adjust the selectors here when the UI changes.
 */
import type { Locator, Page } from "playwright";
import { gotoOutlook } from "./browser.js";

export interface MailSummary {
  index: number;
  id: string | null;
  text: string;
}

/** Folder names per UI language (English, Dutch). */
const FOLDERS: Record<string, string[]> = {
  inbox: ["Inbox", "Postvak IN"],
  drafts: ["Drafts", "Concepten"],
  sent: ["Sent Items", "Verzonden items"],
  deleted: ["Deleted Items", "Verwijderde items"],
  junk: ["Junk Email", "Ongewenste e-mail"],
  archive: ["Archive", "Archiveren"],
};

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const searchBox = (page: Page) => page.getByRole("combobox", { name: /^(Search|Zoek)/i }).first();

/** Tabs are reused: leave search mode left behind by an earlier search_mails call. */
async function exitSearch(page: Page): Promise<void> {
  const box = searchBox(page);
  if (!(await box.count()) || !(await box.inputValue().catch(() => ""))) return;
  await box.fill("");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(500);
}

async function openFolder(page: Page, folder: string): Promise<void> {
  await exitSearch(page);
  const names = FOLDERS[folder.toLowerCase()] ?? [folder];
  // Tree item labels have a leading space and trailing unread counts: " Postvak IN 11 ongelezen".
  const name = new RegExp(`^\\s*(${names.map(escapeRe).join("|")})(\\s|$)`, "i");
  const item = page.getByRole("treeitem", { name }).first();
  if (!(await item.count())) throw new Error(`Folder not found: ${folder}`);
  await item.click();
  await page.waitForTimeout(1500);
  await rows(page).first().waitFor({ timeout: 10_000 }).catch(() => undefined); // empty folders are fine
}

function rows(page: Page): Locator {
  return page.getByRole("listbox").first().getByRole("option");
}

async function summarize(page: Page, limit: number): Promise<MailSummary[]> {
  const all = rows(page);
  const count = Math.min(await all.count(), limit);
  const result: MailSummary[] = [];
  for (let i = 0; i < count; i++) {
    const row = all.nth(i);
    result.push({
      index: i,
      id: await row.getAttribute("data-convid"),
      text: ((await row.getAttribute("aria-label")) ?? (await row.innerText())).trim(),
    });
  }
  return result;
}

export async function listMails(page: Page, folder = "inbox", limit = 20): Promise<MailSummary[]> {
  await openFolder(page, folder);
  return summarize(page, limit);
}

/**
 * Runs an Outlook search (same syntax as the search box: words, from:, subject:,
 * hasattachments:yes, received:>=2026-01-01, ...). Searches all folders by default.
 */
async function runSearch(page: Page, query: string): Promise<void> {
  const box = searchBox(page);
  await box.click();
  await box.fill(query);
  await page.keyboard.press("Enter");
  // Outlook swaps the folder list for an empty list, then fills in results.
  // Wait until the list has rows and is stable for two checks (max ~20 s;
  // a search without results ends with an empty list).
  let last = "";
  let stable = 0;
  for (let waited = 0; waited < 20_000 && stable < 2; waited += 750) {
    await page.waitForTimeout(750);
    const n = await rows(page).count();
    const sig = n ? `${n}:${await rows(page).first().getAttribute("aria-label").catch(() => "")}` : "";
    stable = sig && sig === last ? stable + 1 : 0;
    last = sig;
  }
}

export async function searchMails(page: Page, query: string, limit = 20): Promise<MailSummary[]> {
  await runSearch(page, query);
  return summarize(page, limit);
}

/** Selects a row in a folder, or in the results of `query` when given. */
async function selectRow(page: Page, folder: string, index: number, query?: string): Promise<Locator> {
  if (query) await runSearch(page, query);
  else await openFolder(page, folder);
  const row = rows(page).nth(index);
  if (!(await row.count())) throw new Error(`No mail at index ${index} in ${query ? `search "${query}"` : folder}`);
  await row.click();
  await page.waitForTimeout(1000);
  return row;
}

export async function readMail(page: Page, folder: string, index: number, query?: string): Promise<string> {
  await selectRow(page, folder, index, query);
  return (await page.getByRole("main").first().innerText()).trim();
}

/** Moves the mail to Deleted Items (soft delete, recoverable). */
export async function deleteMail(page: Page, folder: string, index: number): Promise<string> {
  const row = await selectRow(page, folder, index);
  const label = (await row.getAttribute("aria-label")) ?? "";
  await page.keyboard.press("Delete");
  await page.waitForTimeout(1000);
  return label;
}

export async function createDraft(page: Page, to: string[], subject: string, body: string, cc: string[] = []): Promise<void> {
  await page.getByRole("button", { name: /^(New mail|New email|New|Nieuwe e-mail|Nieuw)$/i }).first().click();

  // Recipient fields are contenteditable divs labelled "To"/"Aan" and "Cc"/"CC".
  const fillRecipients = async (labels: string[], list: string[]) => {
    const box = page.locator(labels.map((l) => `[contenteditable="true"][aria-label="${l}"]`).join(", ")).first();
    await box.waitFor({ timeout: 15_000 });
    for (const addr of list) {
      await box.click();
      await page.keyboard.insertText(addr);
      await page.waitForTimeout(500);
      await page.keyboard.press("Enter");
    }
  };
  await fillRecipients(["To", "Aan"], to);
  if (cc.length) await fillRecipients(["Cc", "CC"], cc);

  await page.getByRole("textbox", { name: /^(Subject|Add a subject|Onderwerp)/i }).first().fill(subject);
  const editor = page.getByRole("textbox", { name: /^(Message body|Berichttekst)/i }).first();
  await editor.click();
  await page.keyboard.insertText(body);

  // Save explicitly (Ctrl+S) and wait for OWA's "Draft saved at ..." notice. Only
  // then leave by reloading: Escape can open a "discard draft?" dialog, and
  // reloading before the autosave loses the subject and body.
  await page.waitForTimeout(1000);
  await page.keyboard.press("Control+s");
  await page
    .getByText(/(Draft saved|Concept opgeslagen)/i)
    .first()
    .waitFor({ timeout: 15_000 })
    .catch(() => console.error("create_draft: no 'draft saved' notice seen; verifying in Drafts anyway"));
  await page.waitForTimeout(1500);
  await gotoOutlook(page);
  await page.getByRole("listbox").first().waitFor({ timeout: 60_000 });

  // Verify: the draft must show up in Drafts before we report success.
  await openFolder(page, "drafts");
  const draft = rows(page).filter({ hasText: subject }).first();
  for (let waited = 0; ; waited += 2000) {
    if (await draft.count()) return;
    if (waited >= 20_000) throw new Error(`Draft "${subject}" was not found in Drafts after saving. It may not have been saved.`);
    await page.waitForTimeout(2000);
    await openFolder(page, "drafts"); // re-click to refresh the list
  }
}