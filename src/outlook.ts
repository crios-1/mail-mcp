/**
 * Outlook on the web automation. Every function works on its own tab (Page)
 * from the pool, so calls can run in parallel. OWA has no stable ids, so we
 * rely on ARIA roles/labels; adjust the selectors here when the UI changes.
 */
import type { Locator, Page } from "playwright";

export interface MailSummary {
  index: number;
  id: string | null;
  text: string;
}

const FOLDERS: Record<string, string> = {
  inbox: "Inbox",
  drafts: "Drafts",
  sent: "Sent Items",
  deleted: "Deleted Items",
};

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

async function openFolder(page: Page, folder: string): Promise<void> {
  const name = FOLDERS[folder.toLowerCase()] ?? folder;
  await page.getByRole("treeitem", { name: new RegExp(`^${escapeRe(name)}`, "i") }).first().click();
  await page.waitForTimeout(1500);
}

function rows(page: Page): Locator {
  return page.getByRole("listbox").first().getByRole("option");
}

export async function listMails(page: Page, folder = "inbox", limit = 20): Promise<MailSummary[]> {
  await openFolder(page, folder);
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

async function selectRow(page: Page, folder: string, index: number): Promise<Locator> {
  await openFolder(page, folder);
  const row = rows(page).nth(index);
  if (!(await row.count())) throw new Error(`No mail at index ${index} in ${folder}`);
  await row.click();
  await page.waitForTimeout(1000);
  return row;
}

export async function readMail(page: Page, folder: string, index: number): Promise<string> {
  await selectRow(page, folder, index);
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
  await page.getByRole("button", { name: /^New (mail|email)/i }).first().click();

  const fillRecipients = async (label: RegExp, list: string[]) => {
    const box = page.getByRole("textbox", { name: label }).first();
    for (const addr of list) {
      await box.fill(addr);
      await page.keyboard.press("Enter");
    }
  };
  await fillRecipients(/^To/i, to);
  if (cc.length) await fillRecipients(/^Cc/i, cc);

  await page.getByPlaceholder(/subject/i).first().fill(subject);
  const editor = page.getByRole("textbox", { name: /message body/i }).first();
  await editor.click();
  await page.keyboard.insertText(body);

  // OWA autosaves drafts; give it time, then close the compose form without sending.
  await page.waitForTimeout(3000);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(1000);
}
