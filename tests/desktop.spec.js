// Desktop-only features, with the Rust side replaced by tests/tauri-mock.js
import { test, expect } from "@playwright/test";
import { readFileSync } from "node:fs";
import { installTauriMock } from "./tauri-mock.js";
import { doc, paragraph } from "./helpers.js";

const DOC = "C:\\docs\\AITA for uninstalling Word.docx";
const DRAFT = "C:\\docs\\draft.docx";

test.beforeEach(async ({ page }) => {
  await page.addInitScript(installTauriMock, {
    openPath: DOC,
    files: {
      [DOC]: readFileSync("tests/fixtures/ribbon-rage.docx").toString("base64"),
      [DRAFT]: readFileSync("tests/fixtures/ribbon-rage-draft.docx").toString("base64"),
    },
    versions: [{ ts: 1_700_000_000_000, file: DRAFT }],
  });
  await page.addInitScript(() => localStorage.setItem("reopen", "false"));
  await page.goto("/");
  await page.keyboard.press("Control+O");
  await expect(doc(page)).toHaveAttribute("contenteditable", "true");
});

test("version history previews an old version and restores it", async ({ page }) => {
  await page.locator("#more").click();
  await page.locator('#menu [data-cmd="history"]').click();
  await page.locator(".hv-item").nth(1).click();
  await expect(page.locator("#version-bar")).toBeVisible();
  await expect(doc(page).locator("p").first()).toHaveText("WIBTA if I hid the ribbon on my dad's PC?");
  await expect(doc(page)).not.toHaveAttribute("contenteditable", "true"); // an old version is read-only

  await page.locator("#version-restore").click();
  await expect(page.locator("#version-bar")).toBeHidden();
  await expect(doc(page).locator("p").first()).toHaveText("WIBTA if I hid the ribbon on my dad's PC?");
  await expect(doc(page)).toHaveAttribute("contenteditable", "true");
  await expect(page.locator("#state")).toContainText("Edited");
});

test("switching tabs right after typing still auto-saves the edited file", async ({ page }) => {
  await paragraph(page, "Update: dad").click();
  await page.keyboard.type(" Edited.");
  await page.keyboard.press("Control+T"); // before the autosave timer fires
  await expect.poll(() => page.evaluate(() => window.__writes ?? [])).toEqual([{ path: DOC, size: expect.any(Number) }]);
  await page.waitForTimeout(2000); // the old timer must not save anything else later
  expect(await page.evaluate(() => window.__writes.length)).toBe(1);
});

test("a spelling suggestion replaces the word", async ({ page }) => {
  await paragraph(page, "Update: dad").click();
  await page.keyboard.press("Control+End"); // the last paragraph wraps, so plain End stops mid-paragraph
  await page.keyboard.press("Enter");
  await page.keyboard.type("Long live the Ribon.");
  const point = await paragraph(page, "Long live").evaluate((p) => {
    const n = [...p.querySelectorAll("span")].map((s) => s.firstChild).find((t) => t?.data?.includes("Ribon"));
    const r = document.createRange();
    r.setStart(n, n.data.indexOf("Ribon") + 2);
    r.collapse(true);
    const b = r.getBoundingClientRect();
    return { x: b.x, y: b.y + 4 };
  });
  await page.mouse.click(point.x, point.y, { button: "right" });
  await page.locator("#ctx-spell .suggestion", { hasText: "Ribbon" }).click();
  await expect(paragraph(page, "Long live")).toHaveText("Long live the Ribbon.");
});
