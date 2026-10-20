// Takes the README tour screenshots from the built app: `npm run build`, `npx vite preview`, then `node scripts/screenshots.mjs`.
// The desktop side is mocked (tests/tauri-mock.js), so file dialogs, version history and spelling show sample data.
import { chromium } from "@playwright/test";
import { readFileSync, mkdirSync } from "node:fs";
import { installTauriMock } from "../tests/tauri-mock.js";

const URL = "http://localhost:4173/";
const DOC = "C:\\Users\\you\\Documents\\AITA for uninstalling Word.docx";
const DRAFT = "C:\\Users\\you\\Documents\\WIBTA if I hid the ribbon.docx";
const now = Date.now();
const IMAGE = "C:\\Users\\you\\Pictures\\ribbonless.png";
const mock = {
  openPath: DOC,
  imagePath: IMAGE,
  files: {
    [DOC]: readFileSync("tests/fixtures/ribbon-rage.docx").toString("base64"),
    [DRAFT]: readFileSync("tests/fixtures/ribbon-rage-draft.docx").toString("base64"),
    [IMAGE]: readFileSync("assets/icon.png").toString("base64"),
  },
  versions: [3, 26, 140, 60 * 26, 60 * 75].map((min, i) => ({ ts: now - min * 60_000, file: i < 2 ? DOC : DRAFT })).reverse(),
};

mkdirSync("docs/screens", { recursive: true });
const browser = await chromium.launch();

// A fresh window: theme, outline state, then the sample document opened through the (mocked) Open dialog
async function start({ theme = "light", outline = false, open = true } = {}) {
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: 1.5 });
  await page.addInitScript(installTauriMock, mock);
  await page.addInitScript(
    ([theme, outline]) => {
      if (localStorage.getItem("seeded")) return;
      localStorage.setItem("seeded", "1");
      localStorage.setItem("theme", JSON.stringify(theme));
      localStorage.setItem("outline", JSON.stringify(outline));
      localStorage.setItem("reopen", "false");
      localStorage.setItem("smoothCaret", "false");
    },
    [theme, outline],
  );
  await page.goto(URL);
  await page.locator(".open-btn").waitFor();
  if (open) {
    await page.keyboard.press("Control+O");
    await page.locator(".tab-pane:not([hidden]) article[contenteditable=true]").waitFor();
  }
  await page.waitForTimeout(400);
  return page;
}

const shot = async (page, name) => {
  await page.waitForTimeout(450);
  await page.screenshot({ path: `docs/screens/${name}.png` });
  console.log("shot", name);
};
const para = (page, text) => page.locator(".tab-pane:not([hidden]) article p", { hasText: text }).first();
async function selectWords(page, text, part) {
  await para(page, text).evaluate((p, part) => {
    const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n = walk.nextNode(); n; n = walk.nextNode()) {
      const i = n.data.indexOf(part);
      if (i < 0) continue;
      const r = document.createRange();
      r.setStart(n, i);
      r.setEnd(n, i + part.length);
      p.closest("article").focus();
      getSelection().removeAllRanges();
      getSelection().addRange(r);
      return;
    }
  }, part);
  await page.mouse.move(5, 400);
  await page.mouse.down();
  await page.mouse.up(); // lets the selection bar appear, as after a real drag
  await para(page, text).evaluate((p, part) => {
    const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let n = walk.nextNode(); n; n = walk.nextNode()) {
      const i = n.data.indexOf(part);
      if (i < 0) continue;
      const r = document.createRange();
      r.setStart(n, i);
      r.setEnd(n, i + part.length);
      getSelection().removeAllRanges();
      getSelection().addRange(r);
      return;
    }
  }, part);
}
const caretAtEnd = (page, text) =>
  para(page, text).evaluate((p) => {
    const r = document.createRange();
    r.selectNodeContents(p);
    r.collapse(false);
    p.closest("article").focus();
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  });

// 1. The main window with the outline
let page = await start({ outline: true });
await shot(page, "hero");

// 2. Start screen: templates and recent documents (after the sample has been opened once)
await page.locator("#more").click();
await page.locator('#menu [data-cmd="home"]').click();
await shot(page, "start");
await page.close();

// 3. Dark appearance with the format panel
page = await start({ theme: "dark" });
await para(page, "Throwaway because").click();
await page.locator("#aa").click();
await shot(page, "format-panel");
await page.close();

page = await start();

// 4. Selection bar
await selectWords(page, "Deleted one space", "The space was load-bearing.");
await shot(page, "selection-bar");

// 5. Command bar
await page.keyboard.press("Control+K");
await page.locator("#palette-input").type("ins");
await shot(page, "command-bar");
await page.keyboard.press("Escape");

// 6. Slash menu on a new line
await caretAtEnd(page, "Throwaway because");
await page.keyboard.press("Enter");
await page.keyboard.type("/");
await shot(page, "slash-menu");
await page.keyboard.press("Escape");
await page.keyboard.press("Control+Z");
await page.keyboard.press("Control+Z");

// 7. Find and replace
await page.keyboard.press("Control+H");
await page.locator("#find-input").fill("ribbon");
await page.locator("#replace-input").fill("ribbon (affectionate)");
await shot(page, "find-replace");
await page.locator("#find-done").click();

// 8. Table tools on right-click
const cell = para(page, "ESH");
await cell.scrollIntoViewIfNeeded();
await cell.click({ button: "right" });
await shot(page, "table-menu");
await page.keyboard.press("Escape");

// 9. Links
await para(page, "Update: dad").evaluate((p) => p.scrollIntoView({ block: "center" }));
await selectWords(page, "Update: dad", "Ribbonless");
await page.keyboard.press("Control+Shift+K");
await page.locator("#link-input").fill("github.com/stefanrotari99i/ribbonless");
await shot(page, "link");
await page.keyboard.press("Escape");

// 10. Spelling suggestions
await caretAtEnd(page, "Update: dad");
await page.keyboard.press("Enter");
await page.keyboard.type("Long live the Ribon.");
const typo = para(page, "Long live the Ribon");
// near the top, so the menu opens below the word instead of over it
await typo.evaluate((p) => {
  p.scrollIntoView({ block: "start" });
  let s = p.parentElement;
  while (s && s.scrollHeight <= s.clientHeight) s = s.parentElement;
  if (s) s.scrollTop -= 120;
});
await page.waitForTimeout(300);
const box = await typo.evaluate((p) => {
  const n = [...p.querySelectorAll("span")].map((s) => s.firstChild).find((t) => t?.data?.includes("Ribon"));
  const r = document.createRange();
  const i = n.data.indexOf("Ribon");
  r.setStart(n, i + 1);
  r.setEnd(n, i + 2);
  const b = r.getBoundingClientRect();
  return { x: b.x, y: b.y + b.height / 2 };
});
await page.mouse.click(box.x, box.y, { button: "right" });
await shot(page, "spelling");
await page.keyboard.press("Escape");

// 11. Changes since opening: the margin marks
await caretAtEnd(page, "Throwaway because");
await page.keyboard.type(" Then he asked Word to undo. It undid the tea.");
await caretAtEnd(page, "Your dad was being held");
await page.keyboard.type(" Dad has since been freed.");
await para(page, "Throwaway because").evaluate((p) => p.scrollIntoView({ block: "center" }));
await shot(page, "changes");

// 12. Version history with an old version previewed
await page.locator("#more").click();
await page.locator('#menu [data-cmd="history"]').click();
await page.locator(".hv-item").nth(4).click();
await page.locator("#version-bar").waitFor();
await shot(page, "version-history");
await page.locator("#version-done").click();

// 13. Image resize handles
await caretAtEnd(page, "Throwaway because");
await page.locator("#more").click();
await page.locator('#menu [data-cmd="image"]').click();
await page.locator(".tab-pane:not([hidden]) article img").first().click();
await shot(page, "image");
await page.keyboard.press("Escape");
await page.close();

// 14. Focus mode
page = await start();
await para(page, "Throwaway because").click();
await page.keyboard.press("Control+Shift+F");
await page.mouse.move(640, 500);
await shot(page, "focus-mode");
await page.close();

// 15. Tabs
page = await start();
await page.keyboard.press("Control+T");
await page.locator(".card", { hasText: "Meeting Notes" }).click();
await page.locator(".tab-pane:not([hidden]) article[contenteditable=true]").waitFor();
await shot(page, "tabs");
await page.close();

await browser.close();
