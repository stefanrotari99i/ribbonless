import { expect } from "@playwright/test";
import { DOMParser } from "@xmldom/xmldom";
import JSZip from "jszip";

// Opens the app and fails the test on any error or content security policy violation
export async function openApp(page) {
  const problems = [];
  page.on("pageerror", (e) => problems.push(e.message));
  page.on("console", (m) => {
    if (m.type() === "error" || /Content Security Policy/i.test(m.text())) problems.push(m.text());
  });
  await page.goto("/");
  await expect(page.locator(".new-btn")).toBeVisible();
  return problems;
}

export async function newDocument(page, template) {
  const card = template ? page.locator(".card", { hasText: template }) : page.locator(".new-btn");
  await card.first().click();
  await expect(page.locator(".tab-pane:not([hidden]) article[contenteditable=true]")).toBeVisible();
}

export const doc = (page) => page.locator(".tab-pane:not([hidden]) article");
export const paragraph = (page, text) => doc(page).locator("p", { hasText: text }).first();

// Puts the caret at the end (or start) of the paragraph containing `text`
export async function caretIn(page, text, atStart = false) {
  await paragraph(page, text).evaluate((p, atStart) => {
    const r = document.createRange();
    r.selectNodeContents(p);
    r.collapse(atStart);
    p.closest("article").focus();
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  }, atStart);
}

// Selects `part` inside the paragraph containing `text`
export async function select(page, text, part) {
  await paragraph(page, text).evaluate((p, part) => {
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
    throw new Error(`"${part}" not found`);
  }, part);
}

export async function menu(page, command) {
  await page.locator("#more").click();
  await page.locator(`#menu [data-cmd="${command}"]`).click();
}

// Saves with Ctrl+S (a download in the browser) and returns the file's parts
export async function save(page) {
  const download = page.waitForEvent("download");
  await page.keyboard.press("Control+S");
  const file = await download;
  const bytes = await (await file.createReadStream()).toArray();
  const buffer = Buffer.concat(bytes);
  const zip = await JSZip.loadAsync(buffer);
  const part = async (name) => (await zip.file(name)?.async("string")) ?? null;
  return { buffer, name: file.suggestedFilename(), zip, xml: await part("word/document.xml"), part };
}

// Parses a saved part; throws on XML that Word would refuse to open
export function parseXml(xml) {
  const errors = [];
  const d = new DOMParser({ onError: (level, msg) => level !== "warning" && errors.push(msg) }).parseFromString(xml, "application/xml");
  if (errors.length) throw new Error(`invalid XML: ${errors[0]}`);
  return d;
}

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const all = (node, tag) => Array.from(node.getElementsByTagNameNS(W, tag));
export const val = (node, tag) => all(node, tag)[0]?.getAttributeNS(W, "val") ?? null;
export const text = (node) =>
  all(node, "t")
    .map((t) => t.textContent)
    .join("");
// The body's paragraphs (not those inside tables), as { text, style }
export const bodyParagraphs = (d) =>
  all(d, "p")
    .filter((p) => p.parentNode.localName === "body")
    .map((p) => ({ text: text(p), style: val(p, "pStyle") }));

export async function reopen(page, saved) {
  await page.locator("#file").setInputFiles({ name: saved.name, mimeType: "application/octet-stream", buffer: saved.buffer });
  await expect(doc(page)).toBeVisible();
}
