import { test, expect } from "@playwright/test";
import {
  openApp,
  newDocument,
  doc,
  paragraph,
  caretIn,
  select,
  menu,
  save,
  reopen,
  parseXml,
  all,
  val,
  text,
  bodyParagraphs,
} from "./helpers.js";

let problems;
test.beforeEach(async ({ page }) => {
  problems = await openApp(page);
});
test.afterEach(() => {
  expect(problems, "errors or blocked content in the page").toEqual([]);
});

test("typing with Markdown shortcuts saves real Word styles", async ({ page }) => {
  await newDocument(page);
  await page.keyboard.type("# Quarterly plan");
  await page.keyboard.press("Enter");
  await page.keyboard.type("We ship the **new editor** this month.");
  await page.keyboard.press("Enter");
  await page.keyboard.type("- write tests");
  await page.keyboard.press("Enter");
  await page.keyboard.type("polish the UI");
  await page.keyboard.press("Enter");
  await page.keyboard.press("Enter"); // an empty list item ends the list
  await page.keyboard.type("That's all.");

  const saved = await save(page);
  const d = parseXml(saved.xml);
  expect(bodyParagraphs(d)).toEqual([
    { text: "Quarterly plan", style: "Heading1" },
    { text: "We ship the new editor this month.", style: null },
    { text: "write tests", style: "ListBullet" },
    { text: "polish the UI", style: "ListBullet" },
    { text: "That's all.", style: null },
  ]);
  const boldRun = all(d, "r").find((r) => text(r) === "new editor");
  expect(boldRun && all(boldRun, "b").length).toBe(1);
  expect(saved.xml).not.toContain("​"); // the helper character used while typing never reaches the file

  await reopen(page, saved);
  await expect(paragraph(page, "Quarterly plan")).toHaveClass(/docx_heading1/);
  await expect(doc(page).locator("p", { hasText: "write tests" })).toHaveClass(/docx-num-/);
});

test("undo reverts a formatting change, redo brings it back", async ({ page }) => {
  await newDocument(page);
  await page.keyboard.type("Make this bold");
  await select(page, "Make this bold", "bold");
  await page.keyboard.press("Control+B");
  const weight = () =>
    doc(page)
      .locator("span", { hasText: /^bold$/ })
      .evaluate((s) => getComputedStyle(s).fontWeight);
  expect(await weight()).toBe("700");
  await page.keyboard.press("Control+Z");
  await expect(doc(page).locator("span", { hasText: /^bold$/ })).toHaveCount(0);
  await page.keyboard.press("Control+Y");
  expect(await weight()).toBe("700");
});

test("tables: inserted, grown with Tab and the right-click menu, saved as a valid Word table", async ({ page }) => {
  await newDocument(page);
  await page.keyboard.type("Before the table");
  await menu(page, "table");
  for (const cell of ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"]) {
    await page.keyboard.type(cell);
    await page.keyboard.press("Tab");
  }
  await page.keyboard.type("A4"); // Tab in the last cell added this row
  await paragraph(page, "B1").click({ button: "right" });
  await page.locator('[data-table="colRight"]').click();

  const table = doc(page).locator("table");
  await expect(table.locator("tr")).toHaveCount(4);
  await expect(table.locator("tr").first().locator("td")).toHaveCount(4);

  const d = parseXml((await save(page)).xml);
  const [tbl] = all(d, "tbl");
  const rows = all(tbl, "tr");
  expect(rows).toHaveLength(4);
  for (const row of rows) {
    const cells = all(row, "tc");
    expect(cells).toHaveLength(4);
    for (const cell of cells) expect(all(cell, "p").length, "every Word cell needs a paragraph").toBeGreaterThan(0);
  }
  expect(all(tbl, "gridCol")).toHaveLength(4);
  expect(rows.map((r) => all(r, "tc").map(text).join(","))).toEqual(["A1,B1,,C1", "A2,B2,,C2", "A3,B3,,C3", "A4,,,"]);
});

test("right-clicking a table cell offers table tools even when text elsewhere is selected", async ({ page }) => {
  await newDocument(page);
  await page.keyboard.type("Some text above");
  await menu(page, "table");
  await page.keyboard.type("cell");
  await select(page, "Some text above", "text"); // a selection outside the table
  await paragraph(page, "cell").click({ button: "right" });
  await expect(page.locator("#ctx-table")).toBeVisible();
  await expect(page.locator('[data-table="rowBelow"]')).toBeEnabled();
});

test("pasted control characters never reach the file", async ({ page }) => {
  await newDocument(page);
  await page.keyboard.type("From a PDF: ");
  await doc(page).evaluate((article) => {
    const data = new DataTransfer();
    data.setData("text/plain", "page one\fpage two\u0007 beep");
    article.dispatchEvent(new ClipboardEvent("paste", { clipboardData: data, bubbles: true, cancelable: true }));
  });
  await expect(paragraph(page, "page two")).toHaveText("page two beep");

  const saved = await save(page);
  // eslint-disable-next-line no-control-regex
  expect(saved.xml).not.toMatch(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/);
  expect(bodyParagraphs(parseXml(saved.xml)).map((p) => p.text)).toEqual(["From a PDF: page one", "page two beep"]);
});

test("a custom text color is saved to the file", async ({ page }) => {
  await newDocument(page);
  await page.keyboard.type("Paint this word");
  await select(page, "Paint this word", "word");
  await page.locator("#aa").click();
  const picker = page.locator("#color-custom");
  await picker.dispatchEvent("pointerdown"); // as a real click would, before the native picker opens
  await picker.fill("#12a4b6");
  await expect(picker.locator("..")).toHaveAttribute("aria-pressed", "true");

  const d = parseXml((await save(page)).xml);
  const run = all(d, "r").find((r) => text(r) === "word");
  expect(run && val(run, "color")).toBe("12A4B6");
});

test("find and replace changes every match in one undo step", async ({ page }) => {
  await newDocument(page, "Report");
  const count = async () => (await doc(page).innerText()).split("finding").length - 1;
  const before = await count();
  expect(before).toBeGreaterThan(1);

  await page.keyboard.press("Control+H");
  await page.locator("#find-input").fill("finding");
  await page.locator("#replace-input").fill("result");
  await page.locator("#replace-all").click();
  expect(await count()).toBe(0);

  await page.locator("#find-done").click();
  await caretIn(page, "Summary");
  await page.keyboard.press("Control+Z");
  expect(await count()).toBe(before);
});

test("links are saved as Word hyperlinks with their address", async ({ page }) => {
  await newDocument(page);
  await page.keyboard.type("Read the docs for details.");
  await select(page, "Read the docs", "the docs");
  await page.keyboard.press("Control+Shift+K");
  await page.locator("#link-input").fill("example.com/docs");
  await page.locator("#link-input").press("Enter");

  // Enter in the link box must not reach the document
  await expect(doc(page).locator("p")).toHaveCount(1);
  await expect(doc(page).locator("[data-href]")).toHaveText("the docs");

  const saved = await save(page);
  const d = parseXml(saved.xml);
  const [link] = all(d, "hyperlink");
  expect(text(link)).toBe("the docs");
  const id = link.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
  const rels = await saved.part("word/_rels/document.xml.rels");
  expect(rels).toMatch(
    new RegExp(
      `Id="${id}"[^>]*Target="https://example.com/docs"[^>]*TargetMode="External"|Target="https://example.com/docs"[^>]*Id="${id}"`,
    ),
  );
});

test("inserted images are stored inside the file", async ({ page }) => {
  await newDocument(page);
  await page.keyboard.type("A picture follows.");
  // A real 4x2 blue PNG
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAAEElEQVR4nGOwzn8ARwzIHACbAgxR1RL7/QAAAABJRU5ErkJggg==", "base64");
  await menu(page, "image");
  await page.locator("#image-file").setInputFiles({ name: "dot.png", mimeType: "image/png", buffer: png });
  await expect(doc(page).locator("img")).toHaveCount(1);

  const saved = await save(page);
  const media = Object.keys(saved.zip.files).filter((f) => f.startsWith("word/media/") && f.endsWith(".png"));
  expect(media).toHaveLength(1);
  expect(await saved.part("[Content_Types].xml")).toMatch(/Extension="png"/i);
  const d = parseXml(saved.xml);
  const blip = d.getElementsByTagName("a:blip")[0];
  expect(blip?.getAttribute("r:embed")).toBeTruthy();
  expect(await saved.part("word/_rels/document.xml.rels")).toContain(media[0].replace("word/", ""));
});

test("line spacing and list levels are written to the paragraph", async ({ page }) => {
  await newDocument(page);
  await page.keyboard.type("- first");
  await page.keyboard.press("Enter");
  await page.keyboard.type("nested");
  await caretIn(page, "nested", true);
  await page.keyboard.press("Tab");
  await expect(paragraph(page, "nested")).toHaveClass(/docx-num-\d+-1/);

  await caretIn(page, "first");
  await page.locator("#aa").click();
  await page.locator('#format-pop [data-line="1.5"]').click();

  const d = parseXml((await save(page)).xml);
  const ps = all(d, "p");
  const nested = ps.find((p) => text(p) === "nested");
  expect(val(nested, "ilvl")).toBe("1");
  const first = ps.find((p) => text(p) === "first");
  expect(all(first, "spacing")[0].getAttributeNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "line")).toBe("360");
});

test("Home shows the start screen without closing the document", async ({ page }) => {
  await newDocument(page, "Report");
  await page.locator("#home-btn").click();
  await expect(page.locator(".open-btn")).toBeVisible();
  await expect(page.locator("#home-btn")).toBeHidden();
  await expect(page.locator(".tab")).toHaveCount(2);

  await page.locator(".tab", { hasText: "Report" }).click();
  await expect(paragraph(page, "Findings")).toBeVisible();
  await expect(page.locator(".tab")).toHaveCount(1); // the start screen's tab goes away once left
});

test("tabs keep documents apart", async ({ page }) => {
  await newDocument(page, "Report");
  await page.keyboard.press("Control+T");
  await newDocument(page, "Meeting Notes");
  await expect(page.locator(".tab")).toHaveCount(2);
  await expect(page.locator("#name")).toHaveText("Meeting Notes");

  await page.keyboard.press("Control+F");
  await page.locator("#find-input").fill("Findings"); // only in the Report tab
  await expect(page.locator("#find-count")).toHaveText("No matches");
  await page.locator("#find-done").click();

  await page.locator(".tab", { hasText: "Report" }).click();
  await expect(page.locator("#name")).toHaveText("Report");
  await expect(paragraph(page, "Findings")).toBeVisible();
});
