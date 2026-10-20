// Builds tests/fixtures/ribbon-rage.docx: the sample document used for screenshots and tests.
// A made-up Reddit-style thread; every username is invented.
import { writeFileSync, mkdirSync } from "node:fs";
import JSZip from "jszip";
import { blankDocx } from "../blank.js";

const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// A run: plain string, or { t, b, i, hl, color }
function run(r) {
  if (typeof r === "string") r = { t: r };
  const props = [r.b && "<w:b/>", r.i && "<w:i/>", r.color && `<w:color w:val="${r.color}"/>`, r.hl && `<w:highlight w:val="${r.hl}"/>`]
    .filter(Boolean)
    .join("");
  return `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ""}<w:t xml:space="preserve">${esc(r.t)}</w:t></w:r>`;
}

const p = (style, ...runs) => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}${runs.map(run).join("")}</w:p>`;

function table(rows) {
  const border = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="C9CED8"/>`;
  const cell = (text, head) =>
    `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${head ? '<w:shd w:val="clear" w:color="auto" w:fill="EEF1F8"/>' : ""}</w:tcPr>` +
    `<w:p>${run(head ? { t: text, b: true } : text)}</w:p></w:tc>`;
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map(border).join("")}</w:tblBorders>` +
    `<w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
    rows.map((r, i) => `<w:tr>${r.map((c) => cell(c, i === 0)).join("")}</w:tr>`).join("") +
    `</w:tbl>`
  );
}

const comment = (user, votes, text) =>
  p("ListBullet", { t: `${user}`, b: true }, { t: `  ·  ${votes} upvotes  —  `, color: "86868B" }, text);

const body = [
  p("Title", "AITA for uninstalling Word from my dad's PC?"),
  p("", {
    t: "r/AmItheAsshole  ·  Posted by u/throwaway_ribbon_61 · 9 hours ago  ·  48.9k upvotes  ·  flair: Not the A-hole",
    i: true,
    color: "86868B",
  }),

  p("Heading1", "Original post"),
  p(
    "",
    "Throwaway because my dad knows my main. My dad (61M) spent ",
    { t: "forty-five minutes", b: true },
    " trying to un-bold a single comma. Word responded by bolding the entire document, adding a page break before his signature and asking if he wanted help writing a letter. He did not. He wanted the comma back. So while he was making tea, I uninstalled Word. AITA?",
  ),
  p("", { t: "Edit: wow, this blew up. RIP my inbox.", i: true }),
  p("", { t: "Edit 2: Thanks for the gold, kind stranger!", i: true }),
  p("", { t: "Edit 3: He found Word in the Recycle Bin and restored it. Please send help.", i: true }),

  p("Heading1", "Top comments"),
  comment("u/JudgementDay_", "31.2k", "NTA. Your dad was being held hostage by the Paste Options button. You were the rescue mission."),
  comment("u/InstructionsUnclear", "22.4k", "Instructions unclear, clicked Insert SmartArt. My CV is now a pyramid scheme."),
  comment("u/LoadBearingSpace", "17.9k", "Deleted one space. The table moved to page 9. The space was load-bearing."),
  comment("u/CtrlS_Every4Seconds", "12.6k", "Save every four seconds. Trust no one. This is the way."),
  comment("u/ClippyWidow", "9.8k", "“It looks like you're writing a divorce letter. Would you like help?” Yes, Clippy. With you."),
  comment("u/UsernameChecksOut_", "4.1k", "“Document1 (Recovered) (2) FINAL final USE THIS ONE.docx”. I have nothing to add."),

  p("Heading1", "Verdict tally"),
  table([
    ["Verdict", "Votes", "Top reason"],
    ["NTA", "18,204", "The ribbon had it coming"],
    ["YTA", "3", "All three are the ribbon's alt accounts"],
    ["ESH", "412", "Especially Paste Options"],
    ["INFO", "77", "Did he try turning it off and on again?"],
    ["Ribbon", "1", "Voted for itself. Still bolded."],
  ]),

  p("Heading1", "TL;DR"),
  p("ListNumber", "Dad wanted to un-bold one comma."),
  p("ListNumber", "Word bolded his entire retirement speech instead."),
  p("ListNumber", "I uninstalled Word while the kettle was on."),
  p("ListNumber", "Dad now writes in Ribbonless and says it “feels illegal”."),

  p("Heading2", "Final update"),
  p(
    "",
    "Update: dad switched to Ribbonless. He still hovers where the ribbon used to be, like a phantom limb. Mods locked the thread for being ",
    { t: "too wholesome for this sub", hl: "yellow" },
    ".",
  ),
].join("");

// An earlier, calmer draft, shown when screenshots preview an old version
const draft = [
  p("Title", "WIBTA if I hid the ribbon on my dad's PC?"),
  p("", {
    t: "r/AmItheAsshole  ·  Posted by u/throwaway_ribbon_61 · 2 days ago  ·  12 upvotes  ·  flair: Asshole Pending",
    i: true,
    color: "86868B",
  }),
  p("Heading1", "Original post"),
  p("", "Small question, probably nothing. My dad's Word has quite a lot of buttons. Would it be wrong to collapse them? Just wondering."),
  p("Heading1", "Top comments"),
  comment("u/JudgementDay_", "2", "Need more info. How many buttons are we talking."),
].join("");

async function write(file, content) {
  const zip = await JSZip.loadAsync(await blankDocx());
  const doc = (await zip.file("word/document.xml").async("string")).replace(/<w:body>.*?(<w:sectPr)/s, `<w:body>${content}$1`);
  zip.file("word/document.xml", doc);
  writeFileSync(file, await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
  console.log("wrote", file);
}

mkdirSync("tests/fixtures", { recursive: true });
await write("tests/fixtures/ribbon-rage.docx", body);
await write("tests/fixtures/ribbon-rage-draft.docx", draft);
