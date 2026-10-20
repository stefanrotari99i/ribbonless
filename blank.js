import JSZip from "jszip";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml";

const heading = (id, name, before, after, size, level) =>
  `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>` +
  `<w:pPr><w:keepNext/><w:spacing w:before="${before}" w:after="${after}"/><w:outlineLvl w:val="${level}"/></w:pPr>` +
  `<w:rPr><w:b/><w:color w:val="1D1D1F"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;

const FILES = {
  "[Content_Types].xml":
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="${CT}.styles+xml"/>` +
    `<Override PartName="/word/numbering.xml" ContentType="${CT}.numbering+xml"/></Types>`,
  "_rels/.rels":
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  "word/_rels/document.xml.rels":
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>` +
    `<Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/></Relationships>`,
  "word/styles.xml":
    `<w:styles ${W}><w:docDefaults>` +
    `<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:color w:val="1D1D1F"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/>` +
    `<w:pPr><w:spacing w:after="240" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:b/><w:color w:val="1D1D1F"/><w:sz w:val="52"/><w:szCs w:val="52"/></w:rPr></w:style>` +
    heading("Heading1", "heading 1", 360, 120, 32, 0) +
    heading("Heading2", "heading 2", 280, 80, 26, 1) +
    heading("Heading3", "heading 3", 240, 60, 24, 2) +
    `<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/>` +
    `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:after="60"/></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/>` +
    `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:spacing w:after="60"/></w:pPr></w:style></w:styles>`,
  "word/numbering.xml":
    `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>` +
    `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>` +
    `</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>` +
    `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>` +
    `</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`,
};

// Block kinds: T title, h heading 1, h2 heading 2, p body, l bullet, n numbered
const STYLE_OF = { T: "Title", h: "Heading1", h2: "Heading2", l: "ListBullet", n: "ListNumber" };
const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function paragraph({ k, text = "" }) {
  const style = STYLE_OF[k] ? `<w:pPr><w:pStyle w:val="${STYLE_OF[k]}"/></w:pPr>` : "";
  const run = text ? `<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r>` : "";
  return `<w:p>${style}${run}</w:p>`;
}

export const TEMPLATES = [
  {
    id: "letter",
    name: "Letter",
    blocks: [
      { k: "p", text: "Your Name" },
      { k: "p", text: "Street 1, 1234 AB City" },
      { k: "p", text: "" },
      { k: "p", text: new Date().toLocaleDateString([], { day: "numeric", month: "long", year: "numeric" }) },
      { k: "p", text: "" },
      { k: "p", text: "Recipient Name" },
      { k: "p", text: "Company" },
      { k: "p", text: "" },
      { k: "p", text: "Dear Recipient," },
      { k: "p", text: "Start your letter here. Keep the first paragraph short and say why you are writing." },
      { k: "p", text: "Use the second paragraph for the details, and close with what you would like to happen next." },
      { k: "p", text: "Kind regards," },
      { k: "p", text: "Your Name" },
    ],
  },
  {
    id: "report",
    name: "Report",
    blocks: [
      { k: "T", text: "Report Title" },
      { k: "p", text: "Prepared by Your Name · " + new Date().toLocaleDateString([], { month: "long", year: "numeric" }) },
      { k: "h", text: "Summary" },
      { k: "p", text: "The two or three sentences a busy reader needs: what you looked at, what you found, what you recommend." },
      { k: "h", text: "Findings" },
      { k: "l", text: "First finding, with the number that supports it" },
      { k: "l", text: "Second finding" },
      { k: "l", text: "Third finding" },
      { k: "h", text: "Recommendations" },
      { k: "n", text: "Most important action" },
      { k: "n", text: "Next action" },
      { k: "h", text: "Details" },
      { k: "p", text: "Background, method and supporting material." },
    ],
  },
  {
    id: "meeting",
    name: "Meeting Notes",
    blocks: [
      { k: "T", text: "Meeting Notes" },
      { k: "p", text: new Date().toLocaleDateString([], { weekday: "long", day: "numeric", month: "long" }) + " · Attendees: " },
      { k: "h", text: "Agenda" },
      { k: "n", text: "First topic" },
      { k: "n", text: "Second topic" },
      { k: "h", text: "Notes" },
      { k: "p", text: "" },
      { k: "h", text: "Decisions" },
      { k: "l", text: "" },
      { k: "h", text: "Action Items" },
      { k: "l", text: "Owner — task — due date" },
    ],
  },
  {
    id: "design",
    name: "Design Doc",
    blocks: [
      { k: "T", text: "Project Name" },
      {
        k: "p",
        text: "Author · Status: Draft · " + new Date().toLocaleDateString([], { day: "numeric", month: "short", year: "numeric" }),
      },
      { k: "h", text: "Overview" },
      { k: "p", text: "What this is and why it matters, in one paragraph." },
      { k: "h", text: "Goals" },
      { k: "l", text: "Goal one" },
      { k: "l", text: "Goal two" },
      { k: "h", text: "Non-goals" },
      { k: "l", text: "What this deliberately does not cover" },
      { k: "h", text: "Design" },
      { k: "h2", text: "Approach" },
      { k: "p", text: "How it works." },
      { k: "h", text: "Open Questions" },
      { k: "n", text: "" },
    ],
  },
];

export function blankDocx(blocks = [{ k: "p" }]) {
  const zip = new JSZip();
  for (const [path, body] of Object.entries(FILES)) zip.file(path, HEAD + body);
  zip.file(
    "word/document.xml",
    HEAD +
      `<w:document ${W}><w:body>${blocks.map(paragraph).join("")}` +
      `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>` +
      `</w:body></w:document>`,
  );
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
