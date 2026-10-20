import JSZip from "jszip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const STYLE_NAMES = [
  ["title", "Title"],
  ["heading 1", "Heading 1"],
  ["heading 2", "Heading 2"],
  ["heading 3", "Heading 3"],
  ["list bullet", "Bulleted List"],
  ["list number", "Numbered List"],
];
// Paragraphs holding these are never merged away by Backspace, so images and section breaks survive
const KEEP = ["drawing", "pict", "object", "sectPr", "fldChar", "fldSimple", "AlternateContent", "txbxContent"];
const RPR_ORDER = [
  "rStyle",
  "rFonts",
  "b",
  "bCs",
  "i",
  "iCs",
  "caps",
  "smallCaps",
  "strike",
  "dstrike",
  "outline",
  "shadow",
  "emboss",
  "imprint",
  "noProof",
  "snapToGrid",
  "vanish",
  "webHidden",
  "color",
  "spacing",
  "w",
  "kern",
  "position",
  "sz",
  "szCs",
  "highlight",
  "u",
  "effect",
  "bdr",
  "shd",
  "fitText",
  "vertAlign",
  "rtl",
  "cs",
  "em",
  "lang",
  "eastAsianLayout",
  "specVanish",
  "oMath",
];
const PPR_ORDER = [
  "pStyle",
  "keepNext",
  "keepLines",
  "pageBreakBefore",
  "framePr",
  "widowControl",
  "numPr",
  "suppressLineNumbers",
  "pBdr",
  "shd",
  "tabs",
  "suppressAutoHyphens",
  "kinsoku",
  "wordWrap",
  "overflowPunct",
  "topLinePunct",
  "autoSpaceDE",
  "autoSpaceDN",
  "bidi",
  "adjustRightInd",
  "snapToGrid",
  "spacing",
  "ind",
  "contextualSpacing",
  "mirrorIndents",
  "suppressOverlap",
  "jc",
  "textDirection",
  "textAlignment",
  "textboxTightWrap",
  "outlineLvl",
  "divId",
  "cnfStyle",
  "rPr",
  "sectPr",
  "pPrChange",
];
const ZWSP = "\u200B";
// Characters XML 1.0 forbids; one of them makes Word reject the whole file
// eslint-disable-next-line no-control-regex
const NOT_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MIME = { png: "image/png", jpeg: "image/jpeg", jpg: "image/jpeg", gif: "image/gif" };
const RUN_PARTS = new Set(["rPr", "t", "tab", "br", "lastRenderedPageBreak"]);
const JC = { left: null, start: null, center: "center", right: "right", end: "right", justify: "both" };
const FORMATS = {
  bold: ["b", "fontWeight", "bold", "normal"],
  italic: ["i", "fontStyle", "italic", "normal"],
  underline: ["u", "textDecorationLine", "underline", "none"],
};

const squash = (s) => s.replace(/\s+/g, "");
const xmlText = (p) => [...p.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join("");
const parseXml = (s) => new DOMParser().parseFromString(s, "application/xml");
const escapeClass = (s) => s.replace(/[ .]+/g, "-").replace(/&+/g, "and").toLowerCase();
const isTextRun = (r) => r.getElementsByTagNameNS(W, "t").length > 0 && [...r.children].every((c) => RUN_PARTS.has(c.localName));
const childEl = (parent, name) => [...parent.children].find((c) => c.localName === name);
const topChild = (parent, node) => {
  while (node.parentNode !== parent) node = node.parentNode;
  return node;
};
const pt = (px) => Math.round(parseFloat(px) * 1.5) / 2; // CSS px to points, to the nearest half point
const elementOf = (node) => (node?.nodeType === 1 ? node : node?.parentElement);

export const editableP = (node) => {
  const p = elementOf(node)?.closest("article p");
  return p?.isContentEditable ? p : null;
};

// Inserts `el` into a w:rPr / w:pPr at the position the schema requires
function insertOrdered(parent, el, order) {
  const rank = order.indexOf(el.localName);
  const before = [...parent.children].find((c) => order.indexOf(c.localName) > rank) ?? null;
  parent.insertBefore(el, before);
}

function serialize(xml) {
  const body = new XMLSerializer().serializeToString(xml);
  return body.startsWith("<?xml") ? body : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body;
}

// Word's highlight palette; anything else is written as shading
const HIGHLIGHTS = {
  yellow: "FFFF00",
  green: "00FF00",
  cyan: "00FFFF",
  magenta: "FF00FF",
  blue: "0000FF",
  red: "FF0000",
  darkBlue: "00008B",
  darkCyan: "008B8B",
  darkGreen: "006400",
  darkMagenta: "8B008B",
  darkRed: "8B0000",
  darkYellow: "808000",
  darkGray: "A9A9A9",
  lightGray: "D3D3D3",
  black: "000000",
};
const hex = (rgb) => {
  const m = rgb.match(/\d+(\.\d+)?/g);
  if (!m || (m.length > 3 && +m[3] === 0)) return null;
  return m
    .slice(0, 3)
    .map((v) => Math.round(+v).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
};

function formatOf(n, p) {
  const s = getComputedStyle(n.parentElement);
  let u = false;
  let hl = null;
  for (let a = n.parentElement; a && a !== p; a = a.parentElement) {
    const as = getComputedStyle(a);
    if (as.textDecorationLine.includes("underline")) u = true;
    hl ??= hex(as.backgroundColor);
  }
  if (getComputedStyle(p).textDecorationLine.includes("underline")) u = true;
  const link = n.parentElement.closest("a[href], [data-href]");
  const href = link && p.contains(link) ? (link.dataset.href ?? link.getAttribute("href")) : null;
  const font = s.fontFamily
    .split(",")[0]
    .trim()
    .replace(/^["']|["']$/g, "");
  return { b: parseInt(s.fontWeight) >= 600, i: s.fontStyle === "italic", u, size: pt(s.fontSize), color: hex(s.color), hl, font, href };
}

const sameFormat = (a, b) =>
  a.b === b.b &&
  a.i === b.i &&
  a.u === b.u &&
  a.size === b.size &&
  a.color === b.color &&
  a.hl === b.hl &&
  a.font === b.font &&
  a.href === b.href;

// How plain text looks in this paragraph: formatting equal to this is left to the style
function styleDefault(p) {
  const probe = document.createElement("span");
  const text = probe.appendChild(document.createTextNode("x"));
  p.append(probe);
  const f = formatOf(text, p);
  probe.remove();
  return f;
}

// Splits a paragraph into stretches of equal formatting, as shown on screen
function segments(p) {
  const out = [];
  const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  for (let n = walk.nextNode(); n; n = walk.nextNode()) {
    if (n.nodeName === "BR") {
      out.push({ br: true });
      continue;
    }
    const data = n.nodeType === Node.TEXT_NODE ? n.data.replaceAll(ZWSP, "") : "";
    if (!data) continue;
    const f = formatOf(n, p);
    const last = out.at(-1);
    if (last && !last.br && sameFormat(last, f)) last.text += data;
    else out.push({ text: data, ...f });
  }
  while (out.at(-1)?.br) out.pop(); // the browser keeps a placeholder <br> in emptied paragraphs
  return out;
}

function makeRun(xml, base, seg, def) {
  const el = (tag, val) => {
    const x = xml.createElementNS(W, `w:${tag}`);
    if (val !== undefined) x.setAttributeNS(W, "w:val", val);
    return x;
  };
  const r = el("r");
  if (seg.br) {
    r.appendChild(el("br"));
    return r;
  }
  const rPr = base ? base.cloneNode(true) : el("rPr");
  for (const tag of ["b", "bCs", "i", "iCs", "u", "sz", "szCs", "color", "highlight", "shd"])
    [...rPr.getElementsByTagNameNS(W, tag)].forEach((x) => x.remove());
  if (seg.font !== def.font) {
    [...rPr.getElementsByTagNameNS(W, "rFonts")].forEach((x) => x.remove());
    const f = el("rFonts");
    for (const a of ["ascii", "hAnsi", "cs", "eastAsia"]) f.setAttributeNS(W, `w:${a}`, seg.font);
    insertOrdered(rPr, f, RPR_ORDER);
  }
  if (seg.color && seg.color !== def.color) insertOrdered(rPr, el("color", seg.color), RPR_ORDER);
  if (seg.hl && seg.hl !== def.hl) {
    const name = Object.keys(HIGHLIGHTS).find((k) => HIGHLIGHTS[k] === seg.hl);
    if (name) insertOrdered(rPr, el("highlight", name), RPR_ORDER);
    else {
      const shd = el("shd", "clear");
      shd.setAttributeNS(W, "w:color", "auto");
      shd.setAttributeNS(W, "w:fill", seg.hl);
      insertOrdered(rPr, shd, RPR_ORDER);
    }
  }
  if (seg.b !== def.b) insertOrdered(rPr, el("b", seg.b ? undefined : "0"), RPR_ORDER);
  if (seg.i !== def.i) insertOrdered(rPr, el("i", seg.i ? undefined : "0"), RPR_ORDER);
  if (seg.size !== def.size) {
    insertOrdered(rPr, el("sz", String(seg.size * 2)), RPR_ORDER);
    insertOrdered(rPr, el("szCs", String(seg.size * 2)), RPR_ORDER);
  }
  if (seg.u !== def.u) insertOrdered(rPr, el("u", seg.u ? "single" : "none"), RPR_ORDER);
  if (rPr.children.length) r.appendChild(rPr);
  seg.text.split("\u2003").forEach((part, k) => {
    if (k) r.appendChild(el("tab"));
    if (!part) return;
    const t = r.appendChild(el("t"));
    t.setAttributeNS(XML_NS, "xml:space", "preserve");
    t.textContent = part.replace(NOT_XML, "");
  });
  return r;
}

// Rebuilds only the text runs of a paragraph (images and other content stay put); linkId maps a URL to its relationship
function rewriteParagraph(xml, xp, p, src, linkId) {
  const textRuns = [...xp.getElementsByTagNameNS(W, "r")].filter(isTextRun);
  const srcRun = src && [...src.getElementsByTagNameNS(W, "r")].find(isTextRun);
  const runRpr = (textRuns[0] ?? srcRun) && childEl(textRuns[0] ?? srcRun, "rPr");
  const base = runRpr ?? null;
  const anchor = textRuns[0] ? topChild(xp, textRuns[0]) : null;
  const def = styleDefault(p);
  let link = null; // the w:hyperlink being filled, and its href
  for (const seg of segments(p)) {
    const run = makeRun(xml, base, seg, def);
    if (!seg.href) {
      link = null;
      xp.insertBefore(run, anchor);
      continue;
    }
    if (link?.href !== seg.href) {
      const h = xml.createElementNS(W, "w:hyperlink");
      if (seg.href.startsWith("#")) h.setAttributeNS(W, "w:anchor", seg.href.slice(1));
      else h.setAttributeNS(DOC_REL, "r:id", linkId(seg.href));
      xp.insertBefore(h, anchor);
      link = { el: h, href: seg.href };
    }
    link.el.appendChild(run);
  }
  textRuns.forEach((r) => r.remove());
  [...xp.children]
    .filter((c) => (c.localName === "hyperlink" || c.localName === "smartTag") && !c.getElementsByTagNameNS(W, "r").length)
    .forEach((c) => c.remove());
}

function paragraphProps(xml, xp) {
  return childEl(xp, "pPr") ?? xp.insertBefore(xml.createElementNS(W, "w:pPr"), xp.firstChild);
}

function setParagraphStyle(xml, xp, id) {
  const pPr = paragraphProps(xml, xp);
  [...pPr.children].filter((c) => c.localName === "pStyle" || c.localName === "numPr").forEach((c) => c.remove());
  if (id) {
    const s = xml.createElementNS(W, "w:pStyle");
    s.setAttributeNS(W, "w:val", id);
    insertOrdered(pPr, s, PPR_ORDER);
  }
  if (!pPr.children.length) pPr.remove();
}

function setAlignment(xml, xp, align) {
  const pPr = paragraphProps(xml, xp);
  [...pPr.children].filter((c) => c.localName === "jc").forEach((c) => c.remove());
  const val = JC[align];
  if (val) {
    const jc = xml.createElementNS(W, "w:jc");
    jc.setAttributeNS(W, "w:val", val);
    insertOrdered(pPr, jc, PPR_ORDER);
  }
  if (!pPr.children.length) pPr.remove();
}

// Line spacing as a multiple (1.15 → w:line 276) and space after in points (→ twentieths)
function setSpacing(xml, xp, line, after) {
  const pPr = paragraphProps(xml, xp);
  let sp = childEl(pPr, "spacing");
  if (!sp) {
    sp = xml.createElementNS(W, "w:spacing");
    insertOrdered(pPr, sp, PPR_ORDER);
  }
  if (line != null) {
    sp.setAttributeNS(W, "w:line", String(Math.round(line * 240)));
    sp.setAttributeNS(W, "w:lineRule", "auto");
  }
  if (after != null) sp.setAttributeNS(W, "w:after", String(Math.round(after * 20)));
}

function setNumbering(xml, xp, numId, level) {
  const pPr = paragraphProps(xml, xp);
  [...pPr.children].filter((c) => c.localName === "numPr").forEach((c) => c.remove());
  const numPr = xml.createElementNS(W, "w:numPr");
  for (const [tag, val] of [
    ["ilvl", level],
    ["numId", numId],
  ]) {
    const x = numPr.appendChild(xml.createElementNS(W, `w:${tag}`));
    x.setAttributeNS(W, "w:val", String(val));
  }
  insertOrdered(pPr, numPr, PPR_ORDER);
}

function readStyles(stylesXml) {
  const styles = [{ id: "", key: "normal", label: "Normal", classes: [] }];
  if (!stylesXml) return styles;
  const all = [...stylesXml.getElementsByTagNameNS(W, "style")].filter((s) => s.getAttributeNS(W, "type") === "paragraph");
  // The default paragraph style ("Normal", or its local name) is what unstyled paragraphs show
  const def = all.find((s) => ["1", "true", "on"].includes(s.getAttributeNS(W, "default")));
  if (def) styles[0].classes = [`docx_${escapeClass(def.getAttributeNS(W, "styleId"))}`];
  for (const [name, label] of STYLE_NAMES) {
    const s = all.find((x) => x.getElementsByTagNameNS(W, "name")[0]?.getAttributeNS(W, "val").toLowerCase() === name);
    if (!s) continue;
    const id = s.getAttributeNS(W, "styleId");
    const num = s.getElementsByTagNameNS(W, "numId")[0]?.getAttributeNS(W, "val");
    const lvl = s.getElementsByTagNameNS(W, "ilvl")[0]?.getAttributeNS(W, "val") ?? "0";
    if (name.startsWith("list") && !num) continue;
    styles.push({ id, key: name, label, classes: [`docx_${escapeClass(id)}`, ...(num ? [`docx-num-${num}-${lvl}`] : [])] });
  }
  return styles;
}

function hostOf(node) {
  return elementOf(node)?.closest('[contenteditable="true"]');
}

function caretAt(p, offset) {
  const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  let n;
  while ((n = walk.nextNode()) && offset > n.length) offset -= n.length;
  const r = document.createRange();
  if (n) r.setStart(n, offset);
  else r.setStart(p, p.childNodes.length);
  hostOf(p)?.focus({ preventScroll: true });
  const sel = getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
  elementOf(r.startContainer)?.scrollIntoView({ block: "nearest" });
}

function caretOffset(p) {
  const sel = getSelection();
  if (!sel.rangeCount || !p.contains(sel.anchorNode)) return 0;
  const r = document.createRange();
  r.setStart(p, 0);
  r.setEnd(sel.anchorNode, sel.anchorOffset);
  return r.toString().length;
}

// Returns null when the rendered document can't be matched 1:1 with the file
export async function createEditor(view, data) {
  const zip = await JSZip.loadAsync(data);
  const xml = parseXml(await zip.file("word/document.xml").async("string"));
  const body = xml.getElementsByTagNameNS(W, "body")[0];
  const stylesFile = zip.file("word/styles.xml");
  const styles = readStyles(stylesFile && parseXml(await stylesFile.async("string")));
  const xmlBlocks = [...body.getElementsByTagNameNS(W, "*")].filter((e) => e.localName === "p" || e.localName === "tbl");
  const blocks = [...view.querySelectorAll("article p, article table")];
  const same =
    blocks.length === xmlBlocks.length &&
    blocks.every((b, i) =>
      b.tagName === "TABLE"
        ? xmlBlocks[i].localName === "tbl"
        : xmlBlocks[i].localName === "p" && squash(b.textContent) === squash(xmlText(xmlBlocks[i])),
    );
  if (!same) return null;

  const nodes = new Map(); // block id -> its XML element (kept when deleted, for undo)
  const saved = new Map(); // paragraph id -> HTML at the last commit
  let nextId = 0;
  blocks.forEach((b, i) => {
    const id = String(nextId++);
    b.dataset.xid = id;
    nodes.set(id, xmlBlocks[i]);
    if (b.tagName !== "P") return;
    if (KEEP.some((tag) => xmlBlocks[i].getElementsByTagNameNS("*", tag).length)) b.dataset.keep = "";
    for (const media of b.querySelectorAll("img, svg")) topChild(b, media).contentEditable = "false";
  });
  for (const a of view.querySelectorAll("article")) {
    a.contentEditable = "true";
    a.spellcheck = true;
  }
  blocks.forEach((b) => b.tagName === "P" && saved.set(b.dataset.xid, b.outerHTML));
  const initial = new Map(saved); // paragraph HTML as opened, for "changes since you opened"

  // Package parts touched when images are inserted
  const relsPath = "word/_rels/document.xml.rels";
  const rels = parseXml((await zip.file(relsPath)?.async("string")) ?? `<Relationships xmlns="${PKG_REL}"/>`);
  const types = parseXml(await zip.file("[Content_Types].xml").async("string"));
  const images = new Map(); // image id -> { data, ext, cx, cy, rid? }
  const rowNodes = new Map(); // table row id -> its w:tr
  const cellNodes = new Map(); // table cell id -> its w:tc
  const tableShape = new Map(); // table id -> row/cell layout at the last sync
  let nextImage = 0;
  let nextRow = 0;
  let nextCell = 0;
  let packageDirty = false;
  const X = (tag) => xml.createElementNS(W, `w:${tag}`);
  const shapeOf = (el) => [...el.rows].map((tr) => `${tr.dataset.rid}:${[...tr.cells].map((td) => td.dataset.cid).join(",")}`).join("|");

  // Links screen rows/cells to Word's; tables with merged cells keep their structure fixed
  function mapTable(el, tbl) {
    const xrows = [...tbl.children].filter((c) => c.localName === "tr");
    const xcells = (xtr) => [...xtr.children].filter((c) => c.localName === "tc");
    const merged = (tc) => [...(childEl(tc, "tcPr")?.children ?? [])].some((c) => c.localName === "vMerge" || c.localName === "gridSpan");
    const rows = [...el.rows];
    const simple =
      rows.length === xrows.length &&
      rows.every(
        (tr, i) =>
          tr.cells.length === xcells(xrows[i]).length &&
          [...tr.cells].every((td) => td.colSpan === 1 && td.rowSpan === 1) &&
          !xcells(xrows[i]).some(merged),
      );
    if (!simple) {
      el.dataset.fixed = "";
      return;
    }
    rows.forEach((tr, i) => {
      tr.dataset.rid = `r${nextRow++}`;
      rowNodes.set(tr.dataset.rid, xrows[i]);
      const tcs = xcells(xrows[i]);
      [...tr.cells].forEach((td, j) => {
        td.dataset.cid = `c${nextCell++}`;
        cellNodes.set(td.dataset.cid, tcs[j]);
      });
    });
    tableShape.set(el.dataset.xid, shapeOf(el));
  }
  for (const b of blocks) if (b.tagName === "TABLE") mapTable(b, nodes.get(b.dataset.xid));

  // A new empty paragraph in the document's default style
  const plainParagraph = () => {
    const p = document.createElement("p");
    p.classList.add(...styles[0].classes);
    return p;
  };
  const changed = () => view.dispatchEvent(new Event("input", { bubbles: true }));
  const articles = () => [...view.querySelectorAll("article")];
  const range = () => {
    const s = getSelection();
    return s.rangeCount ? s.getRangeAt(0) : null;
  };
  const inDoc = (r) => r && articles().some((a) => a.contains(r.commonAncestorContainer) || r.commonAncestorContainer.contains(a));

  // Paragraphs touched by the selection, or the one holding the caret
  function selectedParagraphs() {
    const r = range();
    if (!inDoc(r)) return [];
    if (r.collapsed) return [editableP(r.startContainer)].filter(Boolean);
    return [...view.querySelectorAll("article p")].filter((p) => p.isContentEditable && r.intersectsNode(p));
  }

  // ---------- Undo: whole-article DOM copies ----------
  const undoStack = [];
  const redoStack = [];
  let lastTyping = 0;
  const snapshot = () => {
    const p = editableP(getSelection().anchorNode);
    return { copies: articles().map((a) => a.cloneNode(true)), xid: p?.dataset.xid, offset: p ? caretOffset(p) : 0 };
  };
  function record(typing = false) {
    const now = Date.now();
    if (typing && now - lastTyping < 1000) {
      lastTyping = now;
      return;
    }
    lastTyping = typing ? now : 0;
    undoStack.push(snapshot());
    if (undoStack.length > 50) undoStack.shift();
    redoStack.length = 0;
  }
  function restore(from, to) {
    const s = from.pop();
    if (!s) return;
    to.push(snapshot());
    articles().forEach((a, i) => a.replaceWith(s.copies[i]));
    lastTyping = 0;
    const p = s.xid && view.querySelector(`article p[data-xid="${s.xid}"]`);
    if (p) caretAt(p, s.offset);
    changed();
  }
  const undo = () => restore(undoStack, redoStack);
  const redo = () => restore(redoStack, undoStack);

  // ---------- Styles and alignment ----------
  function currentStyle(p) {
    if (!p) return null;
    if ("style" in p.dataset) return p.dataset.style;
    const match = styles.find((s) => s.classes[0] && p.classList.contains(s.classes[0]));
    if (match) return match.id;
    return [...p.classList].some((c) => c.startsWith("docx_") || c.startsWith("docx-num-")) ? null : "";
  }
  const styleKey = (p) => styles.find((s) => s.id === currentStyle(p))?.key;

  function applyStyle(p, id) {
    const style = styles.find((s) => s.id === id);
    if (!style) return;
    [...p.classList].filter((c) => c.startsWith("docx_") || c.startsWith("docx-num-")).forEach((c) => p.classList.remove(c));
    p.classList.add(...style.classes);
    p.dataset.style = id;
    delete p.dataset.lvl; // the new style brings its own numbering, if any
    delete p.dataset.numid;
  }

  // The link at the caret as { href, range } covering all of it, or null
  function linkRange() {
    const r = range();
    if (!inDoc(r)) return null;
    const hrefOf = (x) => x.dataset.href ?? x.getAttribute("href");
    const el = elementOf(r.startContainer)?.closest("a[href], [data-href]");
    const p = el && editableP(el);
    if (!p) return null;
    const href = hrefOf(el);
    const parts = [...p.querySelectorAll("a[href], [data-href]")].filter((x) => hrefOf(x) === href);
    const whole = document.createRange();
    whole.setStartBefore(parts[0]);
    whole.setEndAfter(parts.at(-1));
    return { href, range: whole };
  }

  // ---------- Spacing and list levels ----------

  const LINES = [1, 1.15, 1.5, 2];

  // line: multiple (1, 1.15, 1.5, 2); after: points below the paragraph; undefined leaves either alone
  function setParagraphSpacing({ line, after }) {
    const ps = selectedParagraphs();
    if (!ps.length) return;
    record();
    for (const p of ps) {
      if (line != null) {
        p.style.lineHeight = String(line);
        p.dataset.line = line;
      }
      if (after != null) {
        p.style.marginBottom = `${after}pt`;
        p.dataset.after = after;
      }
    }
    changed();
  }

  const listOf = (p) => {
    const c = [...p.classList].find((x) => x.startsWith("docx-num-"));
    const m = c?.match(/^docx-num-(.+)-(\d+|undefined)$/);
    return m ? { cls: c, numId: m[1], level: +m[2] || 0 } : null;
  };

  // Moves list items a level in or out (Tab / Shift+Tab); false when nothing selected is a list item
  function shiftLevel(delta) {
    const items = selectedParagraphs()
      .map((p) => [p, listOf(p)])
      .filter(([, l]) => l);
    if (!items.length) return false;
    record();
    for (const [p, l] of items) {
      const level = Math.max(0, Math.min(8, l.level + delta));
      p.classList.replace(l.cls, `docx-num-${l.numId}-${level}`);
      p.dataset.lvl = level;
      p.dataset.numid = l.numId;
    }
    changed();
    return true;
  }

  function setStyle(id) {
    const ps = selectedParagraphs();
    if (!ps.length || !styles.some((s) => s.id === id)) return;
    record();
    ps.forEach((p) => applyStyle(p, id));
    changed();
  }

  function setStyleByKey(key) {
    const style = styles.find((s) => s.key === key);
    if (style) setStyle(style.id);
  }

  // A list button toggles: on a list paragraph it goes back to Normal
  function toggleList(key) {
    const ps = selectedParagraphs();
    const style = styles.find((s) => s.key === key);
    if (!ps.length || !style) return;
    setStyle(ps.every((p) => styleKey(p) === key) ? "" : style.id);
  }

  function setAlign(align) {
    const ps = selectedParagraphs();
    if (!ps.length) return;
    record();
    for (const p of ps) {
      p.style.textAlign = align;
      p.dataset.align = align;
    }
    changed();
  }

  // ---------- Text formatting ----------

  // Moves a text node into its own span that carries all styling of its former ancestors
  function isolate(n, p) {
    const chain = [];
    for (let a = n.parentNode; a !== p; a = a.parentNode) chain.push(a);
    let child = n;
    for (const a of chain) {
      if (child.previousSibling) {
        const before = a.cloneNode(false);
        while (a.firstChild !== child) before.appendChild(a.firstChild);
        a.before(before);
      }
      if (child.nextSibling) {
        const after = a.cloneNode(false);
        while (child.nextSibling) after.appendChild(child.nextSibling);
        a.after(after);
      }
      child = a;
    }
    const span = document.createElement("span");
    for (const a of [...chain].reverse()) {
      a.classList.forEach((c) => span.classList.add(c));
      span.style.cssText += ";" + (a.getAttribute("style") ?? "");
      if (a.tagName === "B" || a.tagName === "STRONG") span.style.fontWeight = "bold";
      if (a.tagName === "I" || a.tagName === "EM") span.style.fontStyle = "italic";
      if (a.tagName === "U") span.style.textDecorationLine = "underline";
      const href = a.dataset.href ?? (a.tagName === "A" ? a.getAttribute("href") : null);
      if (href) span.dataset.href = href;
    }
    (chain.at(-1) ?? n).replaceWith(span);
    span.appendChild(n);
    return span;
  }

  // bold/italic/underline/highlight toggle; size, color, font and link set `value` (null removes)
  function format(kind, value) {
    const sel = getSelection();
    const r = range();
    if (!inDoc(r) || r.collapsed) return;
    record();
    // Split at the selection edges; the live range follows the split text
    if (r.endContainer.nodeType === 3 && r.endOffset < r.endContainer.length) r.endContainer.splitText(r.endOffset);
    if (r.startContainer.nodeType === 3 && r.startOffset > 0) r.startContainer.splitText(r.startOffset);
    const texts = [];
    for (const p of selectedParagraphs()) {
      const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      for (let n = walk.nextNode(); n; n = walk.nextNode()) {
        if (!n.data || !r.intersectsNode(n) || !n.parentElement.isContentEditable) continue;
        if (n === r.endContainer && r.endOffset === 0) continue;
        if (n === r.startContainer && r.startOffset === n.length) continue;
        texts.push([n, p]);
      }
    }
    if (!texts.length) return;
    let apply;
    if (kind === "size") apply = (span) => (span.style.fontSize = `${value}pt`);
    else if (kind === "color") apply = (span) => (value ? (span.style.color = value) : span.style.removeProperty("color"));
    else if (kind === "font") apply = (span) => (value ? (span.style.fontFamily = `"${value}"`) : span.style.removeProperty("font-family"));
    else if (kind === "link") apply = (span) => (value ? (span.dataset.href = value) : delete span.dataset.href);
    else if (kind === "highlight") {
      const want = value?.replace("#", "").toUpperCase();
      const off = !want || texts.every(([n, p]) => formatOf(n, p).hl === want);
      apply = (span) => (off ? span.style.removeProperty("background-color") : (span.style.backgroundColor = value));
    } else {
      const [key, prop, on, off] = FORMATS[kind];
      const turnOn = !texts.every(([n, p]) => formatOf(n, p)[key]);
      apply = (span) => (span.style[prop] = turnOn ? on : off);
    }
    const spans = texts.map(([n, p]) => {
      const s = isolate(n, p);
      apply(s);
      return s;
    });
    const nr = document.createRange();
    nr.setStartBefore(spans[0]);
    nr.setEndAfter(spans.at(-1));
    sel.removeAllRanges();
    sel.addRange(nr);
    changed();
  }

  function stepSize(dir) {
    const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
    const cur = state().size ?? 11;
    const next = dir > 0 ? (sizes.find((s) => s > cur) ?? 72) : ([...sizes].reverse().find((s) => s < cur) ?? 8);
    format("size", next);
  }

  // What the format panel and selection bar should show for the current selection
  function state() {
    const ps = selectedParagraphs();
    const r = range();
    const out = { style: null, size: null, b: false, i: false, u: false, align: null, list: null, editable: ps.length > 0 };
    if (!ps.length) return out;
    const ids = new Set(ps.map(currentStyle));
    out.style = ids.size === 1 ? [...ids][0] : null;
    const keys = new Set(ps.map(styleKey));
    out.list = keys.size === 1 && [...keys][0]?.startsWith("list") ? [...keys][0] : null;
    const aligns = new Set(ps.map((p) => getComputedStyle(p).textAlign.replace("start", "left").replace("end", "right")));
    out.align = aligns.size === 1 ? [...aligns][0] : null;
    let texts = [];
    if (r.collapsed) {
      const n = r.startContainer.nodeType === 3 ? r.startContainer : null;
      texts = n ? [[n, ps[0]]] : [];
    } else {
      for (const p of ps) {
        const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        for (let n = walk.nextNode(); n && texts.length < 400; n = walk.nextNode())
          if (n.data.trim() && r.intersectsNode(n)) texts.push([n, p]);
      }
    }
    const fs = texts.length ? texts.map(([n, p]) => formatOf(n, p)) : [styleDefault(ps[0])];
    out.b = fs.every((f) => f.b);
    out.i = fs.every((f) => f.i);
    out.u = fs.every((f) => f.u);
    const sizes = new Set(fs.map((f) => f.size));
    out.size = sizes.size === 1 ? [...sizes][0] : null;
    const one = (key) => {
      const v = new Set(fs.map((f) => f[key]));
      return v.size === 1 ? [...v][0] : undefined;
    };
    out.color = one("color");
    out.hl = one("hl");
    out.font = one("font");
    out.href = one("href");
    // Line spacing and space after of the first paragraph, snapped to the offered values
    const ps0 = getComputedStyle(ps[0]);
    const span0 = getComputedStyle(ps[0].querySelector("span") ?? ps[0]);
    const ratio = ps0.lineHeight === "normal" ? 1.15 : parseFloat(span0.lineHeight) / parseFloat(span0.fontSize);
    out.line = LINES.reduce((best, v) => (Math.abs(v - ratio) < Math.abs(best - ratio) ? v : best), LINES[0]);
    out.after = Math.round(parseFloat(ps0.marginBottom) * 0.75);
    return out;
  }

  // ---------- Paragraph structure ----------

  function newParagraphLike(p) {
    const np = p.cloneNode(false);
    delete np.dataset.keep;
    np.dataset.xid = String(nextId++);
    np.dataset.src = nodes.has(p.dataset.xid) ? p.dataset.xid : p.dataset.src;
    return np;
  }

  function split(p) {
    const r = range();
    if (!r.collapsed) r.deleteContents();
    if (!p.textContent && [...p.classList].some((c) => c.startsWith("docx-num-"))) {
      applyStyle(p, ""); // Enter on an empty list item ends the list
      return;
    }
    const tail = document.createRange();
    tail.setStart(r.startContainer, r.startOffset);
    tail.setEnd(p, p.childNodes.length);
    const frag = tail.extractContents();
    const np = newParagraphLike(p);
    np.append(frag);
    for (const x of [p, np])
      x.querySelectorAll("span, b, i, u").forEach((e) => !e.textContent && !e.querySelector("img, br, svg") && e.remove());
    p.after(np);
    const key = styleKey(p);
    if (!np.textContent && (key === "title" || key?.startsWith("heading"))) applyStyle(np, "");
    caretAt(np, 0);
  }

  // Joins `next` onto the end of `p`; refuses across cells or when it would drop protected content
  function join(p, next) {
    if (p?.tagName !== "P" || next?.tagName !== "P" || next.parentNode !== p.parentNode) return false;
    if (!p.textContent && !("keep" in p.dataset)) {
      p.remove();
      caretAt(next, 0);
      return true;
    }
    if ("keep" in next.dataset) return false;
    const offset = p.textContent.length;
    p.append(...next.childNodes);
    next.remove();
    caretAt(p, offset);
    return true;
  }

  // The block (child of `container`) holding `node`, or null when node is the container itself
  const blockIn = (container, node) => (node === container ? null : topChild(container, node));

  // Deletes a selection that spans paragraphs; tables fully inside it go too
  function deleteSelection() {
    const r = range();
    const c = elementOf(r.commonAncestorContainer);
    if (editableP(c)) {
      r.deleteContents();
      return true;
    }
    if (!(c.tagName === "ARTICLE" || (c.tagName === "TD" && c.isContentEditable))) return false;
    const first = blockIn(c, r.startContainer);
    const last = blockIn(c, r.endContainer);
    const ok = (b) => !b || (b.tagName === "P" && !("keep" in b.dataset));
    if (!ok(first) || !ok(last)) return false;
    r.deleteContents();
    if (first && last && first !== last) {
      const offset = first.textContent.length;
      first.append(...last.childNodes);
      last.remove();
      caretAt(first, offset);
    } else if (first || last) caretAt(first ?? last, first ? first.textContent.length : 0);
    if (![...c.children].some((x) => x.tagName === "P")) {
      // Never leave a document or cell without a paragraph
      const p = plainParagraph();
      p.dataset.xid = String(nextId++);
      c.append(p);
      caretAt(p, 0);
    }
    return true;
  }

  const multiBlock = (r) => r && !r.collapsed && !editableP(r.commonAncestorContainer);

  function selectAll() {
    const all = articles();
    if (!all.length) return;
    const r = document.createRange();
    r.setStart(all[0], 0);
    r.setEnd(all.at(-1), all.at(-1).childNodes.length);
    all[0].focus({ preventScroll: true });
    getSelection().removeAllRanges();
    getSelection().addRange(r);
  }

  // Tab moves between table cells like Word; elsewhere it types a tab
  function tab(p, back) {
    const td = p.closest("td");
    if (!td) {
      document.execCommand("insertText", false, "\u2003");
      return;
    }
    const table = td.closest("table");
    const cells = [...table.querySelectorAll("td")].filter((c) => c.closest("table") === table && c.querySelector("p"));
    const next = cells[cells.indexOf(td) + (back ? -1 : 1)];
    if (!next && !back) {
      // Tab in the last cell adds a row, like Word
      if (!tableOp("rowBelow", td, { quiet: true })) return;
      caretAt(table.rows[table.rows.length - 1].cells[0].querySelector("p"), 0);
      return;
    }
    const target = next?.querySelector("p");
    if (target) caretAt(target, back ? target.textContent.length : 0);
  }

  // ---------- Table structure ----------

  const caretCell = () => {
    const r = range();
    const p = inDoc(r) ? editableP(r.startContainer) : null;
    return p?.closest("td") ?? null;
  };

  // What table tools apply where the caret is
  function tableInfo() {
    const td = caretCell();
    if (!td) return null;
    return { editable: !("fixed" in td.closest("table").dataset) };
  }

  // An empty cell shaped like `td` (its paragraph style and Word cell settings carry over)
  function cellLike(td) {
    const cell = td.cloneNode(false);
    cell.removeAttribute("rowspan");
    cell.removeAttribute("colspan");
    cell.dataset.cid = `c${nextCell++}`;
    cell.dataset.from = td.dataset.cid ?? "";
    const src = td.querySelector(":scope > p");
    const p = src ? src.cloneNode(false) : plainParagraph();
    for (const k of ["keep", "img", "style", "align"]) delete p.dataset[k];
    p.dataset.xid = String(nextId++);
    p.dataset.src = src ? (nodes.has(src.dataset.xid) ? src.dataset.xid : (src.dataset.src ?? "")) : "";
    cell.append(p);
    return cell;
  }

  // Columns share the width evenly after adding or removing one
  function evenColumns(table) {
    const cols = Math.max(1, ...[...table.rows].map((r) => r.cells.length));
    const group = table.querySelector(":scope > colgroup");
    if (group)
      group.replaceChildren(
        ...Array.from({ length: cols }, () => Object.assign(document.createElement("col"), { style: `width: ${100 / cols}%` })),
      );
  }

  function deleteTable(table) {
    const after = paragraphAfter(table);
    table.remove();
    caretAt(after, 0);
  }

  // op: rowAbove | rowBelow | colLeft | colRight | deleteRow | deleteCol | deleteTable; returns false when not possible
  function tableOp(op, td = caretCell(), { quiet = false } = {}) {
    const table = td?.closest("table");
    const ops = ["rowAbove", "rowBelow", "colLeft", "colRight", "deleteRow", "deleteCol", "deleteTable"];
    if (!table || "fixed" in table.dataset || !ops.includes(op)) return false;
    const tr = td.parentElement;
    const col = td.cellIndex;
    const colCount = Math.max(...[...table.rows].map((r) => r.cells.length));
    record();
    let focus = null;
    if (op === "rowAbove" || op === "rowBelow") {
      const row = tr.cloneNode(false);
      row.dataset.rid = `r${nextRow++}`;
      for (const c of tr.cells) row.append(cellLike(c));
      if (op === "rowAbove") tr.before(row);
      else tr.after(row);
      focus = row.cells[col];
    } else if (op === "colLeft" || op === "colRight") {
      for (const row of table.rows) {
        const ref = row.cells[Math.min(col, row.cells.length - 1)];
        const cell = cellLike(ref);
        if (op === "colLeft") ref.before(cell);
        else ref.after(cell);
      }
      evenColumns(table);
      focus = tr.cells[op === "colLeft" ? col : col + 1];
    } else if (op === "deleteRow" && table.rows.length > 1) {
      const other = tr.nextElementSibling ?? tr.previousElementSibling;
      tr.remove();
      focus = other.cells[Math.min(col, other.cells.length - 1)];
    } else if (op === "deleteCol" && colCount > 1) {
      for (const row of table.rows) row.cells[Math.min(col, row.cells.length - 1)]?.remove();
      evenColumns(table);
      focus = tr.cells[Math.min(col, tr.cells.length - 1)];
    } else {
      deleteTable(table); // deleting the last row or column removes the table
    }
    if (focus && !quiet) caretAt(focus.querySelector("p"), 0);
    changed();
    return true;
  }

  // The top-level block (paragraph or table) holding the caret; new content goes after it
  function anchorBlock() {
    const r = range();
    const p = inDoc(r) ? editableP(r.startContainer) : null;
    if (!p) return null;
    return topChild(p.closest("article"), p);
  }

  // Makes sure there is a paragraph to type in after `el`
  function paragraphAfter(el) {
    if (el.nextElementSibling?.tagName === "P") return el.nextElementSibling;
    const p = plainParagraph();
    p.dataset.xid = String(nextId++);
    el.after(p);
    return p;
  }

  function insertTable(rows = 3, cols = 3) {
    const at = anchorBlock();
    if (!at) return;
    record();
    const table = document.createElement("table");
    table.className = "dv-table";
    table.dataset.xid = String(nextId++);
    for (let r = 0; r < rows; r++) {
      const tr = table.insertRow();
      tr.dataset.rid = `r${nextRow++}`;
      for (let c = 0; c < cols; c++) {
        const p = plainParagraph();
        p.dataset.xid = String(nextId++);
        const td = tr.insertCell();
        td.dataset.cid = `c${nextCell++}`;
        td.append(p);
      }
    }
    at.after(table);
    paragraphAfter(table);
    caretAt(table.rows[0].cells[0].querySelector("p"), 0);
    changed();
  }

  // data: image bytes; ext: png | jpeg | gif; width/height in CSS pixels
  function insertImage(data, ext, width, height) {
    const at = anchorBlock();
    if (!at) return;
    record();
    const room = at.closest("article").clientWidth;
    const scale = Math.min(1, room / width);
    const w = Math.round(width * scale);
    const h = Math.round(height * scale);
    const id = String(nextImage++);
    images.set(id, { data, ext, cx: w * 9525, cy: h * 9525 });
    const p = plainParagraph();
    p.dataset.xid = String(nextId++);
    p.dataset.img = id;
    p.dataset.keep = "";
    const holder = document.createElement("span");
    holder.contentEditable = "false";
    const img = new Image(w, h);
    img.src = URL.createObjectURL(new Blob([data], { type: MIME[ext] }));
    holder.append(img);
    p.append(holder);
    at.after(p);
    caretAt(paragraphAfter(p), 0);
    changed();
  }

  // ---------- Commit: on-screen state into the XML, one at a time ----------
  let queue = Promise.resolve(null);
  const commit = () => (queue = queue.then(doCommit, doCommit));

  function place(xe, el, src) {
    const prev = el.previousElementSibling?.dataset?.xid && nodes.get(el.previousElementSibling.dataset.xid);
    const next = el.nextElementSibling?.dataset?.xid && nodes.get(el.nextElementSibling.dataset.xid);
    if (prev && xml.contains(prev)) prev.after(xe);
    else if (next && xml.contains(next)) next.before(xe);
    else if (src && xml.contains(src)) src.after(xe);
    else {
      const sect = childEl(body, "sectPr");
      body.insertBefore(xe, sect ?? null);
    }
  }

  const fragment = (markup) =>
    xml.importNode(
      parseXml(
        `<w:r xmlns:w="${W}" xmlns:r="${DOC_REL}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">${markup}</w:r>`,
      ).documentElement,
      true,
    );

  // A Word table mirroring the DOM table; its paragraphs get mapped as they are made
  function buildTable(el) {
    const cols = Math.max(...[...el.rows].map((r) => r.cells.length));
    const width = Math.round(9000 / cols);
    const border = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>`;
    const tbl = fragment(
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map(border).join("")}</w:tblBorders>` +
        `<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
        `<w:tblGrid>${`<w:gridCol w:w="${width}"/>`.repeat(cols)}</w:tblGrid></w:tbl>`,
    ).firstChild;
    return tbl;
  }

  // A Word cell for a screen cell: reused when it exists, else made with the settings of the cell it was copied from
  function wordCell(td, width) {
    let tc = cellNodes.get(td.dataset.cid);
    if (!tc) {
      tc = X("tc");
      const from = cellNodes.get(td.dataset.from);
      const tcPr = from && childEl(from, "tcPr");
      if (tcPr) {
        const copy = tcPr.cloneNode(true);
        [...copy.children].filter((c) => c.localName === "vMerge" || c.localName === "gridSpan").forEach((c) => c.remove());
        tc.appendChild(copy);
      }
      cellNodes.set(td.dataset.cid, tc);
    }
    if (width) {
      const tcPr = childEl(tc, "tcPr") ?? tc.insertBefore(X("tcPr"), tc.firstChild);
      const w = childEl(tcPr, "tcW") ?? tcPr.insertBefore(X("tcW"), tcPr.firstChild);
      w.setAttributeNS(W, "w:w", String(width));
      w.setAttributeNS(W, "w:type", "dxa");
    }
    // The cell's paragraphs and nested tables, in screen order
    const content = [...td.children]
      .filter((c) => c.tagName === "P" || c.tagName === "TABLE")
      .map((k) => {
        let x = nodes.get(k.dataset.xid);
        if (!x && k.tagName === "P") {
          x = X("p");
          const pPr = nodes.get(k.dataset.src) && childEl(nodes.get(k.dataset.src), "pPr");
          if (pPr) {
            const copy = pPr.cloneNode(true);
            [...copy.getElementsByTagNameNS(W, "sectPr")].forEach((s) => s.remove());
            x.appendChild(copy);
          }
          nodes.set(k.dataset.xid, x);
        } else if (!x) {
          x = buildTable(k);
          nodes.set(k.dataset.xid, x);
        }
        return x;
      });
    [...tc.children].filter((c) => c.localName === "p" || c.localName === "tbl").forEach((c) => c.remove());
    for (const x of content) tc.appendChild(x);
    if (content.at(-1)?.localName !== "p") tc.appendChild(X("p")); // Word needs a paragraph last in a cell
    return tc;
  }

  // Makes the Word table's rows and cells match the screen; only runs when the layout changed
  function syncTable(el, tbl) {
    if ("fixed" in el.dataset) return false;
    const shape = shapeOf(el);
    if (tableShape.get(el.dataset.xid) === shape) return false;
    const rows = [...el.rows];
    const cols = Math.max(1, ...rows.map((r) => r.cells.length));
    let grid = childEl(tbl, "tblGrid");
    const gridCols = grid ? [...grid.children].filter((c) => c.localName === "gridCol") : [];
    const total = gridCols.reduce((s, c) => s + (+c.getAttributeNS(W, "w") || 0), 0) || 9000;
    const width = gridCols.length !== cols ? Math.round(total / cols) : null;
    if (width) {
      if (!grid) grid = tbl.insertBefore(X("tblGrid"), childEl(tbl, "tblPr")?.nextSibling ?? tbl.firstChild);
      grid.replaceChildren(
        ...Array.from({ length: cols }, () => {
          const g = X("gridCol");
          g.setAttributeNS(W, "w:w", String(width));
          return g;
        }),
      );
    }
    for (const tr of rows) {
      let xtr = rowNodes.get(tr.dataset.rid);
      if (!xtr) {
        xtr = X("tr");
        rowNodes.set(tr.dataset.rid, xtr);
      }
      const tcs = [...tr.cells].map((td) => wordCell(td, width));
      [...xtr.children].filter((c) => c.localName === "tc").forEach((c) => c.remove());
      for (const tc of tcs) xtr.appendChild(tc);
    }
    [...tbl.children].filter((c) => c.localName === "tr").forEach((c) => c.remove());
    for (const tr of rows) tbl.appendChild(rowNodes.get(tr.dataset.rid));
    tableShape.set(el.dataset.xid, shape);
    return true;
  }

  // Writes on-screen image sizes (set by resizing) into the Word drawings
  function applyImageSizes(xp, p) {
    const imgs = [...p.querySelectorAll("img")];
    const drawings = [...xp.getElementsByTagNameNS("*", "drawing")];
    if (imgs.length !== drawings.length) return;
    imgs.forEach((img, i) => {
      const box = img.closest("[data-w]");
      if (!box) return;
      const cx = String(Math.round(+box.dataset.w * 9525));
      const cy = String(Math.round(+box.dataset.h * 9525));
      for (const e of drawings[i].getElementsByTagNameNS("*", "extent")) {
        e.setAttribute("cx", cx);
        e.setAttribute("cy", cy);
      }
      for (const x of drawings[i].getElementsByTagNameNS("*", "xfrm")) {
        for (const e of x.children)
          if (e.localName === "ext") {
            e.setAttribute("cx", cx);
            e.setAttribute("cy", cy);
          }
      }
    });
  }

  // Relationship id for an external link, reusing an existing one for the same address
  function linkId(href) {
    const existing = [...rels.documentElement.children].find(
      (r) => r.getAttribute("Type") === `${DOC_REL}/hyperlink` && r.getAttribute("Target") === href,
    );
    if (existing) return existing.getAttribute("Id");
    const taken = new Set([...rels.documentElement.children].map((r) => r.getAttribute("Id")));
    let n = 1;
    while (taken.has(`rIdDvL${n}`)) n++;
    const rel = rels.createElementNS(PKG_REL, "Relationship");
    rel.setAttribute("Id", `rIdDvL${n}`);
    rel.setAttribute("Type", `${DOC_REL}/hyperlink`);
    rel.setAttribute("Target", href);
    rel.setAttribute("TargetMode", "External");
    rels.documentElement.appendChild(rel);
    packageDirty = true;
    return `rIdDvL${n}`;
  }

  function addDrawing(xp, img) {
    if (!img || xp.getElementsByTagNameNS("*", "drawing").length) return;
    if (!img.rid) {
      const taken = new Set([...rels.documentElement.children].map((r) => r.getAttribute("Id")));
      let n = 1;
      while (taken.has(`rIdDv${n}`)) n++;
      img.rid = `rIdDv${n}`;
      const target = `media/dv-image-${Date.now()}-${n}.${img.ext}`;
      zip.file(`word/${target}`, img.data);
      const rel = rels.createElementNS(PKG_REL, "Relationship");
      rel.setAttribute("Id", img.rid);
      rel.setAttribute("Type", `${DOC_REL}/image`);
      rel.setAttribute("Target", target);
      rels.documentElement.appendChild(rel);
      const ext = img.ext.toLowerCase();
      const known = [...types.documentElement.children].some(
        (d) => d.localName === "Default" && d.getAttribute("Extension")?.toLowerCase() === ext,
      );
      if (!known) {
        const d = types.createElementNS(types.documentElement.namespaceURI, "Default");
        d.setAttribute("Extension", ext);
        d.setAttribute("ContentType", MIME[ext]);
        types.documentElement.insertBefore(d, types.documentElement.firstChild);
      }
      packageDirty = true;
    }
    const ids = [...xml.getElementsByTagNameNS("*", "docPr")].map((d) => +d.getAttribute("id") || 0);
    const docId = Math.max(0, ...ids) + 1;
    const run = fragment(
      `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${img.cx}" cy="${img.cy}"/>` +
        `<wp:docPr id="${docId}" name="Picture ${docId}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
        `<pic:nvPicPr><pic:cNvPr id="0" name="Picture ${docId}"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="${img.rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${img.cx}" cy="${img.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
        `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`,
    );
    xp.appendChild(run);
  }

  async function doCommit() {
    const live = [...view.querySelectorAll("article p[data-xid], article table[data-xid]")];
    const present = new Set();
    let dirty = false;
    for (const el of live) {
      const id = el.dataset.xid;
      present.add(id);
      const src = nodes.get(el.dataset.src);
      let xe = nodes.get(id);
      if (!xe && el.tagName === "TABLE") {
        xe = buildTable(el);
        nodes.set(id, xe);
      } else if (!xe) {
        xe = xml.createElementNS(W, "w:p");
        const pPr = src && childEl(src, "pPr");
        if (pPr) {
          const copy = pPr.cloneNode(true);
          [...copy.getElementsByTagNameNS(W, "sectPr")].forEach((x) => x.remove());
          xe.appendChild(copy);
        }
        nodes.set(id, xe);
      }
      const placed = xml.contains(xe);
      if (!placed) {
        place(xe, el, src);
        dirty = true;
      }
      if (el.tagName === "TABLE") {
        if (syncTable(el, xe)) dirty = true;
        continue;
      }
      const html = el.outerHTML;
      if (placed && saved.get(id) === html) continue;
      rewriteParagraph(xml, xe, el, src, linkId);
      if ("img" in el.dataset) addDrawing(xe, images.get(el.dataset.img));
      applyImageSizes(xe, el);
      if ("style" in el.dataset) setParagraphStyle(xml, xe, el.dataset.style);
      if ("align" in el.dataset) setAlignment(xml, xe, el.dataset.align);
      if ("line" in el.dataset || "after" in el.dataset) {
        setSpacing(xml, xe, "line" in el.dataset ? +el.dataset.line : null, "after" in el.dataset ? +el.dataset.after : null);
      }
      if ("lvl" in el.dataset) setNumbering(xml, xe, el.dataset.numid, +el.dataset.lvl);
      saved.set(id, html);
      dirty = true;
    }
    // Map order is document order, so a deleted table goes before its own paragraphs
    for (const [id, xe] of nodes) {
      if (!present.has(id) && xml.contains(xe)) {
        xe.remove();
        dirty = true;
      }
    }
    if (!dirty) return null;
    zip.file("word/document.xml", serialize(xml));
    if (packageDirty) {
      zip.file(relsPath, serialize(rels));
      zip.file("[Content_Types].xml", serialize(types));
      packageDirty = false;
    }
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  // ---------- Input handling ----------

  function onKey(e) {
    const k = e.key.toLowerCase();
    const r = range();
    if (!inDoc(r)) return false;
    const p = editableP(r.startContainer);
    const ctrl = e.ctrlKey && !e.altKey;
    const shortcuts =
      ctrl &&
      {
        z: () => (e.shiftKey ? redo() : undo()),
        y: redo,
        a: selectAll,
        b: () => format("bold"),
        i: () => format("italic"),
        u: () => format("underline"),
        e: () => setAlign("center"),
        l: () => (e.shiftKey ? toggleList("list bullet") : setAlign("left")),
        r: () => setAlign("right"),
        j: () => setAlign("justify"),
        n: e.shiftKey && (() => setStyle("")),
        "[": () => stepSize(-1),
        "]": () => stepSize(1),
      }[k];
    if (shortcuts) {
      e.preventDefault();
      shortcuts();
      return true;
    }
    if (e.ctrlKey && e.altKey && ["1", "2", "3"].includes(e.key)) {
      e.preventDefault();
      setStyleByKey(`heading ${e.key}`);
      return true;
    }
    if (e.ctrlKey || e.altKey) return false;
    if (e.key === "Tab" && p) {
      e.preventDefault();
      // In a list: Shift+Tab outdents; Tab indents at the start of an item or over a selection
      if (listOf(p) && (e.shiftKey || !r.collapsed || atEdge(p, r, false)) && shiftLevel(e.shiftKey ? -1 : 1)) return true;
      record();
      tab(p, e.shiftKey);
      changed();
      return true;
    }
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      if (multiBlock(r)) {
        record();
        if (!deleteSelection()) return true;
      }
      const at = editableP(range().startContainer);
      if (!at) return true;
      record();
      split(at);
      changed();
      return true;
    }
    if ((e.key === "Backspace" || e.key === "Delete") && multiBlock(r)) {
      e.preventDefault();
      const snap = snapshot();
      if (deleteSelection()) {
        undoStack.push(snap);
        redoStack.length = 0;
        changed();
      }
      return true;
    }
    const back = e.key === "Backspace" && p && r.collapsed && atEdge(p, r, false);
    const fwd = e.key === "Delete" && p && r.collapsed && atEdge(p, r, true);
    if (back || fwd) {
      e.preventDefault();
      const snap = snapshot();
      if (back ? join(p.previousElementSibling, p) : join(p, p.nextElementSibling)) {
        undoStack.push(snap);
        redoStack.length = 0;
        lastTyping = 0;
        changed();
      }
      return true;
    }
    return false;
  }

  function atEdge(p, r, end) {
    const t = document.createRange();
    t.selectNodeContents(p);
    if (end) t.setStart(r.startContainer, r.startOffset);
    else t.setEnd(r.startContainer, r.startOffset);
    return t.toString() === "";
  }

  function onBeforeInput(e) {
    if (e.inputType === "historyUndo") {
      e.preventDefault();
      return undo();
    }
    if (e.inputType === "historyRedo") {
      e.preventDefault();
      return redo();
    }
    if (e.inputType === "insertParagraph" || e.inputType.startsWith("format")) return e.preventDefault();
    const r = range();
    if (multiBlock(r)) {
      e.preventDefault();
      record();
      if (!deleteSelection()) return;
      if (e.inputType.startsWith("insert") && e.data) document.execCommand("insertText", false, e.data);
      changed();
      return;
    }
    record(true);
  }

  function cut(e) {
    const r = range();
    if (!inDoc(r) || r.collapsed) return;
    e?.preventDefault();
    e?.clipboardData.setData("text/plain", getSelection().toString());
    record();
    if (deleteSelection()) changed();
  }

  // Pasted lines become paragraphs, like Word
  function paste(text) {
    const r = range();
    if (!inDoc(r)) return;
    record();
    if (multiBlock(r) && !deleteSelection()) return;
    // Page breaks from PDFs become new lines; other control characters are dropped
    const lines = text
      .replace(/\r/g, "")
      .replace(/[\f\v]/g, "\n")
      .replace(NOT_XML, "")
      .split("\n");
    lines.forEach((line, k) => {
      if (k) {
        const p = editableP(range().startContainer);
        if (p) split(p);
      }
      if (line) document.execCommand("insertText", false, line.replace(/\t/g, "\u2003"));
    });
    changed();
  }

  // A range over characters [start, end) of a paragraph's text
  function textRange(p, start, end) {
    const r = document.createRange();
    const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let pos = 0;
    let startSet = false;
    for (let n = walk.nextNode(); n; n = walk.nextNode()) {
      if (!startSet && start <= pos + n.length) {
        r.setStart(n, start - pos);
        startSet = true;
      }
      if (startSet && end <= pos + n.length) {
        r.setEnd(n, end - pos);
        return r;
      }
      pos += n.length;
    }
    return null;
  }

  const BLOCK_SHORTCUTS = {
    "# ": "heading 1",
    "## ": "heading 2",
    "### ": "heading 3",
    "- ": "list bullet",
    "* ": "list bullet",
    "1. ": "list number",
  };
  const INLINE_SHORTCUTS = [
    [/\*\*([^*\s](?:[^*]*[^*\s])?)\*\*$/, "bold", 2, 0],
    [/(^|[\s(])\*([^*\s](?:[^*]*[^*\s])?)\*$/, "italic", 1, 1],
    [/(^|[\s(])_([^_\s](?:[^_]*[^_\s])?)_$/, "italic", 1, 1],
  ];

  // Markdown-style shortcuts, applied right after the triggering keystroke; Ctrl+Z undoes just the conversion
  function autoformat(e) {
    if (e.inputType !== "insertText" || !e.data) return;
    const p = editableP(getSelection().anchorNode);
    if (!p || !getSelection().isCollapsed) return;
    const off = caretOffset(p);
    const before = p.textContent.slice(0, off);
    if (e.data === " " && before in BLOCK_SHORTCUTS) {
      const style = styles.find((s) => s.key === BLOCK_SHORTCUTS[before]);
      if (!style || currentStyle(p) === style.id) return;
      record();
      textRange(p, 0, off)?.deleteContents();
      applyStyle(p, style.id);
      caretAt(p, 0);
      changed();
      return;
    }
    if (e.data !== "*" && e.data !== "_") return;
    for (const [re, kind, marker, lead] of INLINE_SHORTCUTS) {
      const m = before.match(re);
      if (!m) continue;
      const inner = m[lead ? 2 : 1];
      const start = off - m[0].length + (lead ? m[1].length : 0);
      record();
      textRange(p, off - marker, off)?.deleteContents();
      textRange(p, start, start + marker)?.deleteContents();
      const r = textRange(p, start, start + inner.length);
      if (!r) return;
      getSelection().removeAllRanges();
      getSelection().addRange(r);
      format(kind);
      // Continue after the new span in plain text; the invisible marker never reaches the file
      const sel = getSelection();
      const span = elementOf(sel.getRangeAt(0).endContainer.childNodes[sel.getRangeAt(0).endOffset - 1] ?? sel.anchorNode);
      const plain = document.createTextNode(ZWSP);
      (span?.closest("span") ?? span).after(plain);
      const c = document.createRange();
      c.setStart(plain, 1);
      sel.removeAllRanges();
      sel.addRange(c);
      return;
    }
  }

  // Replaces each range's text; formatting of the replaced text carries over. One undo step for all.
  function replaceRanges(ranges, text) {
    const editable = ranges.filter((r) => editableP(r.startContainer) && editableP(r.endContainer));
    if (!editable.length) return 0;
    record();
    for (const r of [...editable].reverse()) {
      r.deleteContents();
      if (r.startContainer.nodeType === 3) r.startContainer.insertData(r.startOffset, text);
      else if (text) r.insertNode(document.createTextNode(text));
    }
    changed();
    return editable.length;
  }

  // Runs a change made outside the editor (image resize, align) as one undo step
  function mutate(fn) {
    record();
    fn();
    changed();
  }

  // Paragraphs that differ from how the document was opened (new ones included)
  function changedIds() {
    return [...view.querySelectorAll("article p[data-xid]")]
      .filter((p) => initial.get(p.dataset.xid) !== p.outerHTML)
      .map((p) => p.dataset.xid);
  }

  // Paragraph styles target spans, so typed text must not sit directly in the paragraph
  function normalize() {
    const p = editableP(getSelection().anchorNode);
    const loose = p ? [...p.childNodes].filter((n) => n.nodeType === 3 && n.data) : [];
    if (!loose.length) return;
    const offset = caretOffset(p);
    for (const n of loose) {
      const span = document.createElement("span");
      n.replaceWith(span);
      span.append(n);
    }
    caretAt(p, offset);
  }

  function focusStart() {
    const p = view.querySelector("article p");
    if (p?.isContentEditable) caretAt(p, 0);
  }

  return {
    styles,
    commit,
    format,
    stepSize,
    setStyle,
    setStyleByKey,
    toggleList,
    setAlign,
    state,
    insertImage,
    insertTable,
    autoformat,
    changedIds,
    replaceRanges,
    mutate,
    tableOp,
    tableInfo,
    setParagraphSpacing,
    shiftLevel,
    linkRange,
    undo,
    redo,
    canUndo: () => undoStack.length > 0,
    canRedo: () => redoStack.length > 0,
    onKey,
    onBeforeInput,
    cut,
    paste,
    normalize,
    selectAll,
    focusStart,
    inDoc: () => inDoc(range()),
  };
}
