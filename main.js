import "@fontsource-variable/inter/wght.css";
import "./fonts.css";
import { renderAsync } from "docx-preview";
import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { open, save as saveDialog, ask, message } from "@tauri-apps/plugin-dialog";
import { readText } from "@tauri-apps/plugin-clipboard-manager";
import { createEditor } from "./editor.js";
import { createFind } from "./find.js";
import { blankDocx, TEMPLATES } from "./blank.js";
import { createImageTools } from "./imagetools.js";
import { createPalette, createSlashMenu } from "./palette.js";
import { createCaret } from "./caret.js";
import { createReader } from "./speech.js";
import { toMarkdown, toHtml } from "./export.js";

const isTauri = !!window.__TAURI_INTERNALS__;
const $ = (id) => document.getElementById(id);
const $$ = (sel) => [...document.querySelectorAll(sel)];
const view = $("view"); // the scroll area; each tab draws into its own pane inside it
const find = createFind(view, {
  replace: (ranges, text) => editor?.replaceRanges(ranges, text) ?? 0,
  canReplace: () => !!editor,
  onReplaced: (n) => hud(!n ? "Nothing replaced" : n === 1 ? "Replaced 1 match" : `Replaced ${n.toLocaleString()} matches`),
  root: () => pane,
});
const IMAGE_EXT = /\.(png|jpe?g|gif|webp|bmp)$/i;
const HEADING = /\bdocx_(title|heading(\d))\b/;
const SNAPSHOT_EVERY = 2 * 60 * 1000;

let zoom = 1;
let doc = null; // { name, path, data, original, startWords, snapshotted, lastSnap, unsnapped }
let editor = null; // null when nothing is open, the document is read-only, or an old version is shown
let dirty = false;
let version = 0; // bumps on every change, so a save knows if more came in meanwhile
let words = 0;
let previewing = null; // { ts, data } while looking at an old version

// Tabs: the variables above always hold the active tab's state; switching swaps them in and out
const tabs = []; // { pane, doc, editor, dirty, version, words, previewing, scroll }
let active = null;
let pane = null; // the active tab's content element

const confirmUser = (msg) => (isTauri ? ask(msg, { title: "Ribbonless", kind: "warning" }) : Promise.resolve(confirm(msg)));
const tell = (msg) => (isTauri ? message(msg, { title: "Ribbonless" }) : alert(msg));
const store = {
  get: (k, fallback) => {
    try {
      return JSON.parse(localStorage.getItem(k)) ?? fallback;
    } catch {
      return fallback;
    }
  },
  set: (k, v) => {
    try {
      localStorage.setItem(k, JSON.stringify(v));
    } catch {}
  },
};
const setting = (k, fallback = true) => store.get(k, fallback);
const raw = (cmd, body, headers = {}) =>
  invoke(cmd, body, { headers: Object.fromEntries(Object.entries(headers).map(([k, v]) => [k, encodeURIComponent(v)])) });
const autoSaving = () => isTauri && setting("autosave") && !!doc?.path && !!editor;
const check = (cmd, on) => $$(`[data-cmd="${cmd}"]`).forEach((b) => b.setAttribute("aria-checked", !!on));
const fileName = (path) => path.split(/[\\/]/).pop();
const plural = (n, word) => `${n.toLocaleString()} ${word}${n === 1 ? "" : "s"}`;

// Short stable key for a path (FNV-1a), used for version folders
function pathKey(path) {
  let h = 0x811c9dc5;
  for (const ch of path.toLowerCase()) h = Math.imul(h ^ ch.codePointAt(0), 0x01000193) >>> 0;
  return h.toString(16).padStart(8, "0");
}

function when(ts) {
  const d = new Date(ts);
  const days = Math.round((new Date().setHours(0, 0, 0, 0) - new Date(ts).setHours(0, 0, 0, 0)) / 86400000);
  const time = d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  if (days === 0) return `Today ${time}`;
  if (days === 1) return `Yesterday ${time}`;
  return `${d.toLocaleDateString([], { day: "numeric", month: "short", year: d.getFullYear() === new Date().getFullYear() ? undefined : "numeric" })} ${time}`;
}

function countWords() {
  // textContent, not innerText: off-screen paragraphs skip layout and innerText would miss them
  const text = [...pane.querySelectorAll(".docx article p")].map((p) => p.textContent).join(" ");
  words = (text.match(/[\p{L}\p{N}][\p{L}\p{N}'’-]*/gu) ?? []).length;
}

// ---------- Status line and writing goal ----------

function renderWords() {
  const goal = setting("goal", 0);
  const btn = $("words");
  btn.replaceChildren();
  if (goal && doc) {
    const written = Math.max(0, words - (doc.startWords ?? words));
    const done = Math.min(1, written / goal);
    const ns = "http://www.w3.org/2000/svg";
    const svg = document.createElementNS(ns, "svg");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.classList.add("ring");
    svg.classList.toggle("done", done >= 1);
    for (const [cls, len] of [
      ["track", 1],
      ["fill", done],
    ]) {
      const c = document.createElementNS(ns, "circle");
      c.setAttribute("cx", 8);
      c.setAttribute("cy", 8);
      c.setAttribute("r", 6);
      c.setAttribute("pathLength", 1);
      c.setAttribute("class", cls);
      c.style.strokeDasharray = `${len} 1`;
      svg.append(c);
    }
    btn.append(svg);
    btn.title = `${plural(written, "word")} of ${goal.toLocaleString()} this session`;
  } else btn.title = "Set a writing goal";
  btn.append(plural(words, "word"));
}

function updateUi() {
  const title = doc ? `${dirty ? "• " : ""}${doc.name} — Ribbonless` : "Ribbonless";
  document.title = title;
  if (isTauri) getCurrentWindow().setTitle(title);
  $("name").textContent = doc ? doc.name.replace(/\.docx$/i, "") : "Ribbonless";
  $("status").hidden = !doc;
  document.body.classList.toggle("no-doc", !doc); // the start screen has no outline
  $("state").textContent = !doc
    ? ""
    : previewing
      ? `Version from ${when(previewing.ts)}`
      : [!editor && "Read-only", dirty ? "Edited" : !doc.path && "Not saved"].filter(Boolean).join(" · ");
  renderWords();
  $("save").hidden = autoSaving() || !(dirty || (doc && !doc.path));
  for (const b of $$("[data-needs-doc]")) b.disabled = !doc;
  for (const b of $$("[data-needs-edit]")) b.disabled = !editor;
  for (const b of $$("[data-needs-path]")) b.disabled = !isTauri || !doc?.path;
  check("theme", document.documentElement.dataset.theme === "dark");
  check("autosave", setting("autosave"));
  check("outline", document.body.classList.contains("outline-open"));
  check("focus", document.body.classList.contains("focus-mode"));
  check("changes", setting("changes"));
  check("smoothCaret", setting("smoothCaret"));
  check("reopen", setting("reopen"));
  check("read", reader.isReading());
  document.body.classList.toggle("smooth-caret", setting("smoothCaret"));
  if (!editor) closePopover();
  syncFormat();
  showChanges();
  if (active) renderTabs();
}

// ---------- Rendering ----------

async function render(continuous = false, data = doc.data) {
  const pos = view.scrollTop / Math.max(1, view.scrollHeight - view.clientHeight);
  const next = document.createElement("div");
  const parsed = await renderAsync(data, next, null, {
    inWrapper: true,
    hideWrapperOnPrint: true,
    ignoreLastRenderedPageBreak: true,
    breakPages: !continuous,
  });
  // Give unstyled paragraphs the default style class; docx-preview's own rule misses the text inside them
  const normal = parsed?.stylesPart?.styles?.find((s) => s.isDefault && s.target === "p")?.cssName;
  if (normal) for (const p of next.querySelectorAll("p")) if (![...p.classList].some((c) => c.startsWith("docx_"))) p.classList.add(normal);
  pane.replaceChildren(...next.childNodes);
  view.scrollTop = pos * (view.scrollHeight - view.clientHeight);
  find.refresh();
}

// Pages when possible; a continuous page when page breaks split paragraphs; read-only as a last resort
async function renderEditable() {
  await render();
  editor = await createEditor(pane, doc.data);
  if (editor) return;
  await render(true);
  editor = await createEditor(pane, doc.data);
  if (!editor) await render();
}

// ---------- Format panel, selection bar ----------

function syncFormat() {
  const st = editor?.state() ?? { editable: false };
  const on = !!editor && st.editable;
  const has = (key) => !!editor?.styles.some((s) => s.key === key);
  for (const b of $$("[data-fmt]")) {
    b.disabled = !on;
    b.setAttribute("aria-pressed", !!st[{ bold: "b", italic: "i", underline: "u" }[b.dataset.fmt]]);
  }
  for (const b of $$("[data-key]")) {
    const style = editor?.styles.find((s) => s.key === b.dataset.key);
    b.disabled = !on || !style;
    b.setAttribute("aria-pressed", !!style && st.style === style.id);
  }
  for (const b of $$("[data-list]")) {
    b.disabled = !on || !has(b.dataset.list);
    b.setAttribute("aria-pressed", st.list === b.dataset.list);
  }
  for (const b of $$("[data-align]")) {
    b.disabled = !on;
    b.setAttribute("aria-pressed", st.align === b.dataset.align);
  }
  for (const b of $$("[data-step]")) b.disabled = !on;
  for (const b of $$("[data-color]")) {
    b.disabled = !on;
    b.setAttribute("aria-pressed", !!b.dataset.color && b.dataset.color.slice(1).toUpperCase() === st.color);
  }
  const custom = $("color-custom");
  const isCustom = on && !!st.color && !$$("[data-color]").some((b) => b.getAttribute("aria-pressed") === "true");
  custom.disabled = !on;
  custom.parentElement.setAttribute("aria-pressed", isCustom);
  custom.parentElement.style.setProperty("--c", isCustom ? `#${st.color}` : "");
  if (isCustom) custom.value = `#${st.color.toLowerCase()}`;
  for (const b of $$("[data-hl]")) {
    b.disabled = !on;
    b.setAttribute("aria-pressed", b.dataset.hl ? b.dataset.hl.slice(1).toUpperCase() === st.hl : on && !st.hl);
  }
  $$(".size-value")[0].textContent = st.size ? `${st.size} pt` : "—";
  for (const b of $$("[data-line]")) {
    b.disabled = !on;
    b.setAttribute("aria-pressed", on && +b.dataset.line === st.line);
  }
  for (const b of $$("[data-after]")) {
    b.disabled = !on;
    b.setAttribute("aria-pressed", on && +b.dataset.after === st.after);
  }
  for (const b of $$("[data-link]")) {
    b.disabled = !on;
    b.setAttribute("aria-pressed", !!st.href);
  }
  syncFontPicker(on ? st.font : undefined, on);
  for (const b of $$('[data-cmd="undo"]')) b.disabled = !editor?.canUndo();
  for (const b of $$('[data-cmd="redo"]')) b.disabled = !editor?.canRedo();
}

// ---------- Font picker ----------

const COMMON_FONTS = ["Aptos", "Calibri", "Cambria", "Georgia", "Segoe UI", "Arial", "Times New Roman", "Verdana", "Consolas"];
let installedFonts = null;

async function loadFonts() {
  if (installedFonts) return;
  installedFonts = isTauri ? await invoke("fonts").catch(() => []) : [];
  const select = $("font-select");
  const option = (name) => Object.assign(new Option(name, name), { style: `font-family: "${name}"` });
  const common = COMMON_FONTS.filter((f) => !isTauri || installedFonts.includes(f) || f === "Aptos");
  const rest = installedFonts.filter((f) => !common.includes(f) && !f.startsWith("@"));
  select.replaceChildren(
    new Option("Default font", ""),
    Object.assign(document.createElement("optgroup"), { label: "Common" }),
    Object.assign(document.createElement("optgroup"), { label: "All fonts" }),
  );
  select.children[1].append(...common.map(option));
  select.children[2].append(...rest.map(option));
}

function syncFontPicker(font, on) {
  const select = $("font-select");
  select.disabled = !on;
  if (font === undefined) {
    select.value = "";
    return;
  }
  if (font && ![...select.options].some((o) => o.value === font)) select.add(new Option(font, font), 1);
  select.value = font ?? "";
}

$("font-select").onfocus = loadFonts;
$("font-select").onmousedown = loadFonts;
// The font dropdown and color picker take focus; keep the text selection to restore before formatting
let panelSelection = null;
function formatFromField(kind, value) {
  if (!editor) return;
  if (panelSelection) {
    getSelection().removeAllRanges();
    getSelection().addRange(panelSelection);
  }
  editor.format(kind, value);
  syncFormat();
}
for (const id of ["font-select", "color-custom"])
  $(id).addEventListener("pointerdown", () => {
    const sel = getSelection();
    panelSelection = sel.rangeCount && !sel.isCollapsed ? sel.getRangeAt(0).cloneRange() : null;
  });
$("font-select").onchange = (e) => formatFromField("font", e.target.value || null);
$("color-custom").onchange = (e) => formatFromField("color", e.target.value.toUpperCase());

function openPopover() {
  const pop = $("format-pop");
  pop.hidden = false;
  $("aa").setAttribute("aria-expanded", true);
  const b = $("aa").getBoundingClientRect();
  pop.style.top = `${b.bottom + 8}px`;
  pop.style.left = `${Math.min(window.innerWidth - pop.offsetWidth - 10, Math.max(10, b.left + b.width / 2 - pop.offsetWidth / 2))}px`;
  syncFormat();
}
function closePopover() {
  $("format-pop").hidden = true;
  $("aa").setAttribute("aria-expanded", false);
}

let mouseDown = false;
let selTimer;
function placeSelbar() {
  const bar = $("selbar");
  const sel = getSelection();
  if (!editor || mouseDown || !sel.rangeCount || sel.isCollapsed || !editor.inDoc() || !$("format-pop").hidden) {
    bar.hidden = true;
    return;
  }
  const rects = [...sel.getRangeAt(0).getClientRects()].filter((r) => r.width);
  if (!rects.length) {
    bar.hidden = true;
    return;
  }
  const top = rects[0];
  const area = view.getBoundingClientRect();
  bar.hidden = false;
  const w = bar.offsetWidth;
  const h = bar.offsetHeight;
  const chrome = Math.max(
    ...[document.querySelector(".toolbar"), $("tabbar"), $("findbar")]
      .filter((el) => !el.hidden)
      .map((el) => el.getBoundingClientRect().bottom),
  );
  const above = top.top - h - 8;
  bar.style.top = `${above > chrome + 4 ? above : rects.at(-1).bottom + 8}px`;
  bar.style.left = `${Math.min(area.right - w - 8, Math.max(area.left + 8, top.left + top.width / 2 - w / 2))}px`;
}

function runFormat(el) {
  if (!editor || el.disabled) return;
  const d = el.dataset;
  if ("link" in d) return openLink();
  if ("line" in d) editor.setParagraphSpacing({ line: +d.line });
  else if ("after" in d) editor.setParagraphSpacing({ after: +d.after });
  else if ("color" in d) editor.format("color", d.color || null);
  else if ("hl" in d) editor.format("highlight", d.hl || null);
  else if (d.fmt) editor.format(d.fmt);
  else if (d.key) editor.setStyleByKey(d.key);
  else if (d.list) editor.toggleList(d.list);
  else if (d.align) editor.setAlign(d.align);
  else if (d.step) editor.stepSize(+d.step);
  syncFormat();
  placeSelbar();
}

for (const panel of [$("format-pop"), $("selbar")]) {
  // Clicking the panels must not take the text selection away (dropdowns still need the click)
  panel.onmousedown = (e) => {
    if (!e.target.closest("select, input")) e.preventDefault();
  };
  panel.onclick = (e) => {
    const el = e.target.closest("button");
    if (el) runFormat(el);
  };
}
$("aa").onclick = (e) => {
  e.stopPropagation();
  if ($("format-pop").hidden) {
    closeMenus();
    openPopover();
  } else closePopover();
  placeSelbar();
};
$("aa").onmousedown = (e) => e.preventDefault();

// ---------- Outline ----------

let headings = [];

function buildOutline() {
  headings = [...pane.querySelectorAll(".docx article p")].filter((p) => HEADING.test(p.className) && p.textContent.trim());
  $("outline-list").replaceChildren(
    ...headings.map((p, i) => {
      const [, kind, n] = p.className.match(HEADING);
      const b = document.createElement("button");
      b.className = `ol-item ol-${kind === "title" ? 0 : Math.min(+n, 3)}`;
      b.textContent = p.textContent.trim();
      b.onclick = () => goTo(p, i);
      return b;
    }),
  );
  $("outline-empty").hidden = headings.length > 0;
  trackOutline();
}

function goTo(p, i = headings.indexOf(p)) {
  p.scrollIntoView({ behavior: "smooth", block: "start" });
  // Skipped paragraphs get their real height as the scroll passes them, so settle exactly at the end
  const started = Date.now();
  view.addEventListener("scrollend", () => Date.now() - started < 2000 && p.scrollIntoView({ block: "start" }), { once: true });
  markOutline(i);
}

function markOutline(i) {
  [...$("outline-list").children].forEach((b, k) => b.classList.toggle("active", k === i));
}

// The heading whose section is at the top of the view
function trackOutline() {
  if (!document.body.classList.contains("outline-open") || !headings.length) return;
  const line = view.getBoundingClientRect().top + 140;
  let at = 0;
  headings.forEach((p, i) => {
    if (p.getBoundingClientRect().top <= line) at = i;
  });
  markOutline(at);
}

function toggleOutline(show = !document.body.classList.contains("outline-open")) {
  document.body.classList.toggle("outline-open", show);
  store.set("outline", show);
  if (show) buildOutline();
  updateUi();
}

// ---------- Changes since opened ----------

const changesStyle = document.head.appendChild(document.createElement("style"));

// Margin bars via a generated stylesheet, so the document itself is never touched
function showChanges() {
  const ids = setting("changes") && editor ? editor.changedIds() : [];
  changesStyle.textContent = ids.length
    ? `${ids.map((id) => `.docx article p[data-xid="${CSS.escape(id)}"]`).join(",")} { position: relative; }
${ids.map((id) => `.docx article p[data-xid="${CSS.escape(id)}"]::after`).join(",")} { content: ""; position: absolute; left: -16px; top: 1px; bottom: 1px; width: 3px; border-radius: 2px; background: var(--change); }`
    : "";
}

// ---------- Focus mode ----------

const focusStyle = document.head.appendChild(document.createElement("style"));

function markCurrentParagraph() {
  if (!document.body.classList.contains("focus-mode")) return (focusStyle.textContent = "");
  const node = getSelection().anchorNode;
  const p = node && (node.nodeType === 1 ? node : node.parentElement)?.closest("article p");
  const id = p?.dataset.xid;
  // A stylesheet rule instead of a class, so the document itself is never touched
  focusStyle.textContent = id ? `.focus-mode .docx article p:not([data-xid="${CSS.escape(id)}"]) { opacity: 0.28; }` : "";
}

async function toggleFocus(on = !document.body.classList.contains("focus-mode")) {
  document.body.classList.toggle("focus-mode", on);
  document.body.classList.remove("reveal");
  if (on) {
    closeMenus();
    closePopover();
    find.close();
  }
  if (isTauri)
    await getCurrentWindow()
      .setFullscreen(on)
      .catch(() => {});
  markCurrentParagraph();
  updateUi();
}

document.addEventListener("mousemove", (e) => {
  if (!document.body.classList.contains("focus-mode")) return;
  document.body.classList.toggle("reveal", e.clientY < 70);
});

// ---------- Remembered place, previews ----------

function rememberPlace() {
  if (!doc?.path || previewing) return;
  const places = store.get("places", {});
  places[doc.path] = { pos: view.scrollTop / Math.max(1, view.scrollHeight - view.clientHeight), zoom };
  const keys = Object.keys(places);
  if (keys.length > 60) delete places[keys[0]];
  store.set("places", places);
}

function restorePlace() {
  const place = doc?.path && store.get("places", {})[doc.path];
  if (!place) return;
  if (place.zoom && place.zoom !== zoom) setZoom(place.zoom, false);
  requestAnimationFrame(() => (view.scrollTop = place.pos * (view.scrollHeight - view.clientHeight)));
}

// A tiny stand-in of the first page for the start screen
function capturePreview() {
  if (!doc?.path) return;
  const first = pane.querySelector(".docx article");
  if (!first) return;
  const blocks = [];
  for (const el of first.children) {
    if (blocks.length >= 16) break;
    if (el.tagName === "TABLE") {
      blocks.push({ k: "t" });
      continue;
    }
    if (el.tagName !== "P") continue;
    const img = el.querySelector("img");
    if (img && !blocks.some((b) => b.k === "i")) {
      try {
        const w = 150;
        const h = Math.max(1, Math.round((img.naturalHeight / img.naturalWidth) * w));
        const c = document.createElement("canvas");
        c.width = w;
        c.height = h;
        c.getContext("2d").drawImage(img, 0, 0, w, h);
        blocks.push({ k: "i", src: c.toDataURL("image/jpeg", 0.6), ratio: h / w });
      } catch {}
      continue;
    }
    const text = el.textContent.trim().slice(0, 140);
    if (!text) continue;
    const h = el.className.match(HEADING);
    blocks.push({ k: h ? (h[1] === "title" ? "T" : "h") : [...el.classList].some((c) => c.startsWith("docx-num-")) ? "l" : "p", text });
  }
  const previews = store.get("previews", {});
  previews[doc.path] = blocks;
  const recent = new Set(store.get("recent", []));
  for (const k of Object.keys(previews)) if (!recent.has(k) && k !== doc.path) delete previews[k];
  store.set("previews", previews);
}

// ---------- Tabs ----------

function newTab() {
  const tab = {
    pane: document.createElement("div"),
    doc: null,
    editor: null,
    dirty: false,
    version: 0,
    words: 0,
    previewing: null,
    scroll: 0,
  };
  tab.pane.className = "tab-pane";
  view.append(tab.pane);
  tabs.push(tab);
  return tab;
}

function switchTab(tab) {
  if (tab === active) return;
  if (saveTimer) autosaveNow(); // don't leave the outgoing tab's edits waiting on a timer
  if (active) {
    Object.assign(active, { doc, editor, dirty, version, words, previewing, scroll: view.scrollTop });
    active.pane.hidden = true;
    // A start screen left behind is not worth a tab
    if (!doc && tabs.length > 1) {
      tabs.splice(tabs.indexOf(active), 1);
      active.pane.remove();
    }
  }
  rememberPlace();
  reader.stop();
  find.close();
  imageTools.deselect();
  closeHistory();
  active = tab;
  ({ doc, editor, dirty, version, words, previewing } = tab);
  pane = tab.pane;
  pane.hidden = false;
  view.scrollTop = tab.scroll;
  buildOutline();
  renderTabs();
  updateUi();
}

// Opens a document in a new tab, or in this one when it is empty; an already open file just gets focus
async function openDocument(name, data, path = null) {
  const open = path && tabs.find((t) => (t === active ? doc : t.doc)?.path === path);
  if (open) return switchTab(open);
  if (doc) switchTab(newTab());
  return load(name, data, path, { force: true });
}

async function closeTab(tab = active, { discard = false } = {}) {
  if (tab !== active) switchTab(tab);
  if (doc) {
    if (previewing) await exitPreview();
    if (autoSaving() && dirty) await save();
    if (dirty && !discard && !(await confirmUser(`Close “${doc.name}” without saving changes?`))) return false;
    rememberPlace();
  }
  const i = tabs.indexOf(tab);
  tabs.splice(i, 1);
  tab.pane.remove();
  active = null;
  if (!tabs.some(tabDirty)) clearBackup();
  if (!tabs.length) {
    switchTab(newTab());
    showStart();
  } else switchTab(tabs[Math.min(i, tabs.length - 1)]);
  saveSession();
  return true;
}

// Saves what auto-save can, then asks once about everything still unsaved
async function closeTabs(list) {
  for (const t of list) {
    if (!tabDirty(t)) continue;
    switchTab(t);
    if (previewing) await exitPreview();
    if (autoSaving() && dirty) await save();
  }
  const names = list.filter(tabDirty).map((t) => `• ${(t === active ? doc : t.doc).name}`);
  if (names.length) {
    const question =
      names.length === 1
        ? `${names[0].slice(2)} has unsaved changes. Close it anyway?`
        : `${names.length} documents have unsaved changes:\n\n${names.join("\n")}\n\nClose them anyway?`;
    if (!(await confirmUser(question))) return false;
  }
  for (const t of list) await closeTab(t, { discard: true });
  return true;
}

function cycleTab(delta) {
  if (tabs.length > 1) switchTab(tabs[(tabs.indexOf(active) + delta + tabs.length) % tabs.length]);
}

// The tab strip shows once two or more tabs are open
function renderTabs() {
  const bar = $("tabbar");
  document.body.classList.toggle("tabs-on", tabs.length > 1);
  bar.hidden = tabs.length < 2;
  bar.querySelector(".tabs").replaceChildren(
    ...tabs.map((t) => {
      const d = t === active ? doc : t.doc;
      const isDirty = tabDirty(t);
      const b = document.createElement("div");
      b.className = "tab";
      b.classList.toggle("active", t === active);
      b.title = d?.path ?? d?.name ?? "Start";
      b.onmousedown = (e) => {
        if (e.button === 1) {
          e.preventDefault();
          closeTab(t);
        }
      };
      b.onclick = (e) => {
        if (!e.target.closest(".tab-close")) switchTab(t);
      };
      b.oncontextmenu = (e) => {
        e.preventDefault();
        e.stopPropagation();
        openTabMenu(t, e);
      };
      const label = Object.assign(document.createElement("span"), {
        className: "tab-name",
        textContent: d ? d.name.replace(/\.docx$/i, "") : "Start",
      });
      const close = Object.assign(document.createElement("button"), {
        className: "tab-close",
        title: "Close tab (Ctrl+W)",
        ariaLabel: "Close tab",
      });
      close.classList.toggle("dirty", !!isDirty);
      close.onclick = () => closeTab(t);
      b.append(label, close);
      return b;
    }),
  );
}

// Right-click on a tab
let menuTab = null;
function openTabMenu(tab, e) {
  closeMenus();
  menuTab = tab;
  const menu = $("tab-menu");
  const i = tabs.indexOf(tab);
  menu.querySelector('[data-tabs="others"]').disabled = tabs.length < 2;
  menu.querySelector('[data-tabs="right"]').disabled = i === tabs.length - 1;
  menu.hidden = false;
  menu.style.left = `${Math.min(e.clientX, window.innerWidth - menu.offsetWidth - 6)}px`;
  menu.style.top = `${Math.min(e.clientY, window.innerHeight - menu.offsetHeight - 6)}px`;
}
$("tab-menu").onmousedown = (e) => e.preventDefault();
$("tab-menu").onclick = (e) => {
  const b = e.target.closest("[data-tabs]");
  if (!b || b.disabled || !tabs.includes(menuTab)) return;
  closeMenus();
  const i = tabs.indexOf(menuTab);
  const pick = {
    close: () => [menuTab],
    others: () => tabs.filter((t) => t !== menuTab),
    right: () => tabs.slice(i + 1),
    all: () => [...tabs],
  }[b.dataset.tabs]();
  closeTabs(pick);
};

// Open files are remembered so the next launch can bring them back
function saveSession() {
  const paths = tabs.map((t) => (t === active ? doc : t.doc)?.path).filter(Boolean);
  store.set("session", { paths, active: active === null ? null : (doc?.path ?? null) });
}

// ---------- Opening ----------

async function load(name, data, path = null, { force = false } = {}) {
  if (dirty && !force && !(await confirmUser("Discard unsaved changes?"))) return false;
  if (dirty) clearBackup();
  rememberPlace();
  reader.stop();
  closeHistory();
  try {
    doc = { name, path, data, original: data };
    editor = null;
    dirty = false;
    previewing = null;
    view.scrollTop = 0;
    await renderEditable();
    countWords();
    doc.startWords = words;
    buildOutline();
    if (path) {
      addRecent(path);
      store.set("last", path);
      restorePlace();
      setTimeout(capturePreview, 300);
    }
    saveSession();
  } catch (e) {
    doc = editor = null;
    showStart(`Could not open ${name}: ${e.message ?? e}`);
  }
  updateUi();
  return true;
}

async function loadPath(path, { quiet = false } = {}) {
  let data;
  try {
    data = await invoke("read_file", { path });
  } catch (e) {
    store.set(
      "recent",
      store.get("recent", []).filter((r) => r !== path),
    );
    if (!doc) showStart();
    if (!quiet) tell(`Could not open ${path}: ${e}`);
    return;
  }
  await openDocument(fileName(path), data, path);
}

async function pickFile() {
  if (!isTauri) return $("file").click();
  const path = await open({ filters: [{ name: "Word document", extensions: ["docx"] }] });
  if (path) loadPath(path);
}

async function newDocument() {
  if (await openDocument("Untitled.docx", await blankDocx())) editor?.focusStart();
}

function addRecent(path) {
  store.set("recent", [path, ...store.get("recent", []).filter((r) => r !== path)].slice(0, 8));
  store.set("opened", { ...store.get("opened", {}), [path]: Date.now() });
}

// A tiny first page drawn from preview blocks (T title, h/h2 heading, p body, l/n list, t table, i image)
function miniPage(blocks) {
  const page = document.createElement("div");
  page.className = "mini-page";
  for (const b of blocks) {
    if (b.k === "i") {
      const img = new Image();
      img.src = b.src;
      img.style.aspectRatio = String(1 / b.ratio);
      page.append(img);
    } else if (b.k === "t") {
      page.append(Object.assign(document.createElement("div"), { className: "m-table" }));
    } else {
      const k = { h2: "h", n: "l" }[b.k] ?? b.k;
      page.append(Object.assign(document.createElement("div"), { className: `m-${k}`, textContent: b.text || "\u00a0" }));
    }
  }
  return page;
}

function card(page, title, meta, onclick, tip = "") {
  const b = document.createElement("button");
  b.className = "card";
  b.title = tip;
  b.onclick = onclick;
  b.append(
    page,
    Object.assign(document.createElement("span"), { className: "card-name", textContent: title }),
    Object.assign(document.createElement("span"), { className: "card-meta", textContent: meta }),
  );
  return b;
}

function previewCard(path) {
  const opened = store.get("opened", {})[path];
  return card(
    miniPage(store.get("previews", {})[path] ?? []),
    fileName(path).replace(/\.docx$/i, ""),
    opened ? when(opened) : path.slice(0, -fileName(path).length - 1),
    () => loadPath(path),
    path,
  );
}

// ---------- Links ----------

let linkTarget = null; // the range the link popover applies to
const hrefOf = (el) => el.dataset.href ?? el.getAttribute("href");

function openLink() {
  if (!editor) return;
  const existing = editor.linkRange();
  const sel = getSelection();
  if (existing) linkTarget = existing.range;
  else if (sel.rangeCount && !sel.isCollapsed && editor.inDoc()) linkTarget = sel.getRangeAt(0).cloneRange();
  else return hud("Select some text first");
  const pop = $("link-pop");
  $("link-input").value = existing?.href ?? "";
  $("link-remove").hidden = !existing;
  pop.hidden = false;
  const r = linkTarget.getBoundingClientRect();
  pop.style.top = `${Math.min(window.innerHeight - pop.offsetHeight - 10, r.bottom + 8)}px`;
  pop.style.left = `${Math.max(10, Math.min(window.innerWidth - pop.offsetWidth - 10, r.left))}px`;
  $("selbar").hidden = true;
  $("link-input").focus();
  $("link-input").select();
}

function closeLink() {
  $("link-pop").hidden = true;
  linkTarget = null;
}

function applyLink(remove = false) {
  if (!linkTarget || !editor) return closeLink();
  let url = $("link-input").value.trim();
  if (url && !remove && !/^([a-z][\w+.-]*:|#)/i.test(url))
    url = url.includes("@") && !url.includes("/") ? `mailto:${url}` : `https://${url}`;
  getSelection().removeAllRanges();
  getSelection().addRange(linkTarget);
  editor.format("link", remove || !url ? null : url);
  closeLink();
  getSelection().collapseToEnd();
}

$("link-input").onkeydown = (e) => {
  if (e.key === "Enter") {
    e.preventDefault();
    applyLink();
  }
  if (e.key === "Escape") {
    e.stopPropagation();
    closeLink();
  }
};
$("link-apply").onclick = () => applyLink();
$("link-remove").onclick = () => applyLink(true);
$("link-pop").onmousedown = (e) => {
  if (e.target.tagName !== "INPUT") e.preventDefault();
};

// Links never navigate the app window; Ctrl+click opens them in the browser
async function followLink(href) {
  if (href.startsWith("#")) return document.getElementById(href.slice(1))?.scrollIntoView({ behavior: "smooth", block: "start" });
  if (!/^(https?:|mailto:)/i.test(href)) return;
  if (isTauri) {
    const { openUrl } = await import("@tauri-apps/plugin-opener");
    openUrl(href).catch(() => {});
  } else window.open(href, "_blank", "noopener");
}

view.addEventListener("click", (e) => {
  const link = e.target.closest?.("a[href], [data-href]");
  if (!link || !pane.contains(link)) return;
  e.preventDefault();
  if (e.ctrlKey || !editor) followLink(hrefOf(link));
});

// A small tooltip with the address, instead of a title attribute that would change the document
let tipTimer;
view.addEventListener("mouseover", (e) => {
  const link = e.target.closest?.("a[href], [data-href]");
  clearTimeout(tipTimer);
  if (!link) {
    $("link-tip").hidden = true;
    return;
  }
  tipTimer = setTimeout(() => {
    const tip = $("link-tip");
    tip.querySelector("span").textContent = hrefOf(link);
    tip.querySelector("kbd").textContent = editor ? "Ctrl+click to open" : "Click to open";
    tip.hidden = false;
    const r = link.getBoundingClientRect();
    tip.style.top = `${r.bottom + 6}px`;
    tip.style.left = `${Math.max(8, Math.min(window.innerWidth - tip.offsetWidth - 8, r.left))}px`;
  }, 450);
});

async function newFromTemplate(t) {
  if (await openDocument(`${t.name}.docx`, await blankDocx(t.blocks))) editor?.focusStart();
}

function showStart(error) {
  const empty = $("start").content.cloneNode(true);
  empty.querySelector(".recent").append(...TEMPLATES.map((t) => card(miniPage(t.blocks), t.name, "Template", () => newFromTemplate(t))));
  const recent = isTauri ? store.get("recent", []) : [];
  if (recent.length) {
    const head = Object.assign(document.createElement("h2"), { className: "start-section", textContent: "Recent" });
    const grid = Object.assign(document.createElement("div"), { className: "recent" });
    grid.append(...recent.map(previewCard));
    empty.querySelector(".recent-wrap").append(head, grid);
  }
  empty.querySelector(".error").textContent = error ?? "";
  empty.querySelector(".open-btn").onclick = pickFile;
  empty.querySelector(".new-btn").onclick = newDocument;
  pane.replaceChildren(empty);
}

// ---------- Saving and versions ----------

// These take the document up front: the user may switch tabs while they await
async function commit(d = doc, ed = editor) {
  if (!ed) return;
  const data = await ed.commit();
  if (data) d.data = data;
}

// Keeps a copy in the version history: the original once per session, then at most every two minutes
async function snapshot(force = false, d = doc) {
  if (!isTauri || !d?.path) return;
  const key = pathKey(d.path);
  const now = Date.now();
  try {
    if (!d.snapshotted) {
      d.snapshotted = true;
      const times = await invoke("version_list", { key });
      const latest = times.length && new Uint8Array(await invoke("version_read", { key, ts: times.at(-1) }));
      const orig = new Uint8Array(d.original instanceof ArrayBuffer ? d.original : await new Blob([d.original]).arrayBuffer());
      const same = latest && latest.length === orig.length && latest.every((b, i) => b === orig[i]);
      if (!same) await raw("version_save", orig, { key, ts: String(now - 1) });
    }
    if (!force && now - (d.lastSnap ?? 0) < SNAPSHOT_EVERY) {
      d.unsnapped = true;
      return;
    }
    await raw("version_save", d.data, { key, ts: String(now) });
    d.lastSnap = now;
    d.unsnapped = false;
  } catch {}
}

// Tab state lives in globals for the active tab and on the tab object for the others
const tabDirty = (tab) => (tab === active ? dirty : tab.dirty);

async function save(as = false) {
  const tab = active;
  const d = doc;
  if (!d || previewing) return;
  const at = version;
  await commit(d, editor);
  let path = d.path;
  if (isTauri && (as || !path)) {
    path = await saveDialog({ defaultPath: d.path ?? d.name, filters: [{ name: "Word document", extensions: ["docx"] }] });
    if (!path) return;
    if (!/\.docx$/i.test(path)) path += ".docx";
  }
  try {
    if (isTauri) {
      await raw("write_file", d.data, { path });
      if (path !== d.path) {
        d.original = d.data;
        d.snapshotted = false;
      }
      d.path = path;
      d.name = fileName(path);
      addRecent(path);
      snapshot(false, d);
      if (tab === active) capturePreview();
    } else {
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([d.data]));
      a.download = d.name;
      a.click();
    }
    const stillDirty = (tab === active ? version : tab.version) !== at;
    if (tab === active) dirty = stillDirty;
    else tab.dirty = stillDirty;
    if (!tabs.some(tabDirty)) clearBackup();
  } catch (e) {
    tell(`Could not save: ${e}`);
  }
  updateUi();
  renderTabs();
}

// Auto-save writes to the file shortly after typing stops; untitled work goes to a backup instead
let saveTimer = null;
function scheduleSave() {
  if (!isTauri) return;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(autosaveNow, autoSaving() ? 1500 : 3000);
}
// Reads the active tab before its first await, so it is safe to call right before switching tabs
async function autosaveNow() {
  clearTimeout(saveTimer);
  saveTimer = null;
  if (!doc || !dirty) return;
  if (autoSaving()) return save();
  const d = doc;
  await commit(d, editor);
  await raw("backup_write", d.data, { meta: JSON.stringify({ name: d.name, path: d.path }) }).catch(() => {});
}
function clearBackup() {
  if (isTauri) invoke("backup_clear").catch(() => {});
}

async function revert() {
  if (!doc?.path || !doc.original) return;
  if (!(await confirmUser("Revert to the version you opened? Your current text stays in Version History."))) return;
  await commit();
  await snapshot(true);
  const { name, path, original } = doc;
  if (isTauri) await raw("write_file", original, { path });
  await load(name, original, path, { force: true });
}

async function openHistory() {
  if (!isTauri || !doc?.path) return;
  await commit();
  await snapshot(true);
  const times = (await invoke("version_list", { key: pathKey(doc.path) }).catch(() => [])).reverse();
  const list = $("history-list");
  const item = (label, sub, onclick, cls = "") => {
    const b = document.createElement("button");
    b.className = `hv-item ${cls}`;
    b.append(
      Object.assign(document.createElement("span"), { textContent: label }),
      Object.assign(document.createElement("small"), { textContent: sub }),
    );
    b.onclick = () => {
      for (const x of list.children) x.classList.remove("active");
      b.classList.add("active");
      onclick();
    };
    return b;
  };
  list.replaceChildren(
    item("Current", "What you're editing now", () => exitPreview(), "active"),
    ...times.map((ts, i) => item(when(ts), i === times.length - 1 ? "Oldest kept" : "", () => previewVersion(ts))),
  );
  $("history-empty").hidden = times.length > 0;
  $("history").hidden = false;
  document.body.classList.add("history-open");
}

function closeHistory() {
  if ($("history").hidden) return;
  $("history").hidden = true;
  document.body.classList.remove("history-open");
  if (previewing) exitPreview();
}

async function previewVersion(ts) {
  if (!doc?.path) return;
  if (!previewing) await commit();
  const data = await invoke("version_read", { key: pathKey(doc.path), ts });
  previewing = { ts, data };
  editor = null;
  await render(false, data);
  $("version-bar-when").textContent = when(ts);
  $("version-bar").hidden = false;
  buildOutline();
  updateUi();
}

async function exitPreview() {
  if (!previewing) return;
  previewing = null;
  $("version-bar").hidden = true;
  await renderEditable();
  buildOutline();
  updateUi();
}

async function restoreVersion() {
  if (!previewing) return;
  const { data } = previewing;
  previewing = null;
  $("version-bar").hidden = true;
  await snapshot(true); // the text being replaced stays in the history
  doc.data = data;
  await renderEditable();
  version++;
  dirty = true;
  closeHistory();
  countWords();
  buildOutline();
  updateUi();
  scheduleSave();
}

// ---------- Insert ----------

// Any image becomes PNG/JPEG/GIF bytes plus its size, ready for the editor
async function insertImageData(bytes, name) {
  if (!editor) return;
  let ext = (name.match(/\.(\w+)$/)?.[1] ?? "png").toLowerCase().replace("jpg", "jpeg");
  let data = new Uint8Array(bytes);
  const bitmap = await createImageBitmap(new Blob([data])).catch(() => null);
  if (!bitmap) return tell(`Could not read ${name} as an image.`);
  if (!["png", "jpeg", "gif"].includes(ext)) {
    const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
    canvas.getContext("2d").drawImage(bitmap, 0, 0);
    data = new Uint8Array(await (await canvas.convertToBlob({ type: "image/png" })).arrayBuffer());
    ext = "png";
  }
  editor.insertImage(data, ext, bitmap.width, bitmap.height);
}

async function pickImage() {
  if (!editor) return;
  if (!isTauri) return $("image-file").click();
  const path = await open({ filters: [{ name: "Image", extensions: ["png", "jpg", "jpeg", "gif", "webp", "bmp"] }] });
  if (path) insertImageData(await invoke("read_file", { path }), path);
}

// ---------- Export ----------

async function exportAs(kind) {
  if (!doc) return;
  if (previewing) await exitPreview();
  const base = doc.name.replace(/\.docx$/i, "");
  const text = kind === "md" ? toMarkdown(pane) : toHtml(pane, base);
  const bytes = new TextEncoder().encode(text);
  if (!isTauri) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([bytes], { type: kind === "md" ? "text/markdown" : "text/html" }));
    a.download = `${base}.${kind}`;
    a.click();
    return;
  }
  const dir = doc.path ? doc.path.slice(0, -fileName(doc.path).length) : "";
  const path = await saveDialog({
    defaultPath: `${dir}${base}.${kind}`,
    filters: [kind === "md" ? { name: "Markdown", extensions: ["md"] } : { name: "Web page", extensions: ["html"] }],
  });
  if (!path) return;
  try {
    await raw("write_file", bytes, { path });
    hud(kind === "md" ? "Exported Markdown" : "Exported HTML");
  } catch (e) {
    tell(`Could not export: ${e}`);
  }
}

// ---------- Zoom, theme ----------

let hudTimer;
function hud(text) {
  const el = $("hud");
  el.textContent = text;
  el.hidden = false;
  el.classList.remove("fade");
  clearTimeout(hudTimer);
  hudTimer = setTimeout(() => el.classList.add("fade"), 900);
}

function setZoom(z, announce = true) {
  zoom = Math.min(3, Math.max(0.3, Math.round(z * 10) / 10));
  view.style.setProperty("--zoom", zoom);
  if (announce) hud(`${Math.round(zoom * 100)}%`);
  placeSelbar();
  caret.update();
}

const systemDark = matchMedia("(prefers-color-scheme: dark)");
function applyTheme() {
  const chosen = store.get("theme", null);
  document.documentElement.dataset.theme = chosen ?? (systemDark.matches ? "dark" : "light");
  // Mica follows the window theme, so keep it in step with ours
  if (isTauri)
    getCurrentWindow()
      .setTheme(chosen)
      .catch(() => {});
  check("theme", document.documentElement.dataset.theme === "dark");
}
systemDark.onchange = applyTheme;

// ---------- Goal popover ----------

function openGoal() {
  if (!doc) return;
  const pop = $("goal-pop");
  const goal = setting("goal", 0);
  $("goal-input").value = goal || "";
  const written = Math.max(0, words - (doc.startWords ?? words));
  $("goal-progress").textContent = goal
    ? `${plural(written, "word")} of ${goal.toLocaleString()} written this session`
    : `${plural(written, "word")} written this session`;
  pop.hidden = false;
  const b = $("words").getBoundingClientRect();
  pop.style.top = `${b.bottom + 8}px`;
  pop.style.left = `${Math.max(10, b.left - 12)}px`;
  $("goal-input").focus();
  $("goal-input").select();
}
function closeGoal() {
  $("goal-pop").hidden = true;
}
function setGoal(value) {
  store.set("goal", Math.max(0, Math.round(+value || 0)));
  closeGoal();
  updateUi();
}
$("words").onclick = (e) => {
  e.stopPropagation();
  $("goal-pop").hidden ? openGoal() : closeGoal();
};
$("goal-set").onclick = () => setGoal($("goal-input").value);
$("goal-clear").onclick = () => setGoal(0);
$("goal-input").onkeydown = (e) => {
  if (e.key === "Enter") setGoal(e.target.value);
  if (e.key === "Escape") closeGoal();
};

// ---------- Commands ----------

const reader = createReader(view, {
  onChange: (on) => {
    $("reader").hidden = !on;
    updateUi();
  },
  root: () => pane,
});
$("reader-stop").onclick = () => reader.stop();

const commands = {
  new: newDocument,
  open: pickFile,
  save: () => save(),
  saveAs: () => save(true),
  revert,
  history: () => ($("history").hidden ? openHistory() : closeHistory()),
  print: async () => {
    if (!doc) return;
    if (previewing) await exitPreview();
    window.print();
  },
  exportMd: () => exportAs("md"),
  exportHtml: () => exportAs("html"),
  undo: () => {
    editor?.undo();
    syncFormat();
  },
  redo: () => {
    editor?.redo();
    syncFormat();
  },
  find: () => doc && find.open(),
  replace: () => doc && find.open(true),
  // The start screen opens beside the document; its tab is one click back
  home: async () => {
    if (!doc) return;
    if (document.body.classList.contains("focus-mode")) await toggleFocus(false);
    commands.newTab();
  },
  newTab: () => {
    switchTab(newTab());
    showStart();
    updateUi();
  },
  closeTab: () => closeTab(),
  closeAllTabs: () => closeTabs([...tabs]),
  closeOtherTabs: () => closeTabs(tabs.filter((t) => t !== active)),
  nextTab: () => cycleTab(1),
  prevTab: () => cycleTab(-1),
  link: () => openLink(),
  palette: () => palette.open(),
  image: pickImage,
  table: () => editor?.insertTable(),
  outline: () => toggleOutline(),
  focus: () => toggleFocus(),
  read: () => doc && reader.toggle(),
  goal: openGoal,
  zoomIn: () => setZoom(zoom + 0.1),
  zoomOut: () => setZoom(zoom - 0.1),
  zoomReset: () => setZoom(1),
  autosave: () => {
    store.set("autosave", !setting("autosave"));
    updateUi();
    if (dirty) scheduleSave();
  },
  changes: () => {
    store.set("changes", !setting("changes"));
    updateUi();
  },
  smoothCaret: () => {
    store.set("smoothCaret", !setting("smoothCaret"));
    updateUi();
    caret.update();
  },
  reopen: () => {
    store.set("reopen", !setting("reopen"));
    updateUi();
  },
  theme: () => {
    store.set("theme", document.documentElement.dataset.theme === "dark" ? "light" : "dark");
    applyTheme();
  },
};

// Everything the command bar can do, in plain words
function paletteItems() {
  const ed = !!editor;
  const has = (key) => !!editor?.styles.some((s) => s.key === key);
  const style = (key) => () => editor?.setStyleByKey(key);
  const items = [
    { group: "Format", title: "Title", run: style("title"), on: ed && has("title") },
    { group: "Format", title: "Heading 1", hint: "Ctrl+Alt+1", run: style("heading 1"), on: ed && has("heading 1") },
    { group: "Format", title: "Heading 2", hint: "Ctrl+Alt+2", run: style("heading 2"), on: ed && has("heading 2") },
    { group: "Format", title: "Heading 3", hint: "Ctrl+Alt+3", run: style("heading 3"), on: ed && has("heading 3") },
    { group: "Format", title: "Body text", keywords: "normal paragraph", hint: "Ctrl+Shift+N", run: style("normal"), on: ed },
    {
      group: "Format",
      title: "Bulleted list",
      hint: "Ctrl+Shift+L",
      run: () => editor.toggleList("list bullet"),
      on: ed && has("list bullet"),
    },
    { group: "Format", title: "Numbered list", run: () => editor.toggleList("list number"), on: ed && has("list number") },
    { group: "Format", title: "Bold", hint: "Ctrl+B", run: () => editor.format("bold"), on: ed },
    { group: "Format", title: "Italic", hint: "Ctrl+I", run: () => editor.format("italic"), on: ed },
    { group: "Format", title: "Underline", hint: "Ctrl+U", run: () => editor.format("underline"), on: ed },
    { group: "Format", title: "Align left", hint: "Ctrl+L", run: () => editor.setAlign("left"), on: ed },
    { group: "Format", title: "Center", keywords: "align middle", hint: "Ctrl+E", run: () => editor.setAlign("center"), on: ed },
    { group: "Format", title: "Align right", hint: "Ctrl+R", run: () => editor.setAlign("right"), on: ed },
    { group: "Format", title: "Justify", hint: "Ctrl+J", run: () => editor.setAlign("justify"), on: ed },
    { group: "Format", title: "Bigger text", keywords: "font size increase", hint: "Ctrl+]", run: () => editor.stepSize(1), on: ed },
    { group: "Format", title: "Smaller text", keywords: "font size decrease", hint: "Ctrl+[", run: () => editor.stepSize(-1), on: ed },
    { group: "Insert", title: "Insert image", keywords: "picture photo", run: commands.image, on: ed },
    { group: "Insert", title: "Insert table", run: commands.table, on: ed },
    { group: "File", title: "Home", keywords: "start screen back recent templates", hint: "Alt+Home", run: commands.home, on: !!doc },
    { group: "File", title: "New tab", hint: "Ctrl+T", run: commands.newTab, on: true },
    { group: "File", title: "Close tab", hint: "Ctrl+W", run: commands.closeTab, on: true },
    { group: "File", title: "Close other tabs", run: commands.closeOtherTabs, on: tabs.length > 1 },
    { group: "File", title: "Close all tabs", hint: "Ctrl+Shift+W", run: commands.closeAllTabs, on: tabs.length > 1 || !!doc },
    { group: "Insert", title: "Add link…", keywords: "hyperlink url web", hint: "Ctrl+Shift+K", run: commands.link, on: ed },
    { group: "Format", title: "Line spacing 1.0", keywords: "single", run: () => editor.setParagraphSpacing({ line: 1 }), on: ed },
    { group: "Format", title: "Line spacing 1.15", run: () => editor.setParagraphSpacing({ line: 1.15 }), on: ed },
    { group: "Format", title: "Line spacing 1.5", run: () => editor.setParagraphSpacing({ line: 1.5 }), on: ed },
    { group: "Format", title: "Line spacing 2.0", keywords: "double", run: () => editor.setParagraphSpacing({ line: 2 }), on: ed },
    { group: "Format", title: "Indent list item", keywords: "nest level", hint: "Tab", run: () => editor.shiftLevel(1), on: ed },
    { group: "Format", title: "Outdent list item", keywords: "level", hint: "Shift+Tab", run: () => editor.shiftLevel(-1), on: ed },
    ...tabs
      .filter((t) => t !== active && t.doc)
      .map((t) => ({
        group: "Tabs",
        title: t.doc.name.replace(/\.docx$/i, ""),
        keywords: "switch tab",
        run: () => switchTab(t),
        on: true,
      })),
    { group: "File", title: "New document", hint: "Ctrl+N", run: commands.new, on: true },
    { group: "File", title: "Open…", hint: "Ctrl+O", run: commands.open, on: true },
    { group: "File", title: "Save", hint: "Ctrl+S", run: commands.save, on: !!doc },
    { group: "File", title: "Save as…", hint: "Ctrl+Shift+S", run: commands.saveAs, on: !!doc },
    { group: "File", title: "Export as PDF…", keywords: "print", hint: "Ctrl+P", run: commands.print, on: !!doc },
    { group: "File", title: "Export as Markdown…", keywords: "md", run: commands.exportMd, on: !!doc },
    { group: "File", title: "Export as HTML…", keywords: "web page", run: commands.exportHtml, on: !!doc },
    { group: "File", title: "Version history", keywords: "restore older time machine", run: commands.history, on: isTauri && !!doc?.path },
    { group: "File", title: "Revert to last opened", run: commands.revert, on: isTauri && !!doc?.path },
    { group: "Edit", title: "Undo", hint: "Ctrl+Z", run: commands.undo, on: ed },
    { group: "Edit", title: "Redo", hint: "Ctrl+Y", run: commands.redo, on: ed },
    { group: "Edit", title: "Find", keywords: "search", hint: "Ctrl+F", run: commands.find, on: !!doc },
    { group: "Edit", title: "Find and replace", keywords: "search change", hint: "Ctrl+H", run: commands.replace, on: ed },
    { group: "Format", title: "Highlight", keywords: "marker yellow", run: () => editor.format("highlight", "#FFFF00"), on: ed },
    { group: "Format", title: "Remove highlight", run: () => editor.format("highlight", null), on: ed },
    { group: "Format", title: "Text color: red", keywords: "colour", run: () => editor.format("color", "#FF3B30"), on: ed },
    { group: "Format", title: "Text color: blue", keywords: "colour", run: () => editor.format("color", "#007AFF"), on: ed },
    { group: "Format", title: "Text color: automatic", keywords: "colour reset black", run: () => editor.format("color", null), on: ed },
    ...TEMPLATES.map((t) => ({ group: "File", title: `New from template: ${t.name}`, run: () => newFromTemplate(t), on: true })),
    { group: "Edit", title: "Select all", hint: "Ctrl+A", run: () => editor.selectAll(), on: ed },
    {
      group: "View",
      title: reader.isReading() ? "Stop reading" : "Read aloud",
      keywords: "speak voice listen",
      run: commands.read,
      on: !!doc,
    },
    { group: "View", title: "Outline", keywords: "sidebar headings navigation", hint: "Ctrl+Shift+O", run: commands.outline, on: !!doc },
    { group: "View", title: "Focus mode", keywords: "distraction free fullscreen", hint: "Ctrl+Shift+F", run: commands.focus, on: !!doc },
    { group: "View", title: "Writing goal…", keywords: "words target", run: commands.goal, on: !!doc },
    {
      group: "View",
      title: setting("changes") ? "Hide changes" : "Show changes",
      keywords: "edits margin",
      run: commands.changes,
      on: !!doc,
    },
    { group: "View", title: "Dark appearance", keywords: "theme light night", run: commands.theme, on: true },
    { group: "View", title: "Zoom in", hint: "Ctrl++", run: commands.zoomIn, on: true },
    { group: "View", title: "Zoom out", hint: "Ctrl+−", run: commands.zoomOut, on: true },
    { group: "View", title: "Actual size", hint: "Ctrl+0", run: commands.zoomReset, on: true },
    ...headings.map((p) => ({
      group: "Go to",
      title: p.textContent.trim().slice(0, 80),
      keywords: "heading section",
      run: () => goTo(p),
      on: true,
    })),
    ...(isTauri ? store.get("recent", []) : [])
      .filter((p) => p !== doc?.path)
      .map((p) => ({ group: "Open recent", title: fileName(p).replace(/\.docx$/i, ""), keywords: p, run: () => loadPath(p), on: true })),
  ];
  return items.filter((i) => i.on);
}

const palette = createPalette({ getItems: paletteItems });

const slash = createSlashMenu({
  editableP: (node) => (editor ? (node?.nodeType === 1 ? node : node?.parentElement)?.closest("article p") : null),
  getItems: () => {
    const has = (key) => !!editor?.styles.some((s) => s.key === key);
    const style = (key) => () => editor.setStyleByKey(key);
    return [
      { title: "Heading 1", keywords: "h1 title big", run: style("heading 1"), on: has("heading 1") },
      { title: "Heading 2", keywords: "h2 subheading", run: style("heading 2"), on: has("heading 2") },
      { title: "Heading 3", keywords: "h3 section", run: style("heading 3"), on: has("heading 3") },
      { title: "Title", run: style("title"), on: has("title") },
      { title: "Body text", keywords: "normal paragraph", run: style("normal"), on: true },
      { title: "Bulleted list", keywords: "ul bullets", run: () => editor.toggleList("list bullet"), on: has("list bullet") },
      { title: "Numbered list", keywords: "ol numbers", run: () => editor.toggleList("list number"), on: has("list number") },
      { title: "Table", run: () => editor.insertTable(), on: true },
      { title: "Image", keywords: "picture photo", run: pickImage, on: true },
    ].filter((i) => i.on);
  },
});

const caret = createCaret(() => setting("smoothCaret") && !!editor);
const imageTools = createImageTools({ view, getEditor: () => editor });

function closeMenus() {
  $("menu").hidden = $("ctx").hidden = $("tab-menu").hidden = true;
  $("more").setAttribute("aria-expanded", false);
}
$("more").onclick = (e) => {
  e.stopPropagation();
  const show = $("menu").hidden;
  closeMenus();
  closePopover();
  $("menu").hidden = !show;
  $("more").setAttribute("aria-expanded", show);
  if (show) syncFormat();
};
$("menu").onmousedown = (e) => e.preventDefault();
$("menu").onclick = (e) => {
  const cmd = e.target.closest("[data-cmd]");
  if (!cmd || cmd.disabled) return;
  closeMenus();
  commands[cmd.dataset.cmd]();
};
$("search").onclick = () => commands.find();
$("outline-btn").onclick = () => toggleOutline();
$("home-btn").onclick = () => commands.home();
document.querySelector(".tab-new").onclick = () => commands.newTab();
$("history-close").onclick = closeHistory;
$("version-restore").onclick = restoreVersion;
$("version-done").onclick = () => {
  closeHistory();
  exitPreview();
};

async function pasteFromClipboard() {
  const text = await (isTauri ? readText() : navigator.clipboard.readText()).catch(() => "");
  if (text) editor?.paste(text);
}

const ctxActions = {
  cut: () => document.execCommand("cut"),
  copy: () => document.execCommand("copy"),
  paste: pasteFromClipboard,
  selectAll: () => editor?.selectAll(),
  bold: () => editor?.format("bold"),
  italic: () => editor?.format("italic"),
  underline: () => editor?.format("underline"),
  read: () => reader.start(),
  link: () => openLink(),
  unlink: () => {
    const l = editor?.linkRange();
    if (!l) return;
    getSelection().removeAllRanges();
    getSelection().addRange(l.range);
    editor.format("link", null);
  },
};

// The word under the pointer, as a range
function wordAt(x, y) {
  const at = document.caretRangeFromPoint(x, y);
  const node = at?.startContainer;
  if (node?.nodeType !== 3) return null;
  const isWord = (ch) => /[\p{L}\p{M}'’-]/u.test(ch ?? "");
  let start = at.startOffset;
  let end = at.startOffset;
  while (start > 0 && isWord(node.data[start - 1])) start--;
  while (end < node.length && isWord(node.data[end])) end++;
  const word = node.data.slice(start, end).replace(/^['’-]+|['’-]+$/g, "");
  if (word.length < 2) return null;
  const range = document.createRange();
  range.setStart(node, start);
  range.setEnd(node, end);
  return { word, range };
}

async function spellingItems(e) {
  const box = $("ctx-spell");
  box.replaceChildren();
  if (!isTauri || !editor || !getSelection().isCollapsed) return;
  const hit = wordAt(e.clientX, e.clientY);
  if (!hit || !hit.range.startContainer.parentElement?.isContentEditable) return;
  const suggestions = await invoke("spell_suggest", { word: hit.word }).catch(() => null);
  if (!suggestions) return;
  const item = (label, action, cls = "") => {
    const b = document.createElement("button");
    b.textContent = label;
    b.className = cls;
    b.onclick = () => {
      closeMenus();
      action();
    };
    return b;
  };
  const replace = (text) => {
    getSelection().removeAllRanges();
    getSelection().addRange(hit.range);
    editor.paste(text);
  };
  box.append(
    ...(suggestions.length ? suggestions.map((s) => item(s, () => replace(s), "suggestion")) : [item("No suggestions", () => {}, "muted")]),
    item("Add to Dictionary", () => invoke("spell_add", { word: hit.word })),
    document.createElement("hr"),
  );
}

// Our own right-click menu instead of the browser's; outside the document there is none, like a native app
document.addEventListener("contextmenu", async (e) => {
  e.preventDefault();
  closeMenus();
  closePopover();
  if (!doc || !e.target.closest?.(".docx")) return;
  const menu = $("ctx");
  $("selbar").hidden = true;
  imageTools.deselect();
  // Right-clicking outside the selection moves the caret there first, like native editors
  const at = document.caretRangeFromPoint(e.clientX, e.clientY);
  const sel = getSelection();
  const insideSelection = sel.rangeCount && !sel.isCollapsed && at && sel.getRangeAt(0).isPointInRange(at.startContainer, at.startOffset);
  if (at && pane.contains(at.startContainer) && !insideSelection) {
    sel.removeAllRanges();
    sel.addRange(at);
  }
  const hasSel = !sel.isCollapsed;
  for (const el of menu.querySelectorAll("[data-edit]")) el.hidden = !editor;
  for (const k of ["cut", "copy", "bold", "italic", "underline"]) menu.querySelector(`[data-ctx="${k}"]`).disabled = !hasSel;
  menu.querySelector('[data-ctx="read"]').firstChild.textContent = hasSel ? "Read Selection Aloud" : "Read Aloud From Here";
  const link = editor?.linkRange();
  const linkItem = menu.querySelector('[data-ctx="link"]');
  linkItem.firstChild.textContent = link ? "Edit Link…" : "Add Link…";
  linkItem.disabled = !link && getSelection().isCollapsed;
  menu.querySelector('[data-ctx="unlink"]').hidden = !link;
  const table = editor?.tableInfo();
  $("ctx-table").hidden = !table;
  for (const b of $("ctx-table").querySelectorAll("button")) {
    b.disabled = !table?.editable;
    b.title = table && !table.editable ? "Tables with merged cells can't be restructured yet" : "";
  }
  await spellingItems(e);
  menu.hidden = false;
  // offset sizes, not getBoundingClientRect: the open animation scales the menu down
  menu.style.left = `${Math.max(6, Math.min(e.clientX, window.innerWidth - menu.offsetWidth - 6))}px`;
  menu.style.top = `${Math.max(6, Math.min(e.clientY, window.innerHeight - menu.offsetHeight - 6))}px`;
});
$("ctx").onmousedown = (e) => e.preventDefault();
$("ctx").onclick = (e) => {
  const tableBtn = e.target.closest("[data-table]");
  if (tableBtn && !tableBtn.disabled) {
    closeMenus();
    editor?.tableOp(tableBtn.dataset.table);
    return;
  }
  const action = e.target.closest("[data-ctx]");
  if (!action || action.disabled) return;
  closeMenus();
  ctxActions[action.dataset.ctx]();
};
document.addEventListener("mousedown", (e) => {
  if (!e.target.closest(".menu, #more")) closeMenus();
  if (!e.target.closest("#format-pop, #aa")) closePopover();
  if (!e.target.closest("#goal-pop, #words")) closeGoal();
  if (!e.target.closest("#link-pop")) closeLink();
  $("link-tip").hidden = true;
  if (!e.target.closest("#slash")) slash.hide();
  if (e.target.closest("#view")) {
    mouseDown = true;
    $("selbar").hidden = true;
  }
});
document.addEventListener("mouseup", () => {
  if (!mouseDown) return;
  mouseDown = false;
  setTimeout(placeSelbar, 10);
});
window.addEventListener("blur", () => {
  closeMenus();
  closePopover();
});

// ---------- Window ----------

$("save").onclick = () => save();
$("file").onchange = (e) => e.target.files[0] && openDocument(e.target.files[0].name, e.target.files[0]);
$("image-file").onchange = async (e) => {
  const f = e.target.files[0];
  if (f) insertImageData(await f.arrayBuffer(), f.name);
  e.target.value = "";
};

if (isTauri) {
  const win = getCurrentWindow();
  $("win-close").onclick = () => win.close();
  $("win-min").onclick = () => win.minimize();
  $("win-max").onclick = () => win.toggleMaximize();
}
window.onfocus = () => document.body.classList.remove("blurred");
window.onblur = () => document.body.classList.add("blurred");
window.addEventListener("resize", () => {
  closePopover();
  placeSelbar();
});
let scrollTick = false;
let placeTimer;
view.addEventListener(
  "scroll",
  () => {
    if (!$("selbar").hidden) placeSelbar();
    slash.hide();
    caret.update();
    clearTimeout(placeTimer);
    placeTimer = setTimeout(rememberPlace, 400);
    if (scrollTick) return;
    scrollTick = true;
    requestAnimationFrame(() => {
      scrollTick = false;
      trackOutline();
    });
  },
  { passive: true },
);

// ---------- Document events ----------

document.addEventListener("selectionchange", () => {
  clearTimeout(selTimer);
  selTimer = setTimeout(() => {
    syncFormat();
    placeSelbar();
    markCurrentParagraph();
  }, 60);
});

view.addEventListener("beforeinput", (e) => editor?.onBeforeInput(e));
view.addEventListener("paste", (e) => {
  e.preventDefault();
  const image = [...e.clipboardData.files].find((f) => f.type.startsWith("image/"));
  if (image) return image.arrayBuffer().then((b) => insertImageData(b, image.name || `pasted.${image.type.split("/")[1]}`));
  editor?.paste(e.clipboardData.getData("text/plain"));
});
view.addEventListener("cut", (e) => {
  if (editor) editor.cut(e);
  else e.preventDefault();
});
let idleTimer;
view.addEventListener("input", (e) => {
  if (!editor) return;
  editor.autoformat(e);
  editor.normalize();
  slash.update();
  version++;
  $("selbar").hidden = true;
  if (!dirty) {
    dirty = true;
    updateUi();
  }
  syncFormat();
  scheduleSave();
  clearTimeout(idleTimer);
  idleTimer = setTimeout(() => {
    find.refresh();
    countWords();
    buildOutline();
    updateUi();
  }, 400);
});

view.addEventListener("dragstart", (e) => e.preventDefault());
// Browser fallback; in the app the webview drag-drop event below handles it
view.ondragover = (e) => {
  e.preventDefault();
  view.classList.add("drag");
};
view.ondragleave = () => view.classList.remove("drag");
view.ondrop = async (e) => {
  e.preventDefault();
  view.classList.remove("drag");
  const f = e.dataTransfer.files[0];
  if (!f) return;
  if (f.type.startsWith("image/")) insertImageData(await f.arrayBuffer(), f.name);
  else openDocument(f.name, f);
};

// Browser shortcuts that make no sense in a desktop app
const BLOCKED = new Set(["F5", "F7", "F12", "BrowserBack", "BrowserForward", "BrowserRefresh"]);
const BLOCKED_CTRL = new Set(["r", "g", "h", "j", "d", "u", "l", "t", "w", "shift+r", "shift+i", "shift+j", "shift+c", "shift+delete"]);

document.addEventListener("keydown", (e) => {
  if (palette.isOpen()) return;
  if (slash.key(e)) return;
  if (e.key === "Escape") {
    if (document.body.classList.contains("focus-mode")) toggleFocus(false);
    if (reader.isReading()) reader.stop();
    if (!$("history").hidden) closeHistory();
    closeMenus();
    closePopover();
    closeGoal();
    find.close();
    $("selbar").hidden = true;
  }
  if (e.ctrlKey && !e.shiftKey && !e.altKey && e.key.toLowerCase() === "k") {
    e.preventDefault();
    palette.open();
    return;
  }
  if (e.altKey && e.key === "Home") {
    e.preventDefault();
    commands.home();
    return;
  }
  // Keys typed into a field (find, link, goal…) never reach the document, even though its selection stays put
  const typingInField = !!e.target.closest?.("input, select, textarea");
  if (editor && !typingInField && editor.inDoc() && editor.onKey(e)) {
    syncFormat();
    return;
  }
  const k = e.key.toLowerCase();
  if (BLOCKED.has(e.key) || (e.altKey && (e.key === "ArrowLeft" || e.key === "ArrowRight"))) return e.preventDefault();
  if (!e.ctrlKey) return;
  const inField = typingInField;
  const actions = {
    o: e.shiftKey ? commands.outline : pickFile,
    n: !e.shiftKey && newDocument,
    s: () => save(e.shiftKey),
    f: e.shiftKey ? commands.focus : commands.find,
    h: commands.replace,
    w: e.shiftKey ? commands.closeAllTabs : commands.closeTab,
    t: commands.newTab,
    tab: e.shiftKey ? commands.prevTab : commands.nextTab,
    k: e.shiftKey && commands.link,
    ...(/^[1-9]$/.test(k) && { [k]: () => tabs[Math.min(+k, tabs.length) - 1] && switchTab(tabs[Math.min(+k, tabs.length) - 1]) }),
    p: commands.print,
    a: !inField && editor && (() => editor.selectAll()),
    "=": commands.zoomIn,
    "+": commands.zoomIn,
    "-": commands.zoomOut,
    0: commands.zoomReset,
  };
  if (actions[k]) {
    e.preventDefault();
    actions[k]();
    return;
  }
  if (BLOCKED_CTRL.has((e.shiftKey ? "shift+" : "") + k)) e.preventDefault();
});
view.addEventListener(
  "wheel",
  (e) => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    setZoom(zoom - Math.sign(e.deltaY) * 0.1);
  },
  { passive: false },
);

applyTheme();
switchTab(newTab());
showStart();
if (store.get("outline", false)) document.body.classList.add("outline-open");
buildOutline();
updateUi();

if (isTauri) {
  // The window starts hidden so it never flashes white
  getCurrentWindow().show();
  getCurrentWebview().onDragDropEvent(async (e) => {
    view.classList.toggle("drag", e.payload.type === "over" || e.payload.type === "enter");
    const path = e.payload.type === "drop" && e.payload.paths[0];
    if (!path) return;
    if (IMAGE_EXT.test(path)) insertImageData(await invoke("read_file", { path }), path);
    else loadPath(path);
  });
  getCurrentWindow().onCloseRequested(async (e) => {
    rememberPlace();
    saveSession();
    // Every tab: auto-save what can be saved, ask about the rest
    for (const tab of [...tabs]) {
      if (tab !== active && !tab.dirty) continue;
      switchTab(tab);
      if (autoSaving() && dirty) await save();
      if (dirty && !(await confirmUser(`Close without saving changes to “${doc.name}”?`))) return e.preventDefault();
    }
    if (doc?.unsnapped && !dirty) await snapshot(true);
    clearBackup();
  });

  // Offer to recover work left behind by a crash
  const meta = await invoke("backup_meta").catch(() => null);
  let started = false;
  if (meta) {
    const { name, path } = JSON.parse(meta);
    if (await confirmUser(`Recover unsaved changes to “${name}”?`)) {
      await load(name, await invoke("backup_data"), path);
      dirty = started = true;
      updateUi();
    } else clearBackup();
  }
  // File passed on the command line ("Open with" / double-click), else pick up where we left off
  const path = await invoke("opened_path");
  if (!started && setting("reopen")) {
    // Bring back the tabs that were open last time, then focus the one that was in front
    const session = store.get("session", null);
    const paths = session?.paths?.length ? session.paths : [store.get("last", null)].filter(Boolean);
    for (const p of paths) await loadPath(p, { quiet: true });
    const front = session?.active && tabs.find((t) => (t === active ? doc : t.doc)?.path === session.active);
    if (front) switchTab(front);
  }
  if (path) loadPath(path);
}
