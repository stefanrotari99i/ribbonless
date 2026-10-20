// Exports the document as shown on screen to Markdown or a self-contained HTML page

const HEADING = /\bdocx_(title|heading(\d))\b/;
const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const mdEsc = (s) => s.replace(/([\\`*_[\]#|<>])/g, "\\$1");

function styleOf(node, p) {
  const s = getComputedStyle(node.parentElement);
  let u = false;
  for (let a = node.parentElement; a && a !== p.parentElement; a = a.parentElement) {
    if (getComputedStyle(a).textDecorationLine.includes("underline")) u = true;
  }
  return { b: parseInt(s.fontWeight) >= 600, i: s.fontStyle === "italic", u };
}

// How plain text looks in this paragraph (heading styles are bold already)
function plainStyle(p) {
  const probe = document.createElement("span");
  const text = probe.appendChild(document.createTextNode("x"));
  p.append(probe);
  const f = styleOf(text, p);
  probe.remove();
  return f;
}

// Text runs with formatting that differs from the paragraph's own; images come back as their element
function runs(p) {
  const base = plainStyle(p);
  const out = [];
  const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  for (let n = walk.nextNode(); n; n = walk.nextNode()) {
    if (n.nodeName === "IMG") {
      out.push({ img: n });
      continue;
    }
    if (n.nodeName === "BR") {
      out.push({ br: true });
      continue;
    }
    if (n.nodeType !== 3 || !n.data) continue;
    const s = styleOf(n, p);
    const f = { b: s.b && !base.b, i: s.i && !base.i, u: s.u && !base.u };
    const text = n.data.replace(/\u2003/g, "\t").replace(/\u200B/g, "");
    if (!text) continue;
    const last = out.at(-1);
    if (last && last.text !== undefined && last.b === f.b && last.i === f.i && last.u === f.u) last.text += text;
    else out.push({ text, ...f });
  }
  return out;
}

function imageData(img) {
  try {
    const canvas = document.createElement("canvas");
    canvas.width = img.naturalWidth;
    canvas.height = img.naturalHeight;
    canvas.getContext("2d").drawImage(img, 0, 0);
    return canvas.toDataURL(/\.jpe?g/i.test(img.src) ? "image/jpeg" : "image/png", 0.9);
  } catch {
    return null;
  }
}

function kindOf(p, headingAll) {
  const h = p.className.match(HEADING);
  if (h) return { tag: "h", level: h[1] === "title" ? 1 : Math.min(+h[2] + (headingAll ? 1 : 0), 6) };
  if ([...p.classList].some((c) => c.startsWith("docx-num-"))) {
    // Numbered lists draw their marker with a CSS counter; bullets use a fixed character
    const marker = getComputedStyle(p, "::before").content;
    return { tag: "li", ordered: marker.includes("counter(") };
  }
  return { tag: "p" };
}

// Top-level blocks in reading order: paragraphs and tables
function blocks(view) {
  return [...view.querySelectorAll(".docx article")].flatMap((a) =>
    [...a.children].filter((c) => c.tagName === "P" || c.tagName === "TABLE"),
  );
}

export function toMarkdown(view) {
  const hasTitle = !!view.querySelector(".docx article p.docx_title");
  const out = [];
  let inList = false;
  const inline = (p) =>
    runs(p)
      .map((r) => {
        if (r.br) return "  \n";
        if (r.img) return "";
        const t = mdEsc(r.text);
        if (!t.trim()) return t;
        const [, pre, core, post] = t.match(/^(\s*)([\s\S]*?)(\s*)$/);
        let s = core;
        if (r.b) s = `**${s}**`;
        if (r.i) s = `*${s}*`;
        if (r.u) s = `<u>${s}</u>`;
        return pre + s + post;
      })
      .join("");
  for (const b of blocks(view)) {
    if (b.tagName === "TABLE") {
      const rows = [...b.rows].map((tr) =>
        [...tr.cells].map((td) => [...td.querySelectorAll("p")].map(inline).join(" ").replace(/\n/g, " ").trim() || " "),
      );
      if (!rows.length) continue;
      const cols = Math.max(...rows.map((r) => r.length));
      const line = (r) => `| ${Array.from({ length: cols }, (_, i) => r[i] ?? " ").join(" | ")} |`;
      out.push("", line(rows[0]), `| ${Array(cols).fill("---").join(" | ")} |`, ...rows.slice(1).map(line), "");
      inList = false;
      continue;
    }
    const text = inline(b).trim();
    const k = kindOf(b, hasTitle);
    if (!text) {
      if (!inList) out.push("");
      continue;
    }
    if (k.tag === "li") {
      if (!inList) out.push("");
      out.push(`${k.ordered ? "1." : "-"} ${text}`);
      inList = true;
      continue;
    }
    if (inList) out.push("");
    inList = false;
    out.push(k.tag === "h" ? `${"#".repeat(k.level)} ${text}` : text, "");
  }
  return (
    out
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim() + "\n"
  );
}

export function toHtml(view, title) {
  const hasTitle = !!view.querySelector(".docx article p.docx_title");
  const inline = (p) =>
    runs(p)
      .map((r) => {
        if (r.br) return "<br>";
        if (r.img) {
          const src = imageData(r.img);
          return src ? `<img src="${src}" alt="">` : "";
        }
        let s = esc(r.text).replace(/\t/g, "&emsp;");
        if (r.b) s = `<strong>${s}</strong>`;
        if (r.i) s = `<em>${s}</em>`;
        if (r.u) s = `<u>${s}</u>`;
        return s;
      })
      .join("");
  const body = [];
  let list = null;
  const closeList = () => {
    if (list) body.push(`</${list}>`);
    list = null;
  };
  for (const b of blocks(view)) {
    if (b.tagName === "TABLE") {
      closeList();
      const rows = [...b.rows].map(
        (tr) => `<tr>${[...tr.cells].map((td) => `<td>${[...td.querySelectorAll("p")].map(inline).join("<br>")}</td>`).join("")}</tr>`,
      );
      body.push(`<table>${rows.join("")}</table>`);
      continue;
    }
    const k = kindOf(b, hasTitle);
    const content = inline(b);
    const align = getComputedStyle(b).textAlign;
    const style = ["center", "right", "justify"].includes(align) ? ` style="text-align:${align}"` : "";
    if (k.tag === "li") {
      const want = k.ordered ? "ol" : "ul";
      if (list !== want) {
        closeList();
        body.push(`<${want}>`);
        list = want;
      }
      body.push(`<li${style}>${content}</li>`);
      continue;
    }
    closeList();
    if (!content.trim()) continue;
    body.push(k.tag === "h" ? `<h${k.level}${style}>${content}</h${k.level}>` : `<p${style}>${content}</p>`);
  }
  closeList();
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(title)}</title>
<style>
  :root { color-scheme: light dark; }
  body { margin: 0 auto; max-width: 46rem; padding: 3rem 1.5rem 5rem; font: 17px/1.6 -apple-system, "Segoe UI", system-ui, sans-serif; color: #1d1d1f; background: #fff; }
  @media (prefers-color-scheme: dark) { body { color: #f5f5f7; background: #1c1c1e; } td { border-color: #3a3a3c !important; } }
  h1, h2, h3, h4 { line-height: 1.25; letter-spacing: -0.01em; margin: 2em 0 0.5em; }
  h1 { font-size: 2.1em; margin-top: 0; }
  img { max-width: 100%; height: auto; border-radius: 6px; }
  table { border-collapse: collapse; width: 100%; margin: 1.2em 0; font-size: 0.92em; }
  td { border: 1px solid #d2d2d7; padding: 0.45em 0.7em; vertical-align: top; }
  tr:first-child td { font-weight: 600; }
</style>
</head>
<body>
${body.join("\n")}
</body>
</html>
`;
}
