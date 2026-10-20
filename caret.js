// A caret that glides between positions instead of jumping; the native one is hidden while it runs
export function createCaret(isActive) {
  const el = document.createElement("div");
  el.className = "caret-line";
  el.hidden = true;
  document.body.append(el);
  let frame = 0;
  let last = "";

  function place() {
    frame = 0;
    const sel = getSelection();
    const node = sel.anchorNode;
    const host = node && (node.nodeType === 1 ? node : node.parentElement)?.closest('.docx article[contenteditable="true"]');
    if (!isActive() || !host || !sel.isCollapsed || !document.hasFocus() || !host.contains(document.activeElement ?? host)) {
      el.hidden = true;
      return;
    }
    const range = sel.getRangeAt(0);
    let rect = [...range.getClientRects()].at(-1);
    if (!rect || !rect.height) {
      // Empty paragraphs have no text box; use the paragraph's first line instead
      const p = (node.nodeType === 1 ? node : node.parentElement).closest("p") ?? node.parentElement;
      const box = p.getBoundingClientRect();
      const size = parseFloat(getComputedStyle(p.querySelector("span") ?? p).fontSize) || 15;
      const left = box.left + (parseFloat(getComputedStyle(p).paddingLeft) || 0) + (parseFloat(getComputedStyle(p).textIndent) || 0);
      rect = { left, top: box.top, height: size * 1.25 };
    }
    const key = `${Math.round(rect.left)},${Math.round(rect.top)},${Math.round(rect.height)}`;
    el.hidden = false;
    el.style.height = `${rect.height}px`;
    el.style.transform = `translate(${rect.left}px, ${rect.top}px)`;
    if (key !== last) {
      // Restart the blink so the caret stays solid while moving
      el.classList.remove("blink");
      void el.offsetWidth;
      el.classList.add("blink");
      last = key;
    }
  }

  const update = () => {
    if (!frame) frame = requestAnimationFrame(place);
  };
  for (const evt of ["selectionchange", "input", "keydown", "mouseup", "focusin", "focusout"]) document.addEventListener(evt, update, true);
  window.addEventListener("resize", update);
  document.addEventListener("scroll", update, true); // the caret is fixed-position, so it has to follow the page
  window.addEventListener("blur", () => (el.hidden = true));
  window.addEventListener("focus", update);
  return { update };
}
