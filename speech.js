// Read aloud with the system voices, highlighting the part being read
export function createReader(view, { onChange, root = () => view }) {
  const highlight = new Highlight();
  CSS.highlights.set("reading", highlight);
  let queue = [];
  let active = false;

  function stop() {
    active = false;
    queue = [];
    speechSynthesis.cancel();
    highlight.clear();
    onChange(false);
  }

  function next() {
    if (!active) return;
    const range = queue.shift();
    if (!range) return stop();
    const text = range
      .toString()
      .replace(/\u2003/g, " ")
      .trim();
    if (!text) return next();
    highlight.clear();
    highlight.add(range);
    const box = range.getBoundingClientRect();
    const area = view.getBoundingClientRect();
    if (box.top < area.top + 120 || box.bottom > area.bottom - 60) {
      view.scrollBy({ top: box.top - area.top - area.height / 3, behavior: "smooth" });
    }
    const u = new SpeechSynthesisUtterance(text);
    u.rate = 1.02;
    u.onend = next;
    u.onerror = (e) => {
      if (e.error !== "interrupted" && e.error !== "canceled") stop();
    };
    speechSynthesis.speak(u);
  }

  // Reads the selection, or from the caret's paragraph to the end of the document
  function start() {
    stop();
    const sel = getSelection();
    const paragraphs = [...root().querySelectorAll(".docx article p")];
    if (sel.rangeCount && !sel.isCollapsed && root().contains(sel.anchorNode)) {
      queue = [sel.getRangeAt(0).cloneRange()];
    } else {
      const node = sel.anchorNode;
      const here = node && root().contains(node) ? (node.nodeType === 1 ? node : node.parentElement).closest("p") : null;
      const from = Math.max(0, paragraphs.indexOf(here));
      queue = paragraphs.slice(from).map((p) => {
        const r = document.createRange();
        r.selectNodeContents(p);
        return r;
      });
    }
    if (!queue.length) return;
    active = true;
    onChange(true);
    next();
  }

  return { start, stop, toggle: () => (active ? stop() : start()), isReading: () => active };
}
