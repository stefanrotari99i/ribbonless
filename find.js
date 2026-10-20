const $ = (id) => document.getElementById(id);

// Find and replace; matches are drawn with CSS highlights, so the document is only touched when replacing
export function createFind(view, { replace, canReplace, onReplaced, root = () => view } = {}) {
  const bar = $("findbar");
  const input = $("find-input");
  const withInput = $("replace-input");
  const all = new Highlight();
  const current = new Highlight();
  CSS.highlights.set("find", all);
  CSS.highlights.set("find-current", current);
  let matches = [];
  let at = -1;
  let matchCase = false;

  function show() {
    $("find-count").textContent = !input.value ? "" : matches.length ? `${at + 1} of ${matches.length}` : "No matches";
    $("find-prev").disabled = $("find-next").disabled = matches.length < 2;
    $("replace-one").disabled = $("replace-all").disabled = !matches.length || !canReplace?.();
  }

  function go(i) {
    if (!matches.length) return show();
    at = (i + matches.length) % matches.length;
    current.clear();
    current.add(matches[at]);
    const r = matches[at].getBoundingClientRect();
    const v = view.getBoundingClientRect();
    // The toolbar and find bar float over the top of the view
    if (r.top < v.top + 160 || r.bottom > v.bottom - 40) view.scrollBy({ top: r.top - v.top - v.height / 3, behavior: "smooth" });
    show();
  }

  function search(keepPlace = false) {
    all.clear();
    current.clear();
    matches = [];
    const q = matchCase ? input.value : input.value.toLowerCase();
    if (q) {
      for (const block of root().querySelectorAll(".docx p")) {
        const texts = [];
        let text = "";
        const walk = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
        for (let n = walk.nextNode(); n; n = walk.nextNode()) {
          texts.push([n, text.length]);
          text += n.data;
        }
        const hay = matchCase ? text : text.toLowerCase();
        const pos = (k, end) => {
          const [n, start] = texts.find(([n, s]) => (end ? k > s && k <= s + n.length : k >= s && k < s + n.length));
          return [n, k - start];
        };
        for (let i = hay.indexOf(q); i >= 0; i = hay.indexOf(q, i + q.length)) {
          const r = document.createRange();
          r.setStart(...pos(i, false));
          r.setEnd(...pos(i + q.length, true));
          matches.push(r);
          all.add(r);
        }
      }
    }
    if (!matches.length) {
      at = -1;
      return show();
    }
    if (keepPlace) {
      at = Math.min(Math.max(at, 0), matches.length - 1);
      current.add(matches[at]);
      return show();
    }
    go(0);
  }

  function setReplace(on) {
    const can = !!canReplace?.();
    bar.classList.toggle("with-replace", on && can);
    $("find-toggle").setAttribute("aria-expanded", on && can);
    $("find-toggle").disabled = !can;
  }

  function open(withReplace = false) {
    bar.hidden = false;
    setReplace(withReplace || bar.classList.contains("with-replace"));
    (withReplace && canReplace?.() ? withInput : input).focus();
    input.select();
    search();
  }

  function close() {
    bar.hidden = true;
    all.clear();
    current.clear();
    matches = [];
  }

  function replaceOne() {
    if (at < 0 || !matches[at]) return;
    const index = at;
    replace([matches[at]], withInput.value);
    search(true);
    if (matches.length) go(Math.min(index, matches.length - 1));
  }

  function replaceAll() {
    if (!matches.length) return;
    const n = replace(matches, withInput.value);
    search();
    onReplaced?.(n);
  }

  input.oninput = () => search();
  input.onkeydown = (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      go(at + (e.shiftKey ? -1 : 1));
    }
    if (e.key === "Escape") close();
  };
  withInput.onkeydown = (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      e.ctrlKey ? replaceAll() : replaceOne();
    }
    if (e.key === "Escape") close();
  };
  $("find-next").onclick = () => go(at + 1);
  $("find-prev").onclick = () => go(at - 1);
  $("find-done").onclick = close;
  $("replace-one").onclick = replaceOne;
  $("replace-all").onclick = replaceAll;
  $("find-toggle").onclick = () => {
    setReplace(!bar.classList.contains("with-replace"));
    (bar.classList.contains("with-replace") ? withInput : input).focus();
  };
  $("find-case").onclick = () => {
    matchCase = !matchCase;
    $("find-case").setAttribute("aria-pressed", matchCase);
    search(true);
  };

  return { open, close, refresh: () => !bar.hidden && search(true) };
}
