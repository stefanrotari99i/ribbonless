// Command bar (Ctrl+K) and the slash menu share one fuzzy-filtered list

// Higher is better; null when the query doesn't match at all
function score(title, query) {
  const t = title.toLowerCase();
  const q = query.toLowerCase().trim();
  if (!q) return 0;
  const at = t.indexOf(q);
  if (at === 0) return 300 - t.length;
  if (at > 0) return (t[at - 1] === " " ? 220 : 150) - at;
  let i = 0;
  let gaps = 0;
  for (const ch of t) {
    if (ch === q[i]) i++;
    else if (i) gaps++;
    if (i === q.length) return 80 - gaps;
  }
  return null;
}

export function filterItems(items, query) {
  // The title decides; keywords only rescue items the title doesn't match
  const rank = (item) => {
    const byTitle = score(item.title, query);
    if (byTitle !== null) return byTitle;
    const byKeyword = item.keywords ? score(item.keywords, query) : null;
    return byKeyword === null ? null : byKeyword - 200;
  };
  const ranked = items
    .map((item) => ({ item, s: rank(item) }))
    .filter((x) => x.s !== null)
    .sort((a, b) => b.s - a.s);
  // Keep each group together, groups ordered by their best match
  const groupAt = new Map();
  for (const { item } of ranked) if (!groupAt.has(item.group)) groupAt.set(item.group, groupAt.size);
  return ranked.sort((a, b) => groupAt.get(a.item.group) - groupAt.get(b.item.group)).map((x) => x.item);
}

// A keyboard-driven list inside `listEl`; returns helpers to fill and navigate it
function listController(listEl, onRun) {
  let shown = [];
  let at = 0;
  function render(items) {
    shown = items;
    at = 0;
    let lastGroup = null;
    listEl.replaceChildren(
      ...items.flatMap((item, i) => {
        const out = [];
        if (item.group && item.group !== lastGroup) {
          const h = document.createElement("div");
          h.className = "pl-group";
          h.textContent = item.group;
          out.push(h);
          lastGroup = item.group;
        }
        const b = document.createElement("button");
        b.className = "pl-item";
        b.dataset.i = i;
        const title = document.createElement("span");
        title.className = "pl-title";
        title.textContent = item.title;
        b.append(title);
        if (item.hint) {
          const hint = document.createElement("kbd");
          hint.textContent = item.hint;
          b.append(hint);
        }
        b.onmousemove = () => select(i);
        b.onmousedown = (e) => e.preventDefault();
        b.onclick = () => run(i);
        out.push(b);
        return out;
      }),
    );
    if (!items.length) {
      const empty = document.createElement("div");
      empty.className = "pl-empty";
      empty.textContent = "No matches";
      listEl.append(empty);
    }
    select(0);
  }
  function select(i) {
    if (!shown.length) return;
    at = (i + shown.length) % shown.length;
    for (const b of listEl.querySelectorAll(".pl-item")) b.classList.toggle("active", +b.dataset.i === at);
    listEl.querySelector(`.pl-item[data-i="${at}"]`)?.scrollIntoView({ block: "nearest" });
  }
  function run(i = at) {
    const item = shown[i];
    if (item) onRun(item);
  }
  function key(e) {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      select(at + 1);
      return true;
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      select(at - 1);
      return true;
    }
    if (e.key === "Enter" || (e.key === "Tab" && shown.length)) {
      e.preventDefault();
      run();
      return true;
    }
    return false;
  }
  return { render, key, count: () => shown.length };
}

export function createPalette({ getItems }) {
  const root = document.getElementById("palette");
  const input = document.getElementById("palette-input");
  let saved = null; // selection to restore before running a command
  const list = listController(document.getElementById("palette-list"), (item) => {
    close();
    if (saved) {
      getSelection().removeAllRanges();
      getSelection().addRange(saved);
    }
    item.run();
  });
  const refresh = () => list.render(filterItems(getItems(), input.value).slice(0, 60));

  function open() {
    const sel = getSelection();
    saved = sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
    root.hidden = false;
    input.value = "";
    refresh();
    input.focus();
  }
  function close() {
    if (root.hidden) return;
    root.hidden = true;
    if (saved) {
      getSelection().removeAllRanges();
      getSelection().addRange(saved);
    }
  }
  input.oninput = refresh;
  input.onkeydown = (e) => {
    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      close();
      return;
    }
    list.key(e);
  };
  root.onmousedown = (e) => {
    if (e.target === root) {
      e.preventDefault();
      close();
    }
  };
  return { open, close, isOpen: () => !root.hidden };
}

// The "/" menu: shown while the caret paragraph reads "/query"
export function createSlashMenu({ getItems, editableP }) {
  const root = document.getElementById("slash");
  let paragraph = null;
  const list = listController(document.getElementById("slash-list"), (item) => {
    const p = paragraph;
    hide();
    if (!p) return;
    // Clear the "/query" text, then apply the choice to the now-empty paragraph
    const r = document.createRange();
    r.selectNodeContents(p);
    getSelection().removeAllRanges();
    getSelection().addRange(r);
    document.execCommand("delete");
    item.run();
  });

  function hide() {
    root.hidden = true;
    paragraph = null;
  }

  // Call after every input; shows, filters or hides the menu
  function update() {
    const sel = getSelection();
    const p = sel.isCollapsed && editableP(sel.anchorNode);
    const m = p && p.textContent.match(/^\/([\p{L}\d ]{0,24})$/u);
    if (!m) return hide();
    const items = filterItems(getItems(), m[1]);
    if (!items.length) return hide();
    paragraph = p;
    list.render(items);
    root.hidden = false;
    const rect = sel.getRangeAt(0).getClientRects()[0] ?? p.getBoundingClientRect();
    const below = rect.bottom + 6;
    const h = root.offsetHeight;
    root.style.top = `${below + h > window.innerHeight - 8 ? rect.top - h - 6 : below}px`;
    root.style.left = `${Math.min(window.innerWidth - root.offsetWidth - 8, rect.left)}px`;
  }

  // Returns true when the menu consumed the key
  function key(e) {
    if (root.hidden) return false;
    if (e.key === "Escape") {
      e.preventDefault();
      hide();
      return true;
    }
    return list.key(e);
  }

  root.onmousedown = (e) => e.preventDefault();
  return { update, key, hide, isOpen: () => !root.hidden };
}
