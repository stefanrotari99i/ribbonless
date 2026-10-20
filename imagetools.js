// Select an image by clicking it: corner handles resize it (aspect ratio kept), a small bar aligns or deletes it
export function createImageTools({ view, getEditor }) {
  const frame = document.getElementById("img-frame");
  const bar = document.getElementById("img-bar");
  let box = null; // the selected <img>
  let drag = null;

  const paragraphOf = (el) => el?.closest("article p");

  function place() {
    if (!box || !box.isConnected) return deselect();
    const r = box.getBoundingClientRect();
    const area = view.getBoundingClientRect();
    if (r.bottom < area.top + 60 || r.top > area.bottom) {
      frame.hidden = bar.hidden = true;
      return;
    }
    frame.hidden = false;
    Object.assign(frame.style, { left: `${r.left}px`, top: `${r.top}px`, width: `${r.width}px`, height: `${r.height}px` });
    bar.hidden = !!drag;
    if (!drag) {
      const top = r.top - bar.offsetHeight - 10;
      bar.style.top = `${top > area.top + 60 ? top : r.bottom + 10}px`;
      bar.style.left = `${Math.max(area.left + 8, Math.min(area.right - bar.offsetWidth - 8, r.left + r.width / 2 - bar.offsetWidth / 2))}px`;
      const align = getComputedStyle(paragraphOf(box)).textAlign;
      for (const b of bar.querySelectorAll("[data-img-align]"))
        b.setAttribute("aria-pressed", align.replace("start", "left").replace("end", "right") === b.dataset.imgAlign);
      bar.querySelector('[data-img="delete"]').disabled = !!paragraphOf(box).textContent.trim();
    }
  }

  function select(img) {
    if (!paragraphOf(img)) return;
    box = img;
    getSelection().removeAllRanges();
    place();
  }

  function deselect() {
    box = null;
    drag = null;
    frame.hidden = bar.hidden = true;
  }

  // Sizes in unzoomed CSS pixels, as stored in the document; the renderer's drawing box follows the image
  function size(w, h) {
    Object.assign(box.style, { width: `${w}px`, height: `${h}px` });
    box.dataset.w = Math.round(w);
    box.dataset.h = Math.round(h);
    const holder = box.parentElement;
    if (holder && holder.style.width) Object.assign(holder.style, { width: `${w}px`, height: `${h}px` });
  }

  view.addEventListener("mousedown", (e) => {
    const img = e.target.closest?.(".docx article img");
    if (img && getEditor() && paragraphOf(img)?.isContentEditable) {
      e.preventDefault();
      select(img);
    } else if (!e.target.closest?.("#img-frame, #img-bar")) deselect();
  });

  frame.addEventListener("pointerdown", (e) => {
    const handle = e.target.closest("[data-corner]");
    if (!handle || !box) return;
    e.preventDefault();
    handle.setPointerCapture(e.pointerId);
    const rect = box.getBoundingClientRect();
    const cssW = parseFloat(getComputedStyle(box).width) || rect.width;
    const cssH = (cssW * rect.height) / rect.width;
    const room = box.closest("article").clientWidth;
    drag = {
      x: e.clientX,
      w: cssW,
      h: cssH,
      scale: rect.width / cssW,
      dir: handle.dataset.corner.includes("l") ? -1 : 1,
      room,
      last: [cssW, cssH],
    };
  });
  frame.addEventListener("pointermove", (e) => {
    if (!drag) return;
    const w = Math.max(32, Math.min(drag.room, drag.w + ((e.clientX - drag.x) / drag.scale) * drag.dir));
    drag.last = [w, (w * drag.h) / drag.w];
    size(...drag.last);
    place();
  });
  const finish = () => {
    if (!drag) return;
    const { w, h, last } = drag;
    drag = null;
    if (Math.round(last[0]) === Math.round(w)) return place();
    // Put the old size back, then apply the new one as a single undoable change
    size(w, h);
    getEditor()?.mutate(() => size(...last));
    place();
  };
  frame.addEventListener("pointerup", finish);
  frame.addEventListener("pointercancel", finish);

  bar.addEventListener("mousedown", (e) => e.preventDefault());
  bar.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    const editor = getEditor();
    if (!b || !box || !editor) return;
    const p = paragraphOf(box);
    if (b.dataset.imgAlign) {
      const r = document.createRange();
      r.setStart(p, 0);
      getSelection().removeAllRanges();
      getSelection().addRange(r);
      editor.setAlign(b.dataset.imgAlign);
      getSelection().removeAllRanges();
      requestAnimationFrame(place);
    } else if (b.dataset.img === "delete" && !p.textContent.trim()) {
      editor.mutate(() => p.remove());
      deselect();
    }
  });

  document.addEventListener(
    "keydown",
    (e) => {
      if (!box) return;
      if (e.key === "Escape") deselect();
      else if ((e.key === "Delete" || e.key === "Backspace") && !paragraphOf(box).textContent.trim()) {
        e.preventDefault();
        e.stopImmediatePropagation();
        const p = paragraphOf(box);
        getEditor()?.mutate(() => p.remove());
        deselect();
      } else if (!e.ctrlKey) deselect();
    },
    true,
  );

  view.addEventListener("scroll", () => box && place(), { passive: true });
  window.addEventListener("resize", () => box && place());
  return { deselect, refresh: () => box && place() };
}
