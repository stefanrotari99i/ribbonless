// Runs before the page draws, so the window never flashes the wrong theme
if (window.__RIBBONLESS_MICA__) document.documentElement.classList.add("mica");
try {
  document.documentElement.dataset.theme =
    JSON.parse(localStorage.getItem("theme")) ?? (matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
} catch {}
