// Stands in for the desktop side (Rust commands and plugins) so desktop-only screens work in a plain browser.
// Passed to page.addInitScript, so it runs before the app; `files` maps paths to base64 document bytes.
export function installTauriMock({ files, versions, openPath, imagePath }) {
  const bytes = (b64) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)).buffer;
  let nextId = 1;
  const handlers = {
    opened_path: () => null,
    backup_meta: () => null,
    backup_write: () => null,
    backup_clear: () => null,
    // Raw-body commands get their arguments as headers; tests read window.__writes
    write_file: (body, { headers }) => {
      (window.__writes ??= []).push({ path: decodeURIComponent(headers.path), size: body.byteLength });
    },
    version_save: () => null,
    read_file: ({ path }) => {
      if (!files[path]) throw new Error(`no such file: ${path}`);
      return bytes(files[path]);
    },
    version_list: () => versions.map((v) => v.ts),
    version_read: ({ ts }) => bytes(files[versions.find((v) => v.ts === ts).file]),
    fonts: () => ["Aptos", "Arial", "Calibri", "Cambria", "Consolas", "Georgia", "Segoe UI", "Times New Roman", "Verdana"],
    spell_suggest: ({ word }) => ({ ribon: ["Ribbon", "Robin", "Rib on"] })[word.toLowerCase()] ?? null,
    spell_add: () => null,
    "plugin:dialog|open": ({ options }) => (options?.filters?.[0]?.name === "Image" ? imagePath : openPath),
    "plugin:dialog|save": ({ options }) => options?.defaultPath ?? openPath,
    "plugin:dialog|ask": () => true,
    "plugin:dialog|message": () => null,
    "plugin:clipboard-manager|read_text": () => "",
    "plugin:event|listen": () => nextId++,
    "plugin:event|unlisten": () => null,
  };
  window.__TAURI_INTERNALS__ = {
    metadata: { currentWindow: { label: "main" }, currentWebview: { windowLabel: "main", label: "main" } },
    transformCallback: (cb) => {
      const id = nextId++;
      window[`_${id}`] = cb;
      return id;
    },
    unregisterCallback: () => {},
    convertFileSrc: (path) => path,
    invoke: async (cmd, args, options) => {
      if (handlers[cmd]) return handlers[cmd](args ?? {}, options ?? {});
      if (cmd.startsWith("plugin:window|") || cmd.startsWith("plugin:webview|") || cmd.startsWith("plugin:opener|")) return null;
      throw new Error(`unmocked command ${cmd}`);
    },
  };
}
