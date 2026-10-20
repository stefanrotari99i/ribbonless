import { readFileSync } from "node:fs";

// `vite preview` serves the built app under the same content security policy as the real app,
// so browser tests catch anything the policy would block
const { csp } = JSON.parse(readFileSync(new URL("./src-tauri/tauri.conf.json", import.meta.url), "utf8")).app.security;
const policy = Object.entries(csp)
  .map(([k, v]) => `${k} ${v}`)
  .join("; ");

export default {
  clearScreen: false,
  server: { port: 1420, strictPort: true },
  preview: { port: 4173, strictPort: true, headers: { "Content-Security-Policy": policy } },
};
