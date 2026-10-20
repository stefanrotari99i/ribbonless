import { defineConfig } from "@playwright/test";

// Tests run the built app in Chromium under the same content security policy as the desktop app.
// The desktop side is mocked in tests/tauri-mock.js; the real Windows parts are covered by the Rust tests.
export default defineConfig({
  testDir: "tests",
  timeout: 30_000,
  fullyParallel: true,
  reporter: [["list"]],
  use: { baseURL: "http://localhost:4173", viewport: { width: 1200, height: 900 } },
  webServer: {
    command: "npm run build && npx vite preview",
    url: "http://localhost:4173",
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },
});
