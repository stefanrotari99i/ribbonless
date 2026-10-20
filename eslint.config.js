import js from "@eslint/js";
import globals from "globals";

export default [
  { ignores: ["dist/", "src-tauri/", "node_modules/", "test-results/", "playwright-report/"] },
  js.configs.recommended,
  {
    languageOptions: { ecmaVersion: 2024, sourceType: "module", globals: { ...globals.browser } },
    rules: {
      "no-unused-vars": ["error", { args: "none", caughtErrors: "none" }],
      "no-empty": ["error", { allowEmptyCatch: true }],
      eqeqeq: ["error", "smart"],
      "prefer-const": "error",
    },
  },
  { files: ["tests/**", "*.config.js"], languageOptions: { globals: { ...globals.node } } },
];
