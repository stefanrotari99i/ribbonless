# Contributing

Thanks for helping. A few things keep the project easy to work on:

## Before you start

- For anything bigger than a small fix, open an issue first so we can agree on the approach.
- Keep the app quiet. New features should not add toolbar buttons; put them in the command bar (`Ctrl+K`),
  the `⋯` menu or the `Aa` panel.
- Never break someone's document. Changes to `editor.js` that touch saving need a test in `tests/`.

## Setup

See "Building from source" in the README. Then:

```bash
npm run lint    # ESLint + Prettier check
npm test        # Playwright end-to-end tests
```

Run `npm run format` to format your changes.

## Code style

- Plain JavaScript modules, no framework, no build-time magic beyond Vite.
- Small, named functions. Comments explain _why_, in one line, and only where the code can't say it.
- No new runtime dependencies without a good reason; the app should stay small.

## Bug reports

Include what you did, what you expected and what happened. If a document misbehaves, a small `.docx` that shows the
problem helps most. Please don't attach documents with private content.
