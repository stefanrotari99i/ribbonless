# Security

Ribbonless opens files that may come from anyone, so security problems matter.

## Reporting

Please report vulnerabilities privately through
[GitHub security advisories](https://github.com/stefanrotari99i/ribbonless/security/advisories/new), not in a public issue.
Include steps to reproduce and, if possible, a document that triggers the problem.

You can expect a first reply within a week. Fixes go out in a new release, and you'll be credited unless you prefer not to be.

## What's in scope

- A document that can run script, load remote content, or read or write files outside what you opened or saved
- A document that makes Ribbonless lose or damage data in other parts of the file
- Anything that gets around the content security policy or the file type rules in `src-tauri/src/main.rs`

## Design

- No network access: the content security policy blocks remote loads, and the app makes no requests of its own.
- Documents can't run code; scripts are limited to the app's own files.
- The page can only read documents and images and only write `.docx`, `.md` and `.html`.
- Links open in the default browser only on Ctrl+click and only for `http`, `https`, `mailto` and `tel`.
