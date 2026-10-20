<p align="center">
  <img src="assets/icon.png" width="112" alt="Ribbonless icon">
</p>

<h1 align="center">Ribbonless</h1>

<p align="center">
  A small, fast editor for Word documents on Windows. Free and open source.
</p>

<p align="center">
  <a href="https://github.com/stefanrotari99i/ribbonless/releases/latest"><img src="https://img.shields.io/github/v/release/stefanrotari99i/ribbonless?style=flat-square&label=download" alt="Latest release"></a>
  <a href="https://github.com/stefanrotari99i/ribbonless/releases"><img src="https://img.shields.io/github/downloads/stefanrotari99i/ribbonless/total?style=flat-square" alt="Downloads"></a>
  <img src="https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-0078d4?style=flat-square" alt="Platform">
  <a href="LICENSE"><img src="https://img.shields.io/github/license/stefanrotari99i/ribbonless?style=flat-square" alt="License"></a>
  <a href="https://github.com/stefanrotari99i/ribbonless/actions/workflows/ci.yml"><img src="https://img.shields.io/github/actions/workflow/status/stefanrotari99i/ribbonless/ci.yml?branch=main&style=flat-square" alt="CI"></a>
</p>

![Ribbonless](docs/screens/hero.png)

## What is Ribbonless?

I got tired of paying for Office just to open the .docx files people send me. Most of the time I only need to read
something, change a few lines and send it back, and for that Word wants a subscription, a Microsoft account and a few
seconds to start. WordPad used to be the free option for this, and it's gone now.

Don't get me wrong, LibreOffice is great and I still use it. But it's a whole office suite, and I wanted something small
that opens fast and stays out of the way. So I made one.

Ribbonless opens a .docx, lets you edit it, and saves a normal .docx back that Word, Google Docs and LibreOffice open
fine. It only rewrites the paragraphs you changed, so the rest of the file stays as it was. No account, no telemetry,
nothing gets uploaded. It's not a full Word replacement, just the part most people actually use.

## Features

- Opens a typical document in about half a second, installer is about 2 MB
- Styles, bold/italic/underline, fonts, sizes, colors, highlight, alignment, line spacing
- Bulleted and numbered lists with levels, tables, images (paste, drag in, resize), links
- Markdown shortcuts while typing (`# `, `- `, `1. `, `**bold**`) and a `/` menu on empty lines
- Command bar (`Ctrl+K`) for everything, so the window stays clean
- Find and replace, spell check with the Windows spellchecker
- Version history: see how the document looked earlier and restore it
- Auto-save, plus a backup for documents you haven't saved yet
- Marks in the margin for every paragraph you changed since opening the file
- Tabs, outline sidebar, focus mode, read aloud, light and dark mode
- Export to PDF, Markdown or HTML

## Screenshots

The document in the screenshots is a fake Reddit thread I made for testing.

|                                                                                      |                                                                        |
| ------------------------------------------------------------------------------------ | ---------------------------------------------------------------------- |
| ![Start screen](docs/screens/start.png) Start screen with templates and recent files | ![Format panel](docs/screens/format-panel.png) Format panel, dark mode |
| ![Version history](docs/screens/version-history.png) Version history                 | ![Command bar](docs/screens/command-bar.png) Command bar (`Ctrl+K`)    |

<details>
<summary>More screenshots</summary>

|                                                                              |                                                                   |
| ---------------------------------------------------------------------------- | ----------------------------------------------------------------- |
| ![Selection bar](docs/screens/selection-bar.png) Selection bar               | ![Slash menu](docs/screens/slash-menu.png) `/` menu               |
| ![Find and replace](docs/screens/find-replace.png) Find and replace          | ![Tables](docs/screens/table-menu.png) Table tools on right-click |
| ![Links](docs/screens/link.png) Adding a link                                | ![Spelling](docs/screens/spelling.png) Spelling suggestions       |
| ![Changes](docs/screens/changes.png) Changed paragraphs marked in the margin | ![Images](docs/screens/image.png) Resizing an image               |
| ![Focus mode](docs/screens/focus-mode.png) Focus mode                        | ![Tabs](docs/screens/tabs.png) Tabs                               |

</details>

## Download

Get the installer from [Releases](https://github.com/stefanrotari99i/ribbonless/releases/latest). Works on Windows 10
and 11 (the frosted window background needs 11).

The installer isn't code-signed yet, so Windows SmartScreen will probably warn you the first time. Click
**More info → Run anyway**, or build it yourself from source.

## Shortcuts

- `Ctrl+K` command bar, `Alt+Home` start screen
- `Ctrl+N` new document, `Ctrl+T` new tab, `Ctrl+W` close tab, `Ctrl+Shift+W` close all tabs
- `Ctrl+Tab` / `Ctrl+1`…`9` switch tabs, right-click a tab for more
- `Ctrl+O` open, `Ctrl+S` save, `Ctrl+Shift+S` save as, `Ctrl+P` print or PDF
- `Ctrl+F` find, `Ctrl+H` replace
- `Ctrl+Alt+1`…`3` headings, `Ctrl+Shift+N` body text, `Ctrl+Shift+L` bullet list
- `Ctrl+L` / `E` / `R` / `J` align, `Ctrl+[` / `]` text size
- `Ctrl+Shift+K` add link, `Ctrl+click` open link
- `Ctrl+Shift+O` outline, `Ctrl+Shift+F` focus mode

## Limitations

- Tracked changes and comments aren't shown yet. They're kept in the file, but editing a paragraph that has tracked
  changes accepts them for that paragraph.
- Headers and footers are shown but can't be edited.
- Tables with merged cells can be edited, but you can't add or remove rows and columns in them.
- Complex layouts (text boxes, columns, floating shapes) open read-only.
- Windows only for now. It's built with [Tauri](https://tauri.app), so Mac and Linux builds should mostly need testing.

## Privacy

Ribbonless makes no network requests. No accounts, no analytics, documents never leave your computer. Documents can't
run scripts or load anything from the internet, and links only open on `Ctrl+click`. More in [SECURITY.md](SECURITY.md).

## Building from source

You need [Node.js](https://nodejs.org) 20+ and [Rust](https://rustup.rs) with the MSVC build tools (see Tauri's
[prerequisites](https://tauri.app/start/prerequisites/)).

```bash
npm install
npm run tauri dev      # run with live reload
npm run tauri build    # installer ends up in src-tauri/target/release/bundle/nsis
```

Tests and lint:

```bash
npx playwright install chromium   # once
npm test
cargo test --manifest-path src-tauri/Cargo.toml
npm run lint
```

Documents are rendered with [docx-preview](https://github.com/VolodymyrBaydalka/docxjs). When you save, `editor.js`
maps each changed paragraph back to its place in the original `document.xml` and rewrites only those, so anything
Ribbonless doesn't understand is left alone.

## Contributing

It's just me working on this right now, so help is really appreciated. The easiest way: if a .docx looks wrong in
Ribbonless, [open an issue](https://github.com/stefanrotari99i/ribbonless/issues/new) with a screenshot (and the file,
if it's not private). If you want to code, anything under Limitations is fair game. See
[CONTRIBUTING.md](CONTRIBUTING.md).

## License

[MIT](LICENSE). "Word" is a trademark of Microsoft; Ribbonless is not affiliated with Microsoft.
