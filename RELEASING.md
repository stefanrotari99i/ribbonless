# Releasing

Releases are built by GitHub Actions. Pushing a tag like `v0.1.0` runs the tests, builds the Windows installer and
creates a **draft** release with `Ribbonless_<version>_x64-setup.exe` attached. Nothing is public until you publish it.

## First time: put the code on GitHub

1. Create an empty public repository named `ribbonless` on GitHub (no README, license or .gitignore; they're here already).
2. In this folder:

   ```bash
   git init -b main
   git add .
   git status            # check: no qa-local/, no .docx files except tests/fixtures
   git commit -m "Ribbonless 0.1.0"
   git remote add origin https://github.com/stefanrotari99i/ribbonless.git
   git push -u origin main
   ```

3. Open the **Actions** tab and wait for the CI run to go green.

## Every release

1. Bump the version in `package.json`, `src-tauri/tauri.conf.json` and `src-tauri/Cargo.toml` (all three the same).
2. Commit, then tag and push:

   ```bash
   git commit -am "Release 0.1.1"
   git tag v0.1.1
   git push origin main v0.1.1
   ```

3. When the workflow finishes, open **Releases**, edit the draft, write a few lines about what changed, and publish.

## Building the installer locally

```bash
npm install
npm run tauri build
```

The installer is written to `src-tauri/target/release/bundle/nsis/`. It installs per user (no admin prompt) and
registers Ribbonless for `.docx` files.

## SmartScreen

The installer isn't code-signed, so Windows shows "Windows protected your PC" on first run (More info → Run anyway).
The warning fades as more people download the same file. To remove it entirely, sign the installer, for example with
Azure Trusted Signing or a certificate from [SignPath](https://signpath.org), which is free for open-source projects.
