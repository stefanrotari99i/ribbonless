#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::path::PathBuf;
use tauri::ipc::{InvokeBody, Request, Response};
use tauri::{AppHandle, Manager};

fn header(req: &Request<'_>, key: &str) -> Result<String, String> {
    let value = req
        .headers()
        .get(key)
        .and_then(|v| v.to_str().ok())
        .ok_or(format!("missing {key}"))?;
    percent_encoding::percent_decode_str(value)
        .decode_utf8()
        .map(|s| s.into_owned())
        .map_err(|e| e.to_string())
}

fn body<'a>(req: &'a Request<'a>) -> Result<&'a [u8], String> {
    match req.body() {
        InvokeBody::Raw(bytes) => Ok(bytes),
        _ => Err("expected raw body".into()),
    }
}

// Write to a temp file first so a crash mid-save never leaves a half-written document
fn write_atomic(path: &str, data: &[u8]) -> Result<(), String> {
    let tmp = format!("{path}.saving");
    let result = std::fs::write(&tmp, data).and_then(|_| std::fs::rename(&tmp, path));
    if result.is_err() {
        let _ = std::fs::remove_file(&tmp); // e.g. Word has the file locked
    }
    result.map_err(|e| e.to_string())
}

#[test]
fn file_types_are_limited() {
    assert!(check_type(r"C:\docs\Report.DOCX", READABLE).is_ok());
    assert!(check_type(r"C:\docs\notes.md", WRITABLE).is_ok());
    assert!(check_type(r"C:\Windows\system32\evil.dll", WRITABLE).is_err());
    assert!(check_type(r"C:\Users\me\.ssh\id_rsa", READABLE).is_err());
    assert!(check_type("no-extension", WRITABLE).is_err());
}

fn backup_dir(app: &AppHandle) -> Result<PathBuf, String> {
    let dir = app.path().app_local_data_dir().map_err(|e| e.to_string())?;
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir)
}

// Windows' own spellchecker, the same one that draws the red underlines
#[cfg(windows)]
mod spell {
    use windows::core::{HSTRING, PWSTR};
    use windows::Win32::Globalization::{
        GetUserDefaultLocaleName, ISpellChecker, ISpellCheckerFactory, SpellCheckerFactory,
    };
    use windows::Win32::System::Com::{
        CoCreateInstance, CoInitializeEx, CoTaskMemFree, CLSCTX_INPROC_SERVER, COINIT_MULTITHREADED,
    };

    fn checker() -> windows::core::Result<ISpellChecker> {
        unsafe {
            let _ = CoInitializeEx(None, COINIT_MULTITHREADED);
            let factory: ISpellCheckerFactory =
                CoCreateInstance(&SpellCheckerFactory, None, CLSCTX_INPROC_SERVER)?;
            let mut buf = [0u16; 85];
            let len = GetUserDefaultLocaleName(&mut buf).max(1) as usize - 1;
            let mut lang = String::from_utf16_lossy(&buf[..len]);
            if !factory.IsSupported(&HSTRING::from(&lang))?.as_bool() {
                lang = "en-US".into();
            }
            factory.CreateSpellChecker(&HSTRING::from(&lang))
        }
    }

    // None when the word is spelled correctly
    pub fn suggest(word: &str) -> windows::core::Result<Option<Vec<String>>> {
        unsafe {
            let checker = checker()?;
            let word = HSTRING::from(word);
            let mut error = None;
            let _ = checker.Check(&word)?.Next(&mut error);
            if error.is_none() {
                return Ok(None);
            }
            let list = checker.Suggest(&word)?;
            let mut out = Vec::new();
            loop {
                let mut item = [PWSTR::null()];
                let mut fetched = 0;
                let _ = list.Next(&mut item, Some(&mut fetched));
                if fetched == 0 || out.len() >= 6 {
                    break;
                }
                out.push(item[0].to_string().unwrap_or_default());
                CoTaskMemFree(Some(item[0].0 as _));
            }
            Ok(Some(out))
        }
    }

    pub fn add(word: &str) -> windows::core::Result<()> {
        unsafe { checker()?.Add(&HSTRING::from(word)) }
    }

    #[test]
    fn lists_installed_fonts() {
        let fonts = super::font_families().unwrap();
        assert!(
            fonts.iter().any(|f| f == "Segoe UI"),
            "Segoe UI ships with every Windows"
        );
    }

    #[test]
    fn suggests_only_for_misspelled_words() {
        assert!(suggest("documnet").unwrap().is_some_and(|s| !s.is_empty()));
        assert!(suggest("document").unwrap().is_none());
    }
}

// Installed font families, as Windows itself lists them
#[cfg(windows)]
fn font_families() -> windows::core::Result<Vec<String>> {
    use windows::core::w;
    use windows::Win32::Graphics::DirectWrite::{
        DWriteCreateFactory, IDWriteFactory, DWRITE_FACTORY_TYPE_SHARED,
    };
    unsafe {
        let factory: IDWriteFactory = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED)?;
        let mut collection = None;
        factory.GetSystemFontCollection(&mut collection, false)?;
        let collection = collection.ok_or_else(windows::core::Error::empty)?;
        let mut out = Vec::new();
        for i in 0..collection.GetFontFamilyCount() {
            let names = collection.GetFontFamily(i)?.GetFamilyNames()?;
            let mut index = 0;
            let mut exists = windows::core::BOOL(0);
            names.FindLocaleName(w!("en-us"), &mut index, &mut exists)?;
            if !exists.as_bool() {
                index = 0;
            }
            let len = names.GetStringLength(index)? as usize;
            let mut buf = vec![0u16; len + 1];
            names.GetString(index, &mut buf)?;
            out.push(String::from_utf16_lossy(&buf[..len]));
        }
        out.sort_by_key(|s| s.to_lowercase());
        out.dedup();
        Ok(out)
    }
}

#[tauri::command]
async fn fonts() -> Vec<String> {
    #[cfg(windows)]
    return font_families().unwrap_or_default();
    #[cfg(not(windows))]
    return Vec::new();
}

#[tauri::command]
async fn spell_suggest(word: String) -> Option<Vec<String>> {
    #[cfg(windows)]
    return spell::suggest(&word).ok().flatten();
    #[cfg(not(windows))]
    return None;
}

#[tauri::command]
async fn spell_add(word: String) {
    #[cfg(windows)]
    let _ = spell::add(&word);
}

#[tauri::command]
fn opened_path() -> Option<String> {
    std::env::args().nth(1)
}

// Even a compromised page can only touch documents, images and exports, never programs or system files
const READABLE: &[&str] = &["docx", "png", "jpg", "jpeg", "gif", "webp", "bmp"];
const WRITABLE: &[&str] = &["docx", "md", "html"];

fn check_type(path: &str, allowed: &[&str]) -> Result<(), String> {
    let ext = std::path::Path::new(path)
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_ascii_lowercase())
        .unwrap_or_default();
    if allowed.contains(&ext.as_str()) {
        Ok(())
    } else {
        Err(format!("not a supported file type: {path}"))
    }
}

#[tauri::command]
fn read_file(path: String) -> Result<Response, String> {
    check_type(&path, READABLE)?;
    std::fs::read(&path)
        .map(Response::new)
        .map_err(|e| e.to_string())
}

#[tauri::command]
fn write_file(request: Request) -> Result<(), String> {
    let path = header(&request, "path")?;
    check_type(&path, WRITABLE)?;
    write_atomic(&path, body(&request)?)
}

#[tauri::command]
fn backup_write(app: AppHandle, request: Request) -> Result<(), String> {
    let dir = backup_dir(&app)?;
    write_atomic(&dir.join("backup.docx").to_string_lossy(), body(&request)?)?;
    std::fs::write(dir.join("backup.json"), header(&request, "meta")?).map_err(|e| e.to_string())
}

#[tauri::command]
fn backup_meta(app: AppHandle) -> Option<String> {
    let dir = backup_dir(&app).ok()?;
    dir.join("backup.docx").exists().then_some(())?;
    std::fs::read_to_string(dir.join("backup.json")).ok()
}

#[tauri::command]
fn backup_data(app: AppHandle) -> Result<Response, String> {
    std::fs::read(backup_dir(&app)?.join("backup.docx"))
        .map(Response::new)
        .map_err(|e| e.to_string())
}

// Version history: one folder per document (keyed by a hash of its path), one file per snapshot
const MAX_VERSIONS: usize = 100;

fn versions_dir(app: &AppHandle, key: &str) -> Result<PathBuf, String> {
    if key.is_empty() || key.len() > 32 || !key.chars().all(|c| c.is_ascii_hexdigit()) {
        return Err("bad key".into());
    }
    let dir = backup_dir(app)?.join("versions").join(key);
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir)
}

fn version_times(dir: &PathBuf) -> Vec<u64> {
    let mut times: Vec<u64> = std::fs::read_dir(dir)
        .map(|it| {
            it.flatten()
                .filter_map(|e| e.file_name().to_str()?.strip_suffix(".docx")?.parse().ok())
                .collect()
        })
        .unwrap_or_default();
    times.sort_unstable();
    times
}

#[tauri::command]
fn version_save(app: AppHandle, request: Request) -> Result<(), String> {
    let dir = versions_dir(&app, &header(&request, "key")?)?;
    let ts: u64 = header(&request, "ts")?.parse().map_err(|_| "bad ts")?;
    write_atomic(
        &dir.join(format!("{ts}.docx")).to_string_lossy(),
        body(&request)?,
    )?;
    let times = version_times(&dir);
    for old in &times[..times.len().saturating_sub(MAX_VERSIONS)] {
        let _ = std::fs::remove_file(dir.join(format!("{old}.docx")));
    }
    Ok(())
}

#[tauri::command]
fn version_list(app: AppHandle, key: String) -> Result<Vec<u64>, String> {
    Ok(version_times(&versions_dir(&app, &key)?))
}

#[tauri::command]
fn version_read(app: AppHandle, key: String, ts: u64) -> Result<Response, String> {
    std::fs::read(versions_dir(&app, &key)?.join(format!("{ts}.docx")))
        .map(Response::new)
        .map_err(|e| e.to_string())
}

#[tauri::command]
fn backup_clear(app: AppHandle) {
    if let Ok(dir) = backup_dir(&app) {
        let _ = std::fs::remove_file(dir.join("backup.json"));
        let _ = std::fs::remove_file(dir.join("backup.docx"));
    }
}

// Mica (the frosted window background) exists from Windows 11, build 22000; Windows 10 gets a normal solid window
#[cfg(windows)]
fn has_mica() -> bool {
    use windows::core::w;
    use windows::Win32::System::Registry::{RegGetValueW, HKEY_LOCAL_MACHINE, RRF_RT_REG_SZ};
    let mut buf = [0u16; 16];
    let mut size = (buf.len() * 2) as u32;
    let status = unsafe {
        RegGetValueW(
            HKEY_LOCAL_MACHINE,
            w!(r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"),
            w!("CurrentBuildNumber"),
            RRF_RT_REG_SZ,
            None,
            Some(buf.as_mut_ptr().cast()),
            Some(&mut size),
        )
    };
    if status.0 != 0 {
        return false;
    }
    let chars = (size as usize / 2).saturating_sub(1);
    String::from_utf16_lossy(&buf[..chars])
        .trim()
        .parse::<u32>()
        .is_ok_and(|build| build >= 22000)
}

#[cfg(not(windows))]
fn has_mica() -> bool {
    false
}

// Builds before the rename used another app id; move their data (settings, recent files, version history) over once
fn migrate_old_data() {
    let Some(base) = std::env::var_os("LOCALAPPDATA").map(PathBuf::from) else {
        return;
    };
    let old = base.join("com.local.docviewer");
    let new = base.join("io.github.stefanrotari99i.ribbonless");
    if old.is_dir() && !new.exists() {
        let _ = std::fs::rename(old, new);
    }
}

fn main() {
    migrate_old_data();
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_clipboard_manager::init())
        .plugin(tauri_plugin_opener::init())
        .setup(|app| {
            // The window is made here rather than from the config, so transparency is only on where Mica can fill it
            let mica = has_mica();
            let mut config = app.config().app.windows[0].clone();
            config.transparent = mica;
            if !mica {
                config.window_effects = None;
            }
            let win = tauri::WebviewWindowBuilder::from_config(app.handle(), &config)?
                .initialization_script(if mica {
                    "window.__RIBBONLESS_MICA__ = true;"
                } else {
                    ""
                })
                .build()?;
            // The page shows the window once styled; this is the fallback if it never does
            std::thread::spawn(move || {
                std::thread::sleep(std::time::Duration::from_secs(3));
                let _ = win.show();
            });
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            opened_path,
            read_file,
            write_file,
            backup_write,
            backup_meta,
            backup_data,
            backup_clear,
            spell_suggest,
            spell_add,
            fonts,
            version_save,
            version_list,
            version_read
        ])
        .run(tauri::generate_context!())
        .expect("error while running app");
}
