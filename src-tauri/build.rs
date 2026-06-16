use std::fs;
use std::path::Path;

fn ensure_windows_icon() {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    let local_icon = manifest_dir.join("icons").join("icon.ico");
    let source_icon = manifest_dir
        .parent()
        .unwrap_or(manifest_dir)
        .join("icons")
        .join("icon.ico");
    let target_icon_dir = manifest_dir.join("icons");
    let target_icon = target_icon_dir.join("icon.ico");

    let source = if local_icon.exists() {
        Some(local_icon)
    } else if source_icon.exists() {
        Some(source_icon)
    } else {
        None
    };

    if let Some(source_icon) = source {
        let _ = fs::create_dir_all(&target_icon_dir);
        let needs_copy = match fs::metadata(&target_icon) {
            Ok(target_meta) => match fs::metadata(&source_icon) {
                Ok(source_meta) => match (source_meta.modified(), target_meta.modified()) {
                    (Ok(source_time), Ok(target_time)) => source_time > target_time,
                    _ => true,
                },
                Err(_) => true,
            },
            Err(_) => true,
        };

        if needs_copy {
            let _ = fs::copy(source_icon, target_icon);
        }
    }
}

fn main() {
    ensure_windows_icon();
    tauri_build::build()
}
