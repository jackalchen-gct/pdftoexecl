use serde::Serialize;
use std::path::{Path, PathBuf};
use std::process::Command;
use tauri::Manager;

#[derive(Serialize)]
struct ConvertResult {
    input: String,
    output: String,
    status: String,
    message: String,
}

fn python_executable() -> String {
    std::env::var("PDFTOEXECL_PYTHON").unwrap_or_else(|_| "python".to_string())
}

fn sidecar_binary_name() -> String {
    match std::env::consts::OS {
        "windows" => format!("converter-sidecar-{}-pc-windows-msvc.exe", std::env::consts::ARCH),
        "macos" => format!("converter-sidecar-{}-apple-darwin", std::env::consts::ARCH),
        "linux" => format!("converter-sidecar-{}-unknown-linux-gnu", std::env::consts::ARCH),
        _ => "converter-sidecar".to_string(),
    }
}

fn bundled_sidecar(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    let path = app
        .path()
        .resolve(
            format!("bin/{}", sidecar_binary_name()),
            tauri::path::BaseDirectory::Resource,
        )
        .map_err(|error| error.to_string())?;

    if path.exists() {
        Ok(path)
    } else {
        Err("Bundled sidecar not found".to_string())
    }
}

fn local_sidecar() -> Result<PathBuf, String> {
    let cwd = std::env::current_dir().map_err(|error| error.to_string())?;
    let sidecar_name = sidecar_binary_name();
    for candidate in [
        cwd.join("src-tauri").join("bin").join(&sidecar_name),
        cwd.parent()
            .unwrap_or(&cwd)
            .join("src-tauri")
            .join("bin")
            .join(&sidecar_name),
    ] {
        if candidate.exists() {
            return Ok(candidate);
        }
    }

    Err("Local sidecar not found".to_string())
}

fn converter_script(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    let exe = std::env::current_exe().map_err(|error| error.to_string())?;
    let exe_dir = exe.parent().ok_or("Cannot locate executable directory")?;
    let bundled = exe_dir.join("converter").join("convert_pdf.py");
    if bundled.exists() {
        return Ok(bundled);
    }

    let resource = app
        .path()
        .resolve("converter/convert_pdf.py", tauri::path::BaseDirectory::Resource)
        .unwrap_or_else(|_| PathBuf::from("converter/convert_pdf.py"));
    if resource.exists() {
        return Ok(resource);
    }

    let cwd = std::env::current_dir().map_err(|error| error.to_string())?;
    for candidate in [
        cwd.join("converter").join("convert_pdf.py"),
        cwd.parent()
            .unwrap_or(&cwd)
            .join("converter")
            .join("convert_pdf.py"),
    ] {
        if candidate.exists() {
            return Ok(candidate);
        }
    }

    Err("Cannot locate converter/convert_pdf.py".to_string())
}

enum ConverterBackend {
    Sidecar(PathBuf),
    Python(PathBuf),
}

fn resolve_backend(app: &tauri::AppHandle) -> Result<ConverterBackend, String> {
    if let Ok(path) = bundled_sidecar(app) {
        return Ok(ConverterBackend::Sidecar(path));
    }

    if let Ok(path) = local_sidecar() {
        return Ok(ConverterBackend::Sidecar(path));
    }

    if cfg!(debug_assertions) {
        return Ok(ConverterBackend::Python(converter_script(app)?));
    }

    Err("Bundled sidecar not found".to_string())
}

#[tauri::command]
async fn convert_pdfs(
    app: tauri::AppHandle,
    pdf_paths: Vec<String>,
    output_dir: String,
) -> Result<Vec<ConvertResult>, String> {
    let backend = resolve_backend(&app)?;
    let output_dir_path = Path::new(&output_dir);
    let mut results = Vec::new();

    for input in pdf_paths {
        let input_path = Path::new(&input);
        let output_path = output_dir_path.join(
            input_path
                .file_stem()
                .ok_or("Invalid PDF file name")?
                .to_string_lossy()
                .to_string()
                + ".xlsx",
        );

        let output = match &backend {
            ConverterBackend::Sidecar(binary) => Command::new(binary)
                .arg(&input)
                .arg(&output_path)
                .output()
                .map_err(|error| error.to_string())?,
            ConverterBackend::Python(script) => Command::new(python_executable())
                .arg(script)
                .arg(&input)
                .arg(&output_path)
                .output()
                .map_err(|error| error.to_string())?,
        };

        if output.status.success() {
            results.push(ConvertResult {
                input,
                output: output_path.to_string_lossy().to_string(),
                status: "success".to_string(),
                message: "轉換成功".to_string(),
            });
        } else {
            results.push(ConvertResult {
                input,
                output: output_path.to_string_lossy().to_string(),
                status: "failed".to_string(),
                message: String::from_utf8_lossy(&output.stderr).trim().to_string(),
            });
        }
    }

    Ok(results)
}

pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![convert_pdfs])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
