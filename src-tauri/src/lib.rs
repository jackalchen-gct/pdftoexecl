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

#[tauri::command]
async fn convert_pdfs(
    app: tauri::AppHandle,
    pdf_paths: Vec<String>,
    output_dir: String,
) -> Result<Vec<ConvertResult>, String> {
    let script = converter_script(&app)?;
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

        let output = Command::new(python_executable())
            .arg(&script)
            .arg(&input)
            .arg(&output_path)
            .output()
            .map_err(|error| error.to_string())?;

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
