use serde::{Deserialize, Serialize};
use std::path::{Path, PathBuf};
use std::process::Command;
use tauri::Manager;
use std::sync::Mutex;
use std::io::BufRead;

struct DaemonState {
    port: u16,
    child: Mutex<std::process::Child>,
}

impl Drop for DaemonState {
    fn drop(&mut self) {
        use std::io::Write;
        if let Ok(mut stream) = std::net::TcpStream::connect(format!("127.0.0.1:{}", self.port)) {
            let cmd = serde_json::json!({ "action": "shutdown" });
            if let Ok(cmd_str) = serde_json::to_string(&cmd) {
                let _ = stream.write_all((cmd_str + "\n").as_bytes());
            }
        }
        if let Ok(mut child) = self.child.lock() {
            let _ = child.kill();
        }
    }
}

fn get_free_port() -> Result<u16, String> {
    let listener = std::net::TcpListener::bind("127.0.0.1:0").map_err(|e| e.to_string())?;
    let port = listener.local_addr().map_err(|e| e.to_string())?.port();
    Ok(port)
}

#[derive(Serialize)]
struct ConvertResult {
    input: String,
    output: String,
    status: String,
    message: String,
    table_count: usize,
    tables: Vec<ExtractedTable>,
    pages: Vec<ExtractedPage>,
    logs: Vec<String>,
}

#[derive(Serialize, Deserialize, Clone)]
struct ExtractedTable {
    page: usize,
    index: usize,
    title: String,
    rows: Vec<Vec<String>>,
    ratio: Option<String>,
    formula: Option<String>,
}

#[derive(Serialize, Deserialize, Clone)]
struct ExtractedPage {
    page: usize,
    text: String,
    thumbnail: String,
}

#[derive(Deserialize)]
struct ConverterPayload {
    input: String,
    output: String,
    status: Option<String>,
    message: Option<String>,
    table_count: usize,
    tables: Vec<ExtractedTable>,
    pages: Vec<ExtractedPage>,
    logs: Option<Vec<String>>,
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

fn parse_payload(stdout: &[u8]) -> Result<ConverterPayload, String> {
    let text = String::from_utf8_lossy(stdout).trim().to_string();
    if text.is_empty() {
        return Err("Converter returned no payload".to_string());
    }
    serde_json::from_str(&text).map_err(|error| error.to_string())
}

#[tauri::command]
async fn convert_pdfs(
    state: tauri::State<'_, DaemonState>,
    pdf_paths: Vec<String>,
    output_dir: String,
    page_selections: std::collections::HashMap<String, Vec<usize>>,
    project_name: Option<String>,
) -> Result<Vec<ConvertResult>, String> {
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

        let pages = page_selections.get(&input).cloned();

        use std::io::{Write, Read};
        let mut stream = std::net::TcpStream::connect(format!("127.0.0.1:{}", state.port))
            .map_err(|e| format!("Failed to connect to daemon: {}", e))?;

        let cmd = serde_json::json!({
            "action": "convert",
            "pdf_path": input.clone(),
            "output_path": output_path.to_string_lossy().to_string(),
            "pages": pages,
            "project_name": project_name,
        });

        let cmd_str = serde_json::to_string(&cmd).map_err(|e| e.to_string())? + "\n";
        stream.write_all(cmd_str.as_bytes()).map_err(|e| e.to_string())?;
        stream.flush().map_err(|e| e.to_string())?;

        let mut response = String::new();
        stream.read_to_string(&mut response).map_err(|e| e.to_string())?;

        match parse_payload(response.as_bytes()) {
            Ok(payload) => {
                let status = payload.status.clone().unwrap_or_else(|| "success".to_string());
                let message = if status == "error" {
                    payload.message.clone().unwrap_or_else(|| "轉換失敗".to_string())
                } else {
                    "轉換成功".to_string()
                };
                results.push(ConvertResult {
                    input: payload.input,
                    output: payload.output,
                    status,
                    message,
                    table_count: payload.table_count,
                    tables: payload.tables,
                    pages: payload.pages,
                    logs: payload.logs.unwrap_or_default(),
                });
            }
            Err(e) => {
                results.push(ConvertResult {
                    input,
                    output: output_path.to_string_lossy().to_string(),
                    status: "failed".to_string(),
                    message: format!("解析 JSON 響應失敗: {}, 原始回應: {}", e, response),
                    table_count: 0,
                    tables: Vec::new(),
                    pages: Vec::new(),
                    logs: Vec::new(),
                });
            }
        }
    }

    Ok(results)
}

#[tauri::command]
async fn get_pdf_previews(
    state: tauri::State<'_, DaemonState>,
    pdf_paths: Vec<String>,
) -> Result<Vec<ConvertResult>, String> {
    let mut results = Vec::new();

    for input in pdf_paths {
        use std::io::{Write, Read};
        let mut stream = std::net::TcpStream::connect(format!("127.0.0.1:{}", state.port))
            .map_err(|e| format!("Failed to connect to daemon: {}", e))?;

        let cmd = serde_json::json!({
            "action": "get_thumbnails",
            "pdf_path": input.clone(),
        });

        let cmd_str = serde_json::to_string(&cmd).map_err(|e| e.to_string())? + "\n";
        stream.write_all(cmd_str.as_bytes()).map_err(|e| e.to_string())?;
        stream.flush().map_err(|e| e.to_string())?;

        let mut response = String::new();
        stream.read_to_string(&mut response).map_err(|e| e.to_string())?;

        match parse_payload(response.as_bytes()) {
            Ok(payload) => {
                let status = payload.status.clone().unwrap_or_else(|| "success".to_string());
                let message = if status == "error" {
                    payload.message.clone().unwrap_or_else(|| "載入失敗".to_string())
                } else {
                    "載入成功".to_string()
                };
                results.push(ConvertResult {
                    input: payload.input,
                    output: payload.output,
                    status,
                    message,
                    table_count: payload.table_count,
                    tables: payload.tables,
                    pages: payload.pages,
                    logs: payload.logs.unwrap_or_default(),
                });
            }
            Err(e) => {
                results.push(ConvertResult {
                    input,
                    output: "".to_string(),
                    status: "failed".to_string(),
                    message: format!("解析 JSON 響應失敗: {}, 原始回應: {}", e, response),
                    table_count: 0,
                    tables: Vec::new(),
                    pages: Vec::new(),
                    logs: Vec::new(),
                });
            }
        }
    }

    Ok(results)
}

#[tauri::command]
async fn get_master_sheets(
    state: tauri::State<'_, DaemonState>,
    output_dir: String,
) -> Result<Vec<String>, String> {
    use std::io::{Write, Read};
    let mut stream = std::net::TcpStream::connect(format!("127.0.0.1:{}", state.port))
        .map_err(|e| format!("Failed to connect to daemon: {}", e))?;

    let cmd = serde_json::json!({
        "action": "get_master_sheets",
        "output_dir": output_dir,
    });

    let cmd_str = serde_json::to_string(&cmd).map_err(|e| e.to_string())? + "\n";
    stream.write_all(cmd_str.as_bytes()).map_err(|e| e.to_string())?;
    stream.flush().map_err(|e| e.to_string())?;

    let mut response = String::new();
    stream.read_to_string(&mut response).map_err(|e| e.to_string())?;

    let payload: serde_json::Value = serde_json::from_str(&response).map_err(|e| e.to_string())?;
    if let Some(sheets) = payload.get("sheets").and_then(|s| s.as_array()) {
        let sheet_names: Vec<String> = sheets.iter().filter_map(|s| s.as_str().map(|str| str.to_string())).collect();
        Ok(sheet_names)
    } else {
        Ok(Vec::new())
    }
}

pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .setup(|app| {
            let port = get_free_port().map_err(|e| e.to_string())?;
            let backend = resolve_backend(app.handle())?;
            let mut child = match &backend {
                ConverterBackend::Sidecar(binary) => {
                    Command::new(binary)
                        .arg("--daemon")
                        .arg(port.to_string())
                        .stdout(std::process::Stdio::piped())
                        .spawn()
                        .map_err(|e| e.to_string())?
                }
                ConverterBackend::Python(script) => {
                    Command::new(python_executable())
                        .arg(script)
                        .arg("--daemon")
                        .arg(port.to_string())
                        .stdout(std::process::Stdio::piped())
                        .spawn()
                        .map_err(|e| e.to_string())?
                }
            };
            
            let stdout = child.stdout.as_mut().ok_or("Failed to capture sidecar stdout")?;
            let mut reader = std::io::BufReader::new(stdout);
            let mut ready_line = String::new();
            reader.read_line(&mut ready_line).map_err(|e| e.to_string())?;
            
            if !ready_line.contains("ready") {
                return Err(format!("Daemon startup failed, output: {}", ready_line).into());
            }
            
            app.manage(DaemonState {
                port,
                child: Mutex::new(child),
            });
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![convert_pdfs, get_pdf_previews, get_master_sheets])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
