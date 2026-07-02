import { invoke } from "@tauri-apps/api/core";
import { open } from "@tauri-apps/plugin-dialog";
import { listen } from "@tauri-apps/api/event";
import { useCallback, useEffect, useMemo, useState } from "react";
import StatsHero from "./components/StatsHero";
import VersionModal from "./components/VersionModal";
import PreviewModal from "./components/PreviewModal";

type ExtractedTable = {
  page: number;
  index: number;
  title: string;
  rows: string[][];
  ratio?: string;
  formula?: string;
};

type ExtractedPage = {
  page: number;
  text: string;
  thumbnail: string;
};

type ConvertResult = {
  input: string;
  output: string;
  status: "success" | "failed";
  message: string;
  table_count: number;
  tables: ExtractedTable[];
  pages: ExtractedPage[];
  logs: string[];
};

function fileName(path: string) {
  return path.split(/[\\/]/).pop() ?? path;
}

function previewPath(path: string, maxLength = 64) {
  if (path.length <= maxLength) return path;
  return `…${path.slice(path.length - maxLength)}`;
}

function nonEmptyText(text: string) {
  return text.trim().length > 0;
}

type PreviewTarget = {
  title: string;
  page: ExtractedPage;
};

export default function App() {
  const [pdfs, setPdfs] = useState<string[]>([]);
  const [outputDir, setOutputDir] = useState<string>("");
  const [projectName, setProjectName] = useState<string>("");
  const [existingSheets, setExistingSheets] = useState<string[]>([]);
  const [overwrite, setOverwrite] = useState<boolean>(false);
  const [busy, setBusy] = useState(false);
  const [results, setResults] = useState<ConvertResult[]>([]);
  const [pdfPreviews, setPdfPreviews] = useState<ConvertResult[]>([]);
  const [loadingPreviews, setLoadingPreviews] = useState(false);
  const [selectedPages, setSelectedPages] = useState<Record<string, number[]>>({});
  const [preview, setPreview] = useState<PreviewTarget | null>(null);
  const [globalError, setGlobalError] = useState<string | null>(null);
  const [showVersionInfo, setShowVersionInfo] = useState(false);

  const successCount = useMemo(
    () => results.filter((result) => result.status === "success").length,
    [results],
  );

  const tableCount = useMemo(
    () => results.reduce((sum, result) => sum + result.table_count, 0),
    [results],
  );

  const pageCount = useMemo(
    () => pdfPreviews.reduce((sum, preview) => sum + preview.pages.length, 0),
    [pdfPreviews],
  );

  const hasUnselectedMultiPagePdfs = useMemo(() => {
    return pdfs.some((path) => {
      const preview = pdfPreviews.find((p) => p.input === path);
      if (!preview) return false;
      if (preview.pages.length > 1) {
        return !selectedPages[path] || selectedPages[path].length === 0;
      }
      return false;
    });
  }, [pdfs, pdfPreviews, selectedPages]);

  const [isDragging, setIsDragging] = useState(false);
  const [debugLogs, setDebugLogs] = useState<string[]>([]);

  const loadPdfPaths = useCallback(async (paths: string[]) => {
    setPdfs(paths);
    setResults([]);
    setPreview(null);
    setPdfPreviews([]);
    setSelectedPages({});
    setGlobalError(null);

    // Auto-fill project name from first PDF filename if currently empty
    if (paths.length > 0 && projectName.trim() === "") {
      const fullPath = paths[0];
      const fileName = fullPath.split(/[\\/]/).pop() || "";
      const defaultName = getDefaultProjectName(fileName);
      setProjectName(defaultName);
    }

    setLoadingPreviews(true);
    try {
      const previews = await invoke<ConvertResult[]>("get_pdf_previews", {
        pdfPaths: paths,
      });
      setPdfPreviews(previews);

      // Default unchecked for multi-page PDFs, auto-select page 1 for single-page PDFs
      const initialSelections: Record<string, number[]> = {};
      previews.forEach((p) => {
        if (p.pages.length === 1) {
          initialSelections[p.input] = [1];
        } else {
          initialSelections[p.input] = [];
        }
      });
      setSelectedPages(initialSelections);
    } catch (e) {
      console.error("Failed to load PDF previews:", e);
      setGlobalError(`載入 PDF 預覽失敗：${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setLoadingPreviews(false);
    }
  }, []);

  async function choosePdfs() {
    const selected = await open({
      multiple: true,
      filters: [{ name: "PDF", extensions: ["pdf"] }],
    });
    if (!selected) return;
    const paths = Array.isArray(selected) ? selected : [selected];
    await loadPdfPaths(paths);
  }

  useEffect(() => {
    const logMessage = (msg: string) => {
      setDebugLogs((prev) => [`[${new Date().toLocaleTimeString()}] ${msg}`, ...prev.slice(0, 49)]);
    };

    const handleDragOver = (e: DragEvent) => {
      e.preventDefault();
      if (e.dataTransfer) {
        e.dataTransfer.dropEffect = "copy";
      }
    };
    const handleDragEnter = (e: DragEvent) => {
      e.preventDefault();
      if (e.dataTransfer) {
        e.dataTransfer.dropEffect = "copy";
      }
      logMessage("DOM Event: window dragenter triggered");
    };
    const handleDragLeave = (e: DragEvent) => {
      logMessage("DOM Event: window dragleave triggered");
    };
    const handleDrop = (e: DragEvent) => {
      e.preventDefault();
      logMessage("DOM Event: window drop triggered");
    };

    window.addEventListener("dragover", handleDragOver);
    window.addEventListener("dragenter", handleDragEnter);
    window.addEventListener("dragleave", handleDragLeave);
    window.addEventListener("drop", handleDrop);

    let active = true;
    const promises = [
      listen<{ paths: string[] }>("tauri://drag-drop", (event) => {
        setIsDragging(false);
        const paths = event.payload?.paths;
        logMessage(`Tauri Event: tauri://drag-drop, paths: ${JSON.stringify(paths)}`);
        if (paths && paths.length > 0) {
          const pdfPaths = paths.filter((p) => p.toLowerCase().endsWith(".pdf"));
          if (pdfPaths.length > 0) {
            loadPdfPaths(pdfPaths);
          }
        }
      }),
      listen("tauri://drag-enter", (event) => {
        setIsDragging(true);
        logMessage(`Tauri Event: tauri://drag-enter, payload: ${JSON.stringify(event.payload)}`);
      }),
      listen("tauri://drag-leave", (event) => {
        setIsDragging(false);
        logMessage(`Tauri Event: tauri://drag-leave, payload: ${JSON.stringify(event.payload)}`);
      })
    ];

    let unlisteners: (() => void)[] = [];

    Promise.all(promises).then((fns) => {
      if (active) {
        unlisteners = fns;
      } else {
        fns.forEach((fn) => fn());
      }
    });

    return () => {
      active = false;
      unlisteners.forEach((fn) => fn());
      window.removeEventListener("dragover", handleDragOver);
      window.removeEventListener("dragenter", handleDragEnter);
      window.removeEventListener("dragleave", handleDragLeave);
      window.removeEventListener("drop", handleDrop);
    };
  }, [loadPdfPaths]);


  const togglePageSelection = (pdfPath: string, pageNum: number) => {
    setSelectedPages((prev) => {
      const current = prev[pdfPath] ?? [];
      if (current.includes(pageNum)) {
        return { ...prev, [pdfPath]: current.filter((p) => p !== pageNum) };
      } else {
        return { ...prev, [pdfPath]: [...current, pageNum] };
      }
    });
  };

  const checkExistingSheets = useCallback(async (dir: string) => {
    if (!dir) {
      setExistingSheets([]);
      return;
    }
    try {
      const sheets = await invoke<string[]>("get_master_sheets", { outputDir: dir });
      setExistingSheets(sheets);
    } catch (e) {
      console.error("Failed to fetch master sheets:", e);
    }
  }, []);

  useEffect(() => {
    checkExistingSheets(outputDir);
  }, [outputDir, checkExistingSheets]);

  async function chooseOutputDir() {
    const selected = await open({ directory: true, multiple: false });
    if (typeof selected === "string") {
      setOutputDir(selected);
    }
  }

  async function convert() {
    if (!pdfs.length || !outputDir) return;
    setBusy(true);
    setResults([]);
    setGlobalError(null);
    try {
      const converted = await invoke<ConvertResult[]>("convert_pdfs", {
        pdfPaths: pdfs,
        outputDir,
        pageSelections: selectedPages,
        projectName: projectName.trim() || null,
      });
      setResults(converted);
      setPreview(null);
      await checkExistingSheets(outputDir);
    } catch (e) {
      console.error("Conversion failed:", e);
      setGlobalError(`轉換失敗：${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setBusy(false);
    }
  }

  return (
    <main className="shell">
      <div className="ambient ambient-one" />
      <div className="ambient ambient-two" />

      <header style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.5rem" }}>
        <h1 style={{ fontSize: "1.8rem", margin: 0, fontWeight: 900, color: "#1d3b2a", maxWidth: "none", whiteSpace: "nowrap" }}>
          PDF to Excel Converter
        </h1>

        <span 
          className="badge neutral" 
          style={{ fontSize: "0.85rem", padding: "4px 10px", cursor: "pointer", userSelect: "none" }}
          onClick={() => setShowVersionInfo(true)}
        >
          v0.6.0
        </span>
      </header>

      <StatsHero
        pdfsCount={pdfs.length}
        successCount={successCount}
        tableCount={tableCount}
        pageCount={pageCount}
      />

      {globalError && (
        <div style={{
          background: "rgba(171, 54, 37, 0.12)",
          border: "1px solid rgba(171, 54, 37, 0.28)",
          borderRadius: "18px",
          padding: "14px 18px",
          color: "#7f2717",
          fontWeight: 700,
          marginBottom: "16px",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          fontSize: "0.95rem"
        }}>
          <span>{globalError}</span>
          <button 
            onClick={() => setGlobalError(null)}
            style={{
              padding: "4px 12px",
              fontSize: "0.82rem",
              background: "rgba(255, 255, 255, 0.6)",
              border: "1px solid rgba(171, 54, 37, 0.18)",
              height: "auto"
            }}
          >
            關閉
          </button>
        </div>
      )}

      <section className="panel command-panel" style={{ paddingBottom: "24px" }}>
        <div className="grid">
          {/* 輸入 PDF Card */}
          <div 
            className={`card interactive drag-drop-zone ${isDragging ? "drag-over" : ""}`}
            onClick={choosePdfs}
          >
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "12px" }}>
              <h2 style={{ margin: 0 }}>輸入 PDF</h2>
              {pdfs.length > 0 && <span style={{ fontSize: "0.82rem", color: "#2c724b", fontWeight: "bold" }}>🔄 點擊重選</span>}
            </div>
            {isDragging ? (
              <div className="drag-active-overlay">
                <div className="drag-active-inner">
                  <span className="drag-icon">📥</span>
                  <strong>放開以載入 PDF 檔案</strong>
                </div>
              </div>
            ) : pdfs.length ? (
              <ul className="file-list">
                {pdfs.map((path) => (
                  <li key={path}>
                    <strong>{fileName(path)}</strong>
                    <span>{previewPath(path)}</span>
                  </li>
                ))}
              </ul>
            ) : (
              <div className="drag-placeholder">
                <span style={{ fontSize: "1.8rem", marginBottom: "8px" }}>📄</span>
                <p className="muted" style={{ fontWeight: 600 }}>尚未選擇 PDF。</p>
                <p className="muted-hint" style={{ fontSize: "0.82rem", marginTop: "4px", opacity: 0.7 }}>
                  拖曳 PDF 檔案至此，或點擊此區域選擇檔案
                </p>
              </div>
            )}
          </div>

          {/* 輸出位置 Card */}
          <div className="card interactive" onClick={chooseOutputDir}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "12px" }}>
              <h2 style={{ margin: 0 }}>輸出位置</h2>
              {outputDir && <span style={{ fontSize: "0.82rem", color: "#2c724b", fontWeight: "bold" }}>📁 點擊變更</span>}
            </div>
            {outputDir ? (
              <div style={{ display: "flex", flexDirection: "column", gap: "8px", justifyContent: "center", minHeight: "110px", padding: "12px", background: "rgba(255, 250, 240, 0.5)", border: "1px solid rgba(47, 42, 31, 0.06)", borderRadius: "18px" }}>
                <span style={{ fontSize: "1.5rem" }}>📂</span>
                <code className="path" style={{ fontSize: "0.88rem", color: "#1d3b2a", wordBreak: "break-all" }}>{outputDir}</code>
              </div>
            ) : (
              <div className="drag-placeholder">
                <span style={{ fontSize: "1.8rem", marginBottom: "8px" }}>📁</span>
                <p className="muted" style={{ fontWeight: 600 }}>尚未選擇輸出資料夾。</p>
                <p className="muted-hint" style={{ fontSize: "0.82rem", marginTop: "4px", opacity: 0.7 }}>
                  點擊此區域選擇 Excel 檔案輸出路徑
                </p>
              </div>
            )}
          </div>

          {/* 專案名稱 Card */}
          <div className="card" onClick={(e) => e.stopPropagation()}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "12px" }}>
              <h2 style={{ margin: 0 }}>專案名稱</h2>
              {projectName.trim() && (
                <span 
                  style={{ fontSize: "0.82rem", color: "#6d6254", cursor: "pointer", fontWeight: "bold" }}
                  onClick={() => {
                    setProjectName("");
                    setOverwrite(false);
                  }}
                >
                  清除 ❌
                </span>
              )}
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: "10px", justifyContent: "center", minHeight: "110px", padding: "8px 12px" }}>
              <input
                type="text"
                placeholder="輸入專案名稱以命名 Excel 及 Sheet..."
                value={projectName}
                onChange={(e) => {
                  setProjectName(e.target.value);
                  setOverwrite(false); // Reset overwrite checkbox when name changes
                }}
                style={{
                  width: "100%",
                  padding: "12px 14px",
                  fontSize: "0.95rem",
                  borderRadius: "14px",
                  border: projectName.trim() && existingSheets.includes(projectName.trim()) && !overwrite ? "2px solid #ab3625" : "1px solid rgba(47, 42, 31, 0.2)",
                  background: "rgba(255, 255, 255, 0.95)",
                  color: "#1d3b2a",
                  outline: "none",
                  fontWeight: 600,
                  boxShadow: "inset 0 2px 4px rgba(0,0,0,0.02)"
                }}
              />
              {projectName.trim() && existingSheets.includes(projectName.trim()) ? (
                <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: "4px", color: "#ab3625", fontSize: "0.82rem", fontWeight: "bold" }}>
                    <span>⚠️ 此專案名稱已存在於匯總管理表！</span>
                  </div>
                  <label style={{ display: "flex", alignItems: "center", gap: "6px", cursor: "pointer", fontSize: "0.85rem", color: "#ab3625", fontWeight: "bold", userSelect: "none" }}>
                    <input
                      type="checkbox"
                      checked={overwrite}
                      onChange={(e) => setOverwrite(e.target.checked)}
                      style={{ cursor: "pointer" }}
                    />
                    <span>覆蓋同名專案</span>
                  </label>
                </div>
              ) : (
                <div style={{ color: "#6d6254", fontSize: "0.82rem", opacity: 0.8 }}>
                  {projectName.trim() ? "✨ 專案名稱可用" : "💡 設定後將於 Row 1 新增專案標題"}
                </div>
              )}
            </div>
          </div>
        </div>

        {/* 轉成 Excel 開始轉換按鈕 */}
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", marginTop: "24px" }}>
          <button 
            className="primary" 
            onClick={(e) => {
              e.stopPropagation();
              convert();
            }} 
            disabled={busy || !pdfs.length || !outputDir || (projectName.trim() !== "" && existingSheets.includes(projectName.trim()) && !overwrite) || hasUnselectedMultiPagePdfs}
            style={{ 
              width: "100%", 
              maxWidth: "480px", 
              padding: "1rem 2rem", 
              fontSize: "1.1rem",
              borderRadius: "999px",
              boxShadow: !pdfs.length || !outputDir || (projectName.trim() !== "" && existingSheets.includes(projectName.trim()) && !overwrite) || hasUnselectedMultiPagePdfs ? "none" : "0 8px 26px rgba(44, 114, 75, 0.22)"
            }}
          >
            {busy ? (
              <span style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "8px" }}>
                <span className="spinner" /> 正在轉換為 Excel 中...
              </span>
            ) : (
              "🚀 開始轉換為 Excel"
            )}
          </button>
          {hasUnselectedMultiPagePdfs && (
            <div style={{ color: "#ab3625", fontSize: "0.9rem", fontWeight: "bold", marginTop: "10px", textAlign: "center" }}>
              ⚠️ 請在下方預覽區勾選多頁 PDF 要轉換的頁數！
            </div>
          )}
        </div>
      </section>

      {/* PDF Page Previews (shown immediately after loading PDFs) */}
      {(loadingPreviews || pdfPreviews.length > 0) && (
        <section className="panel">
          <div className="section-head">
            <div>
              <h2>PDF Previews</h2>
            </div>
            {loadingPreviews && <span className="badge neutral">載入縮圖中...</span>}
          </div>

          <div className="result-stack">
            {pdfPreviews.map((preview) => (
              <div key={preview.input} className="preview-group" style={{ marginBottom: "1.5rem" }}>
                <h3 style={{ margin: "0 0 12px 0", fontSize: "1.1rem" }}>{fileName(preview.input)}</h3>
                {preview.pages.length === 1 ? (
                  <div className="single-page-preview" style={{ maxWidth: "480px", margin: "0 auto", textAlign: "center" }}>
                    <div
                      className="page-thumb"
                      style={{ aspectRatio: "auto", height: "auto", cursor: "pointer", borderRadius: "16px", overflow: "hidden" }}
                      onClick={() => setPreview({ title: fileName(preview.input), page: preview.pages[0] })}
                    >
                      {preview.pages[0].thumbnail ? (
                        <img
                          src={preview.pages[0].thumbnail}
                          alt="Page preview"
                          style={{ width: "100%", height: "auto", display: "block", background: "#fffdf7", border: "1px solid rgba(47, 42, 31, 0.12)", borderRadius: "16px" }}
                        />
                      ) : (
                        <div className="page-thumb-fallback" style={{ padding: "40px" }}>
                          <strong>Page 1</strong>
                          <span>無縮圖資料</span>
                        </div>
                      )}
                    </div>
                    <div style={{ marginTop: "10px" }}>
                      <span className="badge neutral">單頁 PDF（自動選取轉換）</span>
                    </div>
                  </div>
                ) : preview.pages.length > 1 ? (
                  <div>
                    <div className="page-control-row" style={{ display: "flex", gap: "10px", marginBottom: "12px" }}>
                      <button
                        type="button"
                        style={{ padding: "6px 14px", fontSize: "0.85rem", height: "auto", borderRadius: "10px" }}
                        onClick={() => {
                          setSelectedPages((prev) => ({
                            ...prev,
                            [preview.input]: preview.pages.map((p) => p.page),
                          }));
                        }}
                      >
                        全選 ({preview.pages.length} 頁)
                      </button>
                      <button
                        type="button"
                        style={{ padding: "6px 14px", fontSize: "0.85rem", height: "auto", borderRadius: "10px" }}
                        onClick={() => {
                          setSelectedPages((prev) => ({
                            ...prev,
                            [preview.input]: [],
                          }));
                        }}
                      >
                        全取消勾選
                      </button>
                    </div>
                    <div className="page-grid">
                      {preview.pages.map((page) => {
                        const isSelected = selectedPages[preview.input]?.includes(page.page) ?? false;
                        return (
                          <div
                            key={`${preview.input}-page-${page.page}`}
                            className="page-card"
                            style={{ padding: "12px", display: "flex", flexDirection: "column", gap: "8px" }}
                          >
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%" }}>
                              <label style={{ display: "flex", alignItems: "center", gap: "6px", cursor: "pointer", fontWeight: "bold", fontSize: "0.92rem", color: "#1f1a16" }}>
                                <input
                                  type="checkbox"
                                  checked={isSelected}
                                  onChange={() => togglePageSelection(preview.input, page.page)}
                                  style={{ width: "16px", height: "16px", cursor: "pointer" }}
                                />
                                第 {page.page} 頁
                              </label>
                              <button
                                className="zoom-btn"
                                style={{ padding: "4px 10px", fontSize: "0.8rem", height: "auto", borderRadius: "8px", fontWeight: "bold" }}
                                onClick={() => setPreview({ title: fileName(preview.input), page })}
                              >
                                放大
                              </button>
                            </div>

                            <div
                              className="page-thumb"
                              style={{ cursor: "pointer", width: "100%" }}
                              onClick={() => togglePageSelection(preview.input, page.page)}
                            >
                              {page.thumbnail ? (
                                <img src={page.thumbnail} alt={`Page ${page.page} preview`} style={{ opacity: isSelected ? 1 : 0.4, transition: "opacity 140ms ease" }} />
                              ) : (
                                <div className="page-thumb-fallback">
                                  <strong>Page {page.page}</strong>
                                  <span>無縮圖資料</span>
                                </div>
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                ) : (
                  <p className="muted">正在載入頁面縮圖...</p>
                )}
              </div>
            ))}
          </div>
        </section>
      )}

      <section className="panel">
        <div className="section-head">
          <div>
            <h2>結果預覽</h2>
          </div>
        </div>

        {results.length ? (
          <div className="result-stack">
            {results.map((result) => (
              <article className={`result-card ${result.status === "success" ? "success" : "failed"}`} key={result.input}>
                <header className="result-header">
                  <div>
                    <p className="result-name" style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                      {fileName(result.input)}
                      {result.tables.some((t) => t.ratio) && (
                        <span className="badge success" style={{ fontSize: "0.8rem", padding: "2px 8px", background: "rgba(44, 114, 75, 0.12)", color: "#2c724b", border: "1px solid rgba(44, 114, 75, 0.2)" }}>
                          比例: {result.tables.find((t) => t.ratio)?.ratio}
                        </span>
                      )}
                    </p>
                    <p className="muted">{previewPath(result.input, 96)}</p>
                    {result.tables.some((t) => t.formula) && (
                      <p style={{ margin: "4px 0 0 0", fontSize: "0.85rem", color: "#2c724b", fontWeight: "bold" }}>
                        🧮 {result.tables.find((t) => t.formula)?.formula}
                      </p>
                    )}
                  </div>
                  <div className="status-row">
                    <span className={`badge ${result.status === "success" ? "success" : "failed"}`}>{result.status === "success" ? "成功" : "失敗"}</span>
                    <span className="badge neutral">{result.table_count} tables</span>
                  </div>
                </header>

                <p className="result-message">{result.message}</p>

                {result.output && (
                  <div className="output-pill" style={{ marginBottom: "16px" }}>
                    <span>Excel</span>
                    <code>{result.output}</code>
                  </div>
                )}

                {result.logs && result.logs.length > 0 && (
                  <div style={{ marginTop: "12px", marginBottom: "16px" }}>
                    <details>
                      <summary style={{ fontSize: "0.85rem", fontWeight: "bold", color: "#2c724b", cursor: "pointer", padding: "4px 8px", background: "rgba(44, 114, 75, 0.05)", borderRadius: "8px", width: "fit-content" }}>
                        📋 檢視解析日誌 (Parsing Logs)
                      </summary>
                      <div style={{
                        marginTop: "8px",
                        padding: "10px 14px",
                        background: "#1e1e1e",
                        color: "#d4d4d4",
                        fontFamily: "Consolas, Monaco, monospace",
                        fontSize: "0.88rem",
                        borderRadius: "8px",
                        maxHeight: "180px",
                        overflowY: "auto",
                        whiteSpace: "pre-wrap",
                        border: "1px solid rgba(0, 0, 0, 0.15)"
                      }}>
                        {result.logs.map((log, i) => (
                          <div key={i} style={{ borderBottom: "1px solid #2d2d2d", padding: "4px 0", color: log.includes("fallback") ? "#e5c07b" : log.includes("Ignored") ? "#e06c75" : "#abb2bf" }}>
                            {log}
                          </div>
                        ))}
                      </div>
                    </details>
                  </div>
                )}

                <section className="subpanel" style={{ borderRadius: "18px" }}>
                  <div className="subpanel-head">
                    <h3>抽出的表格</h3>
                    <span>{result.tables.length} tables</span>
                  </div>

                  {result.tables.length ? (
                    <div className="table-stack">
                      {result.tables.map((table) => (
                        <article className="table-preview" key={`${result.input}-${table.page}-${table.index}`}>
                          <div className="table-scroll">
                            <table>
                              <tbody>
                                {table.rows.map((row, rowIndex) => (
                                  <tr key={`${table.title}-${rowIndex}`}>
                                    {row.map((cell, cellIndex) => (
                                      <td key={`${table.title}-${rowIndex}-${cellIndex}`}>{cell || " "}</td>
                                    ))}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </article>
                      ))}
                    </div>
                  ) : (
                    <p className="muted">此 PDF 無格線表格或屬於純文字格式，已在 Excel 中輸出為 Text 分頁，未擷取到格線表格。</p>
                  )}
                </section>
              </article>
            ))}
          </div>
        ) : (
          <p className="muted">執行後會在這裡顯示每個 PDF 的輸出檔案與表格。</p>
        )}
      </section>

      <PreviewModal preview={preview} onClose={() => setPreview(null)} />
      <VersionModal isOpen={showVersionInfo} onClose={() => setShowVersionInfo(false)} />
      {/* Debug Logs Section */}
      <section className="panel" style={{ marginTop: "20px", fontSize: "0.85rem", opacity: 0.8 }}>
        <details>
          <summary style={{ padding: "8px 12px", background: "rgba(0, 0, 0, 0.04)", borderRadius: "8px", cursor: "pointer" }}>
            <strong>🔧 偵錯資訊 (Debug Logs)</strong>
            <span style={{ fontSize: "0.75rem", color: "#6d6254", marginLeft: "10px" }}>
              {isDragging ? "正在拖曳中 (Dragging)" : "無拖曳活動"}
            </span>
          </summary>
          <div style={{ padding: "12px", background: "#1e1e1e", color: "#67d87c", fontFamily: "monospace", borderRadius: "8px", marginTop: "10px", maxHeight: "150px", overflowY: "auto" }}>
            <p style={{ margin: "0 0 8px 0", color: "#aaa" }}>
              提示：若拖曳檔案時出現禁止符號，請確認命令提示字元 (CMD/PowerShell) 與檔案總管是否皆為一般使用者權限（非系統管理員身分）。
            </p>
            {debugLogs.length === 0 ? (
              <p style={{ margin: 0, color: "#888" }}>尚無事件記錄。請嘗試拖曳檔案進來...</p>
            ) : (
              debugLogs.map((log, i) => (
                <div key={i} style={{ borderBottom: "1px solid #333", padding: "4px 0" }}>{log}</div>
              ))
            )}
          </div>
        </details>
      </section>
    </main>
  );
}

function getDefaultProjectName(fileName: string): string {
  // Remove extension .pdf (case insensitive)
  let name = fileName.replace(/\.pdf$/i, "");
  
  // Remove "比赫-" or "比赫" prefix
  name = name.replace(/^比赫\s*[-_]?\s*/i, "");

  // Remove common date formats:
  // 1. YYYYMMDD (e.g., 20260317)
  name = name.replace(/\b\d{8}\b/g, "");
  // 2. YYYY-MM-DD or YYYY/MM/DD
  name = name.replace(/\b\d{4}[-/]\d{2}[-/]\d{2}\b/g, "");
  // 3. YYYY年MM月DD日
  name = name.replace(/\b\d{4}年\d{1,2}月\d{1,2}日\b/g, "");
  
  // Clean up: collapse multiple spaces, strip trailing/leading spaces and trailing dashes/underscores
  name = name.replace(/\s*[-_]+\s*$/, "")
             .replace(/\s+/g, " ")
             .trim();
  return name;
}
