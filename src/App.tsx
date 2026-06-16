import { invoke } from "@tauri-apps/api/core";
import { open } from "@tauri-apps/plugin-dialog";
import { useMemo, useState } from "react";

type ExtractedTable = {
  page: number;
  index: number;
  title: string;
  rows: string[][];
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
  const [busy, setBusy] = useState(false);
  const [results, setResults] = useState<ConvertResult[]>([]);
  const [pdfPreviews, setPdfPreviews] = useState<ConvertResult[]>([]);
  const [loadingPreviews, setLoadingPreviews] = useState(false);
  const [selectedPages, setSelectedPages] = useState<Record<string, number[]>>({});
  const [preview, setPreview] = useState<PreviewTarget | null>(null);
  const [globalError, setGlobalError] = useState<string | null>(null);

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

  async function choosePdfs() {
    const selected = await open({
      multiple: true,
      filters: [{ name: "PDF", extensions: ["pdf"] }],
    });
    if (!selected) return;
    const paths = Array.isArray(selected) ? selected : [selected];
    setPdfs(paths);
    setResults([]);
    setPreview(null);
    setPdfPreviews([]);
    setSelectedPages({});
    setGlobalError(null);

    setLoadingPreviews(true);
    try {
      const previews = await invoke<ConvertResult[]>("get_pdf_previews", {
        pdfPaths: paths,
      });
      setPdfPreviews(previews);

      // Auto-select all pages of loaded PDFs by default
      const initialSelections: Record<string, number[]> = {};
      previews.forEach((p) => {
        initialSelections[p.input] = p.pages.map((page) => page.page);
      });
      setSelectedPages(initialSelections);
    } catch (e) {
      console.error("Failed to load PDF previews:", e);
      setGlobalError(`載入 PDF 預覽失敗：${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setLoadingPreviews(false);
    }
  }

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
      });
      setResults(converted);
      setPreview(null);
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

      <section className="hero" style={{ marginTop: "1rem" }}>
        <p className="intro" style={{ marginTop: 0 }}>
          先把表格型 PDF 抽成 Excel，再把抽出的表格與 PDF 頁面直接顯示在畫面上，方便比對與 debug。
        </p>

        <div className="stats">
          <article>
            <span>選取 PDF</span>
            <strong>{pdfs.length}</strong>
          </article>
          <article>
            <span>成功轉換</span>
            <strong>{successCount}</strong>
          </article>
          <article>
            <span>抽出表格</span>
            <strong>{tableCount}</strong>
          </article>
          <article>
            <span>PDF 頁數</span>
            <strong>{pageCount}</strong>
          </article>
        </div>
      </section>

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

      <section className="panel command-panel">
        <div className="actions">
          <button onClick={choosePdfs}>選擇 PDF</button>
          <button onClick={chooseOutputDir}>選擇輸出資料夾</button>
          <button className="primary" onClick={convert} disabled={busy || !pdfs.length || !outputDir}>
            {busy ? "轉換中..." : "轉成 Excel"}
          </button>
        </div>

        <div className="grid">
          <div className="card">
            <h2>輸入 PDF</h2>
            {pdfs.length ? (
              <ul className="file-list">
                {pdfs.map((path) => (
                  <li key={path}>
                    <strong>{fileName(path)}</strong>
                    <span>{previewPath(path)}</span>
                  </li>
                ))}
              </ul>
            ) : (
              <p className="muted">尚未選擇 PDF。</p>
            )}
          </div>
          <div className="card">
            <h2>輸出位置</h2>
            <p className={outputDir ? "path" : "muted"}>{outputDir || "尚未選擇輸出資料夾。"}</p>
          </div>
        </div>
      </section>

      {/* PDF Page Previews (shown immediately after loading PDFs) */}
      {(loadingPreviews || pdfPreviews.length > 0) && (
        <section className="panel">
          <div className="section-head">
            <div>
              <p className="section-kicker">PDF Previews</p>
              <h2>PDF 頁面預覽</h2>
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
            <p className="section-kicker">結果預覽</p>
            <h2>轉換結果</h2>
          </div>
          <p className="muted">每個 PDF 會顯示 Excel 路徑與抽出的表格。</p>
        </div>

        {results.length ? (
          <div className="result-stack">
            {results.map((result) => (
              <article className={`result-card ${result.status}`} key={result.input}>
                <header className="result-header">
                  <div>
                    <p className="result-name">{fileName(result.input)}</p>
                    <p className="muted">{previewPath(result.input, 96)}</p>
                  </div>
                  <div className="status-row">
                    <span className={`badge ${result.status}`}>{result.status === "success" ? "成功" : "失敗"}</span>
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

      {preview && (
        <div className="preview-overlay" role="button" tabIndex={0} onClick={() => setPreview(null)}>
          <div className="preview-modal" onClick={(event) => event.stopPropagation()} style={{ maxWidth: "960px" }}>
            <header className="preview-head">
              <div>
                <p className="section-kicker">Page Preview</p>
                <h2>{preview.title} - 第 {preview.page.page} 頁</h2>
              </div>
              <button onClick={() => setPreview(null)}>關閉</button>
            </header>

            <div className="preview-image">
              {preview.page.thumbnail && <img src={preview.page.thumbnail} alt={`Page ${preview.page.page} full preview`} />}
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
