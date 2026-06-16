import { invoke } from "@tauri-apps/api/core";
import { open } from "@tauri-apps/plugin-dialog";
import { useState } from "react";

type ConvertResult = {
  input: string;
  output: string;
  status: "success" | "failed";
  message: string;
};

function fileName(path: string) {
  return path.split(/[\\/]/).pop() ?? path;
}

export default function App() {
  const [pdfs, setPdfs] = useState<string[]>([]);
  const [outputDir, setOutputDir] = useState<string>("");
  const [busy, setBusy] = useState(false);
  const [results, setResults] = useState<ConvertResult[]>([]);

  async function choosePdfs() {
    const selected = await open({
      multiple: true,
      filters: [{ name: "PDF", extensions: ["pdf"] }],
    });
    if (!selected) return;
    setPdfs(Array.isArray(selected) ? selected : [selected]);
    setResults([]);
  }

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
    try {
      const converted = await invoke<ConvertResult[]>("convert_pdfs", {
        pdfPaths: pdfs,
        outputDir,
      });
      setResults(converted);
    } finally {
      setBusy(false);
    }
  }

  return (
    <main className="shell">
      <section className="hero">
        <p className="eyebrow">Quotation PDF Converter</p>
        <h1>PDF 報價單轉 Excel</h1>
        <p className="intro">
          MVP 先處理 Nvidia 與比赫這類表格型 PDF；AMD 文字型報價會在下一階段加入專屬規則。
        </p>
      </section>

      <section className="panel">
        <div className="actions">
          <button onClick={choosePdfs}>選擇 PDF</button>
          <button onClick={chooseOutputDir}>選擇輸出資料夾</button>
          <button className="primary" onClick={convert} disabled={busy || !pdfs.length || !outputDir}>
            {busy ? "轉換中..." : "轉成 Excel"}
          </button>
        </div>

        <div className="grid">
          <div>
            <h2>輸入 PDF</h2>
            {pdfs.length ? (
              <ul>
                {pdfs.map((path) => (
                  <li key={path}>{fileName(path)}</li>
                ))}
              </ul>
            ) : (
              <p className="muted">尚未選擇 PDF。</p>
            )}
          </div>
          <div>
            <h2>輸出位置</h2>
            <p className={outputDir ? "" : "muted"}>{outputDir || "尚未選擇輸出資料夾。"}</p>
          </div>
        </div>
      </section>

      <section className="panel">
        <h2>轉換結果</h2>
        {results.length ? (
          <ul className="results">
            {results.map((result) => (
              <li className={result.status} key={result.input}>
                <strong>{fileName(result.input)}</strong>
                <span>{result.message}</span>
                {result.output && <code>{result.output}</code>}
              </li>
            ))}
          </ul>
        ) : (
          <p className="muted">執行後會在這裡顯示每個 PDF 的輸出檔案。</p>
        )}
      </section>
    </main>
  );
}

