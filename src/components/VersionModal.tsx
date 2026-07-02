type VersionModalProps = {
  isOpen: boolean;
  onClose: () => void;
};

export default function VersionModal({ isOpen, onClose }: VersionModalProps) {
  if (!isOpen) return null;

  return (
    <div className="preview-overlay" role="button" tabIndex={0} onClick={onClose}>
      <div className="preview-modal" onClick={(event) => event.stopPropagation()} style={{ maxWidth: "600px" }}>
        <header className="preview-head">
          <div>
            <p className="section-kicker">Version History</p>
            <h2>版本更新日誌</h2>
          </div>
          <button onClick={onClose}>關閉</button>
        </header>

        <div style={{ padding: "8px 0", maxHeight: "400px", overflowY: "auto" }}>
          <div style={{ marginBottom: "20px" }}>
            <h3 style={{ margin: "0 0 8px 0", fontSize: "1.1rem", color: "#1d3b2a" }}>v0.6.0 <span style={{ fontSize: "0.85rem", color: "#6d6254", fontWeight: "normal" }}> (2026-07-02)</span></h3>
            <ul style={{ paddingLeft: "20px", margin: 0, lineHeight: "1.6" }}>
              <li><strong>專案匯總管理與自訂名稱功能</strong>
                <ul>
                  <li><strong>專案名稱自訂與對齊</strong>：UI 新增「專案名稱」輸入框。設定後會將產出 Excel 的工作表（Sheet）重新命名為專案名稱，並在 Excel 第一列（Row 1）插入 24 級 Calibri 粗體合併大標題，其餘表格自動下移 1 行，維持完美對齊。</li>
                  <li><strong>匯總管理表自動整合</strong>：每次轉檔時，除了產出獨立 Excel，還會自動將該專案工作表複製合併至 <code>專案匯總管理表.xlsx</code> 中，以便集中檢視與管理所有專案。</li>
                  <li><strong>重複名稱阻擋與保護</strong>：程式啟動或變更目錄時，會自動掃描 <code>專案匯總管理表.xlsx</code> 中的所有工作表。當輸入重複名稱時，UI 會即時警示並禁用轉換按鈕，防止覆蓋歷史資料。</li>
                  <li><strong>修正 UI 版本號顯示</strong>：修復先前主介面上版本顯示徽章未同步更新的問題，現已統一顯示為對應的 v0.6.0。</li>
                </ul>
              </li>
            </ul>
          </div>

          <div style={{ marginBottom: "20px" }}>
            <h3 style={{ margin: "0 0 8px 0", fontSize: "1.1rem", color: "#6d6254" }}>v0.5.0 <span style={{ fontSize: "0.85rem", color: "#6d6254", fontWeight: "normal" }}> (2026-07-01)</span></h3>
            <ul style={{ paddingLeft: "20px", margin: 0, lineHeight: "1.6" }}>
              <li><strong>比赫 PDF 報價單「版型二」排版與外框線優化</strong>
                <ul>
                  <li><strong>數量與單位分拆</strong>：Qty 與單位拆為兩欄（數量欄與單位欄分開），與舊版欄位結構完全一致。新版表頭合併 Column D 與 E，置中顯示 "Qty"；數據列的數量填入數字，單位置中顯示。</li>
                  <li><strong>自適應起點與分隔列</strong>：主表價格與試算表公式一律對齊 Column H 欄。分隔列寬度與計算表起點依備註欄是否存在，動態平移（無備註在 I / J，有備註在 J / K），完美對齊。</li>
                  <li><strong>區塊大外框與分隔線</strong>：上方區塊（公司 Logo、Quotation 表頭、客資區）套用完整的大外框，並於 Quotation 表頭下方繪製分隔線；備註區（Remarks）套用動態寬度大框，移除多餘的空白列。</li>
                  <li><strong>細部解析優化</strong>：調換項目內容比對優先級以正確提取項目內容；新增 Mail 解析提取與 dynamic 客資標籤。</li>
                </ul>
              </li>
            </ul>
          </div>

          <div style={{ marginBottom: "20px" }}>
            <h3 style={{ margin: "0 0 8px 0", fontSize: "1.1rem", color: "#6d6254" }}>v0.4.0 <span style={{ fontSize: "0.85rem", color: "#6d6254", fontWeight: "normal" }}> (2026-06-29)</span></h3>
            <ul style={{ paddingLeft: "20px", margin: 0, lineHeight: "1.6" }}>
              <li><strong>比赫 PDF 報價單「版型二」儲存格樣式與線框優化</strong>
                <ul>
                  <li>取消客戶資訊區內部實線，改為單一乾淨外框包覆。</li>
                  <li>支援內容欄與備註欄多行排版自動折行，完美還原 PDF 換行結構。</li>
                  <li>優化保固列檢測，支援備份搜尋內容欄（Description）以應對專案欄為空的情形。</li>
                  <li>當原始檔案無專案名稱時自動保持空白，避免產生多餘的 "專案 :" 標籤。</li>
                </ul>
              </li>
              <li><strong>介面與結果預覽優化</strong>
                <ul>
                  <li>結果預覽卡片新增「計算比例徽章」與「保固計算公式」明細展示。</li>
                </ul>
              </li>
            </ul>
          </div>

          <div style={{ marginBottom: "20px" }}>
            <h3 style={{ margin: "0 0 8px 0", fontSize: "1.1rem", color: "#6d6254" }}>v0.3.0 <span style={{ fontSize: "0.85rem", color: "#6d6254", fontWeight: "normal" }}> (2026-06-29)</span></h3>
            <ul style={{ paddingLeft: "20px", margin: 0, lineHeight: "1.6" }}>
              <li><strong>比赫 PDF 報價單「版型一」圖檔嵌入排版優化 (Option 1)</strong>
                <ul>
                  <li>將 PDF 頂部的公司/客戶資訊與底部的備註條款裁剪為 PNG 圖檔並嵌入 Excel。</li>
                  <li>完美還原 Logo 與備註排版，同時保持中間表格為可編輯、可計算與黑塗效果。</li>
                </ul>
              </li>
              <li><strong>總計列 (TTL) 匹配精度提升</strong>
                <ul>
                  <li>使用單字邊界 <code>\b(ttl|total)\b</code> 判定總計列，避免備註中的 <code>settlement</code> 單字被誤匹配。</li>
                </ul>
              </li>
            </ul>
          </div>

          <div style={{ marginBottom: "20px" }}>
            <h3 style={{ margin: "0 0 8px 0", fontSize: "1.1rem", color: "#6d6254" }}>v0.2.0 <span style={{ fontSize: "0.85rem", color: "#6d6254", fontWeight: "normal" }}> (2026-06-29)</span></h3>
            <ul style={{ paddingLeft: "20px", margin: 0, lineHeight: "1.6" }}>
              <li><strong>比赫 PDF 報價單自動偵測與轉換優化</strong>
                <ul>
                  <li>自動合併跨列標題，正確抓取「3年保固 單價(USD)/台」。</li>
                  <li>調整欄位順序：將 <code>Qty</code> (及其單位) 置於價格欄位之前。</li>
                  <li>新增保固年期自動辨識，並依比例試算保固金額與重算 TTL 總額。</li>
                  <li>Excel 報價輸出自動塗黑：將資料列之單價、保固單價、總計與備註欄位套用黑色底色與黑色字型，實現遮罩效果。</li>
                </ul>
              </li>
              <li><strong>介面與體驗優化</strong>
                <ul>
                  <li>點擊版本號可開啟此更新歷史視窗。</li>
                </ul>
              </li>
            </ul>
          </div>

          <div style={{ borderTop: "1px solid rgba(0, 0, 0, 0.08)", paddingTop: "12px" }}>
            <h3 style={{ margin: "0 0 8px 0", fontSize: "1.1rem", color: "#6d6254" }}>v0.1.0 <span style={{ fontSize: "0.85rem", color: "#888", fontWeight: "normal" }}> (初始版本)</span></h3>
            <ul style={{ paddingLeft: "20px", margin: 0, lineHeight: "1.6", color: "#666" }}>
              <li>支援多格式 PDF 轉 Excel 報價單轉換器。</li>
              <li>支援多檔案拖曳載入與單/多頁預覽功能。</li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
}
