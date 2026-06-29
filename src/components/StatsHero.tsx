import { useMemo } from "react";

type StatsHeroProps = {
  pdfsCount: number;
  successCount: number;
  tableCount: number;
  pageCount: number;
};

export default function StatsHero({
  pdfsCount,
  successCount,
  tableCount,
  pageCount,
}: StatsHeroProps) {
  return (
    <section className="hero" style={{ marginTop: "0.5rem" }}>
      <div className="stats">
        <article>
          <span>選取 PDF</span>
          <strong>{pdfsCount}</strong>
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
  );
}
