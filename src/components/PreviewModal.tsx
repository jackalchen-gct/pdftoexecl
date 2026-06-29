type ExtractedPage = {
  page: number;
  text: string;
  thumbnail: string;
};

type PreviewTarget = {
  title: string;
  page: ExtractedPage;
};

type PreviewModalProps = {
  preview: PreviewTarget | null;
  onClose: () => void;
};

export default function PreviewModal({ preview, onClose }: PreviewModalProps) {
  if (!preview) return null;

  return (
    <div className="preview-overlay" role="button" tabIndex={0} onClick={onClose}>
      <div className="preview-modal" onClick={(event) => event.stopPropagation()} style={{ maxWidth: "960px" }}>
        <header className="preview-head">
          <div>
            <p className="section-kicker">Page Preview</p>
            <h2>{preview.title} - 第 {preview.page.page} 頁</h2>
          </div>
          <button onClick={onClose}>關閉</button>
        </header>

        <div className="preview-image">
          {preview.page.thumbnail && <img src={preview.page.thumbnail} alt={`Page ${preview.page.page} full preview`} />}
        </div>
      </div>
    </div>
  );
}
