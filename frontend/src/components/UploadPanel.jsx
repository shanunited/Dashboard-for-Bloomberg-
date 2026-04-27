export default function UploadPanel({
  yesterdayFile,
  todayFile,
  loading,
  onYesterdayChange,
  onTodayChange,
  onCompare,
}) {
  return (
    <section className="panel upload-panel">
      <div className="panel-header">
        <h2>File Upload</h2>
        <p>Upload yesterday and today Excel files, then run comparison.</p>
      </div>
      <div className="upload-grid">
        <label className="upload-box">
          <span className="upload-label">Yesterday Excel File</span>
          <input
            type="file"
            accept=".xlsx,.xlsm,.xls"
            onChange={(event) => onYesterdayChange(event.target.files?.[0] || null)}
          />
          <span className="upload-file">{yesterdayFile?.name || "No file selected"}</span>
        </label>

        <label className="upload-box">
          <span className="upload-label">Today Excel File</span>
          <input
            type="file"
            accept=".xlsx,.xlsm,.xls"
            onChange={(event) => onTodayChange(event.target.files?.[0] || null)}
          />
          <span className="upload-file">{todayFile?.name || "No file selected"}</span>
        </label>
      </div>

      <button
        className="compare-button"
        type="button"
        onClick={onCompare}
        disabled={loading || !yesterdayFile || !todayFile}
      >
        {loading ? "Comparing..." : "Compare Files"}
      </button>
    </section>
  );
}
