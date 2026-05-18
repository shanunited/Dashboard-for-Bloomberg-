import { useState } from "react";
import { generateReport } from "../api";

function todayString() {
  const d = new Date();
  return `${d.getDate()}-${d.getMonth() + 1}-${d.getFullYear()}`;
}

export default function ReportGenerator() {
  const [file, setFile] = useState(null);
  const [date, setDate] = useState(todayString());
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState(false);

  async function handleGenerate() {
    if (!file) {
      setError("Please select an Excel file.");
      return;
    }
    if (!date.trim()) {
      setError("Please enter a report date.");
      return;
    }

    setLoading(true);
    setError("");
    setSuccess(false);

    try {
      const blob = await generateReport({ reportFile: file, reportDate: date.trim() });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `Stock_Report_${date.trim()}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setSuccess(true);
    } catch (err) {
      const detail =
        err?.response?.data
          ? await err.response.data.text().then((t) => {
              try { return JSON.parse(t).detail; } catch { return t; }
            })
          : err?.message || "Report generation failed.";
      setError(detail);
    } finally {
      setLoading(false);
    }
  }

  return (
    <section className="panel report-generator-panel">
      <div className="panel-header">
        <h2>Daily Stock Scoring Report</h2>
        <p>
          Upload today&apos;s Excel file and generate a professional PDF report with Top 3 &amp;
          Bottom 3 stocks per index, scorecards, and key metrics.
        </p>
      </div>

      <div className="rg-body">
        <div className="rg-row">
          <label className="upload-box rg-upload">
            <span className="upload-label">Excel File</span>
            <input
              type="file"
              accept=".xlsx,.xlsm,.xls"
              onChange={(e) => {
                setFile(e.target.files?.[0] || null);
                setSuccess(false);
                setError("");
              }}
            />
            <span className="upload-file">{file?.name || "No file selected"}</span>
          </label>

          <div className="rg-date-wrap">
            <span className="upload-label">Report Date</span>
            <input
              className="rg-date-input"
              type="text"
              value={date}
              onChange={(e) => { setDate(e.target.value); setSuccess(false); setError(""); }}
              placeholder="e.g. 12-5-2026"
            />
            <span className="rg-date-hint">Used in filename and footer (DD-M-YYYY)</span>
          </div>
        </div>

        <div className="rg-actions">
          <button
            className="compare-button rg-button"
            type="button"
            onClick={handleGenerate}
            disabled={loading || !file}
          >
            {loading ? "Generating PDF…" : "Generate PDF Report"}
          </button>

          {success && (
            <span className="rg-success">
              ✓ PDF downloaded successfully
            </span>
          )}
        </div>

        {error && <div className="error-banner rg-error">{error}</div>}
      </div>
    </section>
  );
}
