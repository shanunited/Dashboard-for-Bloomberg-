import { useState } from "react";
import { generateCombinedReport } from "../api";

function filenameDate(fileName) {
  const match = fileName?.match(/(\d{1,2}[-_]\d{1,2}[-_]\d{2,4})/);
  if (match) {
    return match[1].replaceAll("_", "-");
  }
  return new Date().toLocaleDateString("en-GB").replaceAll("/", "-");
}

async function responseErrorDetail(err, fallback) {
  if (!err?.response?.data) {
    return err?.message || fallback;
  }
  const text = await err.response.data.text();
  try {
    return JSON.parse(text).detail || fallback;
  } catch {
    return text || fallback;
  }
}

export default function ReportGenerator() {
  const [stockFile, setStockFile] = useState(null);
  const [indexFile, setIndexFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState(false);

  async function handleGenerate() {
    if (!stockFile) {
      setError("Please upload the Daily Stock Dashboard Sheet.");
      return;
    }
    if (!indexFile) {
      setError("Please upload the Daily Index Sheet.");
      return;
    }

    setLoading(true);
    setError("");
    setSuccess(false);

    try {
      const blob = await generateCombinedReport({ stockFile, indexFile });
      const reportDate = filenameDate(indexFile.name);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `Combined_Index_Stock_Report_${reportDate}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setSuccess(true);
    } catch (err) {
      setError(await responseErrorDetail(err, "Combined report generation failed."));
    } finally {
      setLoading(false);
    }
  }

  return (
    <section className="panel report-generator-panel">
      <div className="panel-header">
        <h2>Daily Combined Report Generator</h2>
        <p>
          Upload the Daily Stock Dashboard Sheet and Daily Index Sheet to generate one combined PDF report.
        </p>
      </div>

      <div className="rg-body">
        <div className="rg-row">
          <label className="upload-box rg-upload">
            <span className="upload-label">Upload Daily Stock Dashboard Sheet</span>
            <input
              type="file"
              accept=".xlsx,.xlsm,.xls"
              onChange={(e) => {
                setStockFile(e.target.files?.[0] || null);
                setSuccess(false);
                setError("");
              }}
            />
            <span className="upload-file">{stockFile?.name || "No file selected"}</span>
          </label>

          <label className="upload-box rg-upload">
            <span className="upload-label">Upload Daily Index Sheet</span>
            <input
              type="file"
              accept=".xlsx,.xlsm,.xls"
              onChange={(e) => {
                setIndexFile(e.target.files?.[0] || null);
                setSuccess(false);
                setError("");
              }}
            />
            <span className="upload-file">{indexFile?.name || "No file selected"}</span>
          </label>
        </div>

        <div className="rg-actions">
          <button
            className="compare-button rg-button"
            type="button"
            onClick={handleGenerate}
            disabled={loading || !stockFile || !indexFile}
          >
            {loading ? "Generating PDF..." : "Generate Combined Report"}
          </button>

          {success && (
            <span className="rg-success">
              PDF downloaded successfully
            </span>
          )}
        </div>

        {error && <div className="error-banner rg-error">{error}</div>}
      </div>
    </section>
  );
}
