import { useState } from "react";
import { processIndexRp } from "../api";

const emptyResponse = {
  kpis: null,
  rows: [],
};

function formatValue(value) {
  if (value === null || value === undefined || value === "") {
    return "--";
  }
  if (typeof value === "number") {
    return Number.isInteger(value) ? String(value) : value.toFixed(2);
  }
  return value;
}

function getCallClass(value) {
  const text = String(value || "").toUpperCase();
  if (text === "BUY") {
    return "change-positive";
  }
  if (text === "SELL") {
    return "change-negative";
  }
  return "change-neutral";
}

function IndexRpKpis({ kpis }) {
  const items = [
    ["Total Indexes", kpis?.total_indexes ?? 0],
    ["Daily BUY Calls", kpis?.daily_buy_calls ?? 0],
    ["Weekly BUY Calls", kpis?.weekly_buy_calls ?? 0],
    ["Above 30WMA", kpis?.above_30wma ?? 0],
    ["Strongest Index", kpis?.strongest_index ?? "--"],
    ["Weakest Index", kpis?.weakest_index ?? "--"],
  ];

  return (
    <section className="kpi-grid">
      {items.map(([label, value]) => (
        <div className="kpi-card" key={label}>
          <div className="kpi-label">{label}</div>
          <div className="kpi-value index-rp-kpi-value">{value}</div>
        </div>
      ))}
    </section>
  );
}

function IndexRpTable({ rows }) {
  return (
    <section className="panel table-panel">
      <div className="panel-header compact">
        <h2>Index RP Ranking</h2>
      </div>
      <div className="table-wrapper">
        <table className="ranking-table">
          <thead>
            <tr>
              <th>Index Name</th>
              <th>LTP</th>
              <th>Daily RP Call</th>
              <th>Weekly RP Call</th>
              <th>RSI</th>
              <th>30WMA</th>
              <th>RSI Score</th>
              <th>30WMA Score</th>
              <th>Weekly RP Score</th>
              <th>Daily RP Score</th>
              <th>Total</th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan="11" className="empty-cell">
                  No rows available.
                </td>
              </tr>
            ) : (
              rows.map((row) => (
                <tr key={`${row.index_name}-${row.total}`}>
                  <td className="company-cell" title={row.index_name || ""}>
                    <span className="company-text">{formatValue(row.index_name)}</span>
                  </td>
                  <td>{formatValue(row.ltp)}</td>
                  <td>
                    <span className={`change-pill ${getCallClass(row.daily_rp_call)}`}>
                      {formatValue(row.daily_rp_call)}
                    </span>
                  </td>
                  <td>
                    <span className={`change-pill ${getCallClass(row.weekly_rp_call)}`}>
                      {formatValue(row.weekly_rp_call)}
                    </span>
                  </td>
                  <td>{formatValue(row.rsi)}</td>
                  <td>{formatValue(row.wma_30)}</td>
                  <td>{formatValue(row.rsi_score)}</td>
                  <td>{formatValue(row.wma_30_score)}</td>
                  <td>{formatValue(row.weekly_rp_score)}</td>
                  <td>{formatValue(row.daily_rp_score)}</td>
                  <td>{formatValue(row.total)}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </section>
  );
}

export default function IndexRpDashboard() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [response, setResponse] = useState(emptyResponse);

  async function handleProcess() {
    if (!file) {
      setError("Please upload an Index RP Excel file.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const data = await processIndexRp({ indexRpFile: file });
      setResponse(data);
    } catch (requestError) {
      const detail =
        requestError?.response?.data?.detail ||
        requestError?.message ||
        "Index RP processing failed.";
      setError(detail);
    } finally {
      setLoading(false);
    }
  }

  return (
    <>
      <section className="panel upload-panel">
        <div className="panel-header">
          <h2>Index RP Upload</h2>
          <p>Upload the Index RP Excel file. The backend reads Sheet2 and ranks indexes by Total.</p>
        </div>

        <div className="upload-grid single-upload-grid">
          <label className="upload-box">
            <span className="upload-label">Upload Index RP Excel File</span>
            <input
              type="file"
              accept=".xlsx,.xlsm,.xls"
              onChange={(event) => {
                setFile(event.target.files?.[0] || null);
                setError("");
              }}
            />
            <span className="upload-file">{file?.name || "No file selected"}</span>
          </label>
        </div>

        <button
          className="compare-button"
          type="button"
          onClick={handleProcess}
          disabled={loading || !file}
        >
          {loading ? "Processing..." : "Process Index RP"}
        </button>
      </section>

      {error ? <div className="error-banner">{error}</div> : null}

      {response.rows.length > 0 ? (
        <>
          <IndexRpKpis kpis={response.kpis} />
          <IndexRpTable rows={response.rows} />
        </>
      ) : (
        <section className="empty-state panel">
          <h2>Index RP Dashboard Waiting For File</h2>
          <p>Upload one Index RP Excel file and click <strong>Process Index RP</strong>.</p>
        </section>
      )}
    </>
  );
}
