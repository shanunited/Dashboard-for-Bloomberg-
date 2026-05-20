import { useEffect, useMemo, useState } from "react";
import { generateSectorQuickGlance } from "../api";

const emptyResponse = {
  kpis: null,
  sectors: [],
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

function KpiCards({ kpis }) {
  const items = [
    ["Total Sectors", kpis?.total_sectors ?? 0],
    ["Total Stocks", kpis?.total_stocks ?? 0],
    ["Best Stock Overall", kpis?.best_stock_overall ?? "--"],
    ["Weakest Stock Overall", kpis?.weakest_stock_overall ?? "--"],
  ];

  return (
    <section className="kpi-grid">
      {items.map(([label, value]) => (
        <div className="kpi-card" key={label}>
          <div className="kpi-label">{label}</div>
          <div className="kpi-value quick-glance-kpi-value">{value}</div>
        </div>
      ))}
    </section>
  );
}

function MiniStockTable({ title, rows }) {
  return (
    <section className="quick-glance-table-wrap">
      <div className="quick-glance-subtitle">{title}</div>
      <div className="table-wrapper quick-glance-table-scroll">
        <table className="ranking-table quick-glance-table">
          <thead>
            <tr>
              <th>Rank</th>
              <th>Company</th>
              <th>FS</th>
              <th>TS</th>
              <th>Total</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={`${title}-${row.rank}-${row.company_name}`}>
                <td>{formatValue(row.rank)}</td>
                <td className="company-cell" title={row.company_name || ""}>
                  <span className="company-text">{formatValue(row.company_name)}</span>
                </td>
                <td>{formatValue(row.fs)}</td>
                <td>{formatValue(row.ts)}</td>
                <td>{formatValue(row.total)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function SectorCard({ sector }) {
  return (
    <section className="panel quick-glance-sector-card">
      <div className="panel-header compact">
        <h2>{sector.sector}</h2>
      </div>
      <div className="quick-glance-pair">
        <MiniStockTable title="Top 3 Stocks" rows={sector.top_stocks} />
        <MiniStockTable title="Bottom 3 Stocks" rows={sector.bottom_stocks} />
      </div>
    </section>
  );
}

export default function SectorQuickGlance({ initialFile = null }) {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [response, setResponse] = useState(emptyResponse);
  const [search, setSearch] = useState("");
  const [sortMode, setSortMode] = useState("sector");

  useEffect(() => {
    if (initialFile && !file) {
      setFile(initialFile);
    }
  }, [initialFile, file]);

  const sectors = useMemo(() => {
    const query = search.trim().toLowerCase();
    let filtered = response.sectors.filter((sector) => {
      if (!query) {
        return true;
      }
      const stockNames = [...sector.top_stocks, ...sector.bottom_stocks]
        .map((stock) => stock.company_name || "")
        .join(" ")
        .toLowerCase();
      return sector.sector.toLowerCase().includes(query) || stockNames.includes(query);
    });

    filtered = [...filtered].sort((a, b) => {
      if (sortMode === "highest") {
        return (b.top_stocks[0]?.total ?? -Infinity) - (a.top_stocks[0]?.total ?? -Infinity);
      }
      if (sortMode === "lowest") {
        return (a.bottom_stocks[0]?.total ?? Infinity) - (b.bottom_stocks[0]?.total ?? Infinity);
      }
      return a.sector.localeCompare(b.sector);
    });

    return filtered;
  }, [response.sectors, search, sortMode]);

  async function handleGenerate() {
    if (!file) {
      setError("Please upload the Daily Stock Dashboard Sheet.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const data = await generateSectorQuickGlance({ stockFile: file });
      setResponse(data);
    } catch (requestError) {
      const detail =
        requestError?.response?.data?.detail ||
        requestError?.message ||
        "Sector quick glance processing failed.";
      setError(detail);
    } finally {
      setLoading(false);
    }
  }

  return (
    <>
      <section className="panel upload-panel">
        <div className="panel-header">
          <h2>Sector Quick Glance</h2>
          <p>View Top 3 and Bottom 3 stocks across all sectors without selecting each index manually.</p>
        </div>

        <div className="upload-grid single-upload-grid">
          <label className="upload-box">
            <span className="upload-label">Upload Daily Stock Dashboard Sheet</span>
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
          onClick={handleGenerate}
          disabled={loading || !file}
        >
          {loading ? "Generating..." : "Generate Quick Glance"}
        </button>
      </section>

      {error ? <div className="error-banner">{error}</div> : null}

      {response.sectors.length > 0 ? (
        <>
          <KpiCards kpis={response.kpis} />

          <section className="panel quick-glance-controls">
            <input
              className="quick-glance-search"
              type="search"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Search sector or stock"
            />
            <select
              className="index-select quick-glance-sort"
              value={sortMode}
              onChange={(event) => setSortMode(event.target.value)}
            >
              <option value="sector">Sector Name</option>
              <option value="highest">Highest Top Stock Total</option>
              <option value="lowest">Lowest Bottom Stock Total</option>
            </select>
          </section>

          <div className="quick-glance-grid">
            {sectors.map((sector) => (
              <SectorCard sector={sector} key={sector.sector} />
            ))}
          </div>
        </>
      ) : (
        <section className="empty-state panel">
          <h2>Sector Quick Glance Waiting For File</h2>
          <p>Upload the Daily Stock Dashboard Sheet and click <strong>Generate Quick Glance</strong>.</p>
        </section>
      )}
    </>
  );
}
