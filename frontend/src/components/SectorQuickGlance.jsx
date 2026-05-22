import { useEffect, useMemo, useState } from "react";
import { generateSectorQuickGlance } from "../api";

function formatValue(value) {
  if (value === null || value === undefined || value === "") {
    return "--";
  }
  if (typeof value === "number") {
    return Number.isInteger(value) ? String(value) : value.toFixed(2);
  }
  return value;
}

function formatScore(value) {
  if (value === null || value === undefined || value === "") {
    return "--";
  }
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return value;
  }
  return Number.isInteger(number) ? String(number) : number.toFixed(2);
}

function normalizeKey(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/^nifty\s+/i, "")
    .replace(/[^a-z0-9]/g, "");
}

function findIndexScore(sectorName, indexRpRows) {
  const sectorKey = normalizeKey(sectorName);
  if (!sectorKey) {
    return null;
  }

  const exactMatch = indexRpRows.find((row) => normalizeKey(row.index_name) === sectorKey);
  if (exactMatch) {
    return exactMatch.total;
  }

  const partialMatch = indexRpRows.find((row) => {
    const indexKey = normalizeKey(row.index_name);
    return indexKey.includes(sectorKey) || sectorKey.includes(indexKey);
  });
  return partialMatch?.total ?? null;
}

function KpiCards({ summary }) {
  const items = [
    ["Total Sectors", summary?.total_sectors ?? 0],
    ["Total Stocks", summary?.total_stocks ?? 0],
    ["Best Performing Sector", summary?.best_sector ?? "--"],
    ["Best Stock Overall", summary?.best_stock_overall ?? "--"],
    ["Weakest Stock Overall", summary?.weakest_stock_overall ?? "--"],
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

function StockList({ title, rows, tone }) {
  return (
    <section className={`quick-glance-stock-section ${tone}`}>
      <div className="quick-glance-subtitle">{title}</div>
      <div className="quick-glance-stock-list">
        {rows.map((row) => (
          <div className="quick-glance-stock-row" key={`${title}-${row.rank}-${row.company_name}`}>
            <div className="quick-glance-stock-main">
              <span className="quick-glance-rank">{row.rank}</span>
              <span className="quick-glance-company" title={row.company_name || ""}>
                {formatValue(row.company_name)}
              </span>
            </div>
            <div className="quick-glance-metrics">
              <span>FS: {formatValue(row.fs)}</span>
              <span>TS: {formatValue(row.ts)}</span>
              <span className="quick-glance-total">Total: {formatValue(row.total)}</span>
            </div>
          </div>
        ))}
      </div>
    </section>
  );
}

function SectorCard({ sector }) {
  return (
    <section className="panel quick-glance-sector-card">
      <div className="quick-glance-card-header">
        <div>
          <h2>{sector.sector}</h2>
          <p>{sector.index_score !== null ? "Index Total" : "Avg Total"}: {formatScore(sector.display_score)}</p>
        </div>
        <span className="quick-glance-score-badge">{formatScore(sector.display_score)}</span>
      </div>
      <div className="quick-glance-stack">
        <StockList title="Top 3 Stocks" rows={sector.top_stocks} tone="positive" />
        <StockList title="Bottom 3 Stocks" rows={sector.bottom_stocks} tone="negative" />
      </div>
    </section>
  );
}

export default function SectorQuickGlance({
  file,
  onFileChange,
  initialFile = null,
  response,
  onResponseChange,
  indexRpRows = [],
  search,
  onSearchChange,
  sortMode,
  onSortModeChange,
}) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    if (initialFile && !file) {
      onFileChange(initialFile);
    }
  }, [initialFile, file, onFileChange]);

  const sectors = useMemo(() => {
    const query = search.trim().toLowerCase();
    let filtered = response.sectors.map((sector) => {
      const indexScore = findIndexScore(sector.sector, indexRpRows);
      const fallbackScore = sector.sort_score ?? sector.sector_score ?? 0;
      return {
        ...sector,
        index_score: indexScore,
        display_score: indexScore ?? fallbackScore,
        performance_score: indexScore ?? fallbackScore,
      };
    }).filter((sector) => {
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
      if (sortMode === "weakest") {
        return (a.performance_score ?? Infinity) - (b.performance_score ?? Infinity);
      }
      if (sortMode === "sector") {
        return a.sector.localeCompare(b.sector);
      }
      return (b.performance_score ?? -Infinity) - (a.performance_score ?? -Infinity);
    });

    return filtered;
  }, [response.sectors, indexRpRows, search, sortMode]);

  async function handleGenerate() {
    if (!file) {
      setError("Please upload the Daily Stock Dashboard Sheet.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const data = await generateSectorQuickGlance({ stockFile: file });
      onResponseChange(data);
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
          <p>Best-performing sectors first, with Top 3 and Bottom 3 stocks in each sector.</p>
        </div>

        <div className="upload-grid single-upload-grid">
          <label className="upload-box">
            <span className="upload-label">Upload Daily Stock Dashboard Sheet</span>
            <input
              type="file"
              accept=".xlsx,.xlsm,.xls"
              onChange={(event) => {
                onFileChange(event.target.files?.[0] || null);
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
          <KpiCards summary={response.summary || response.kpis} />

          <section className="panel quick-glance-controls">
            <input
              className="quick-glance-search"
              type="search"
              value={search}
              onChange={(event) => onSearchChange(event.target.value)}
              placeholder="Search sector or stock"
            />
            <select
              className="index-select quick-glance-sort"
              value={sortMode}
              onChange={(event) => onSortModeChange(event.target.value)}
            >
              <option value="performance">Best Performing Sectors</option>
              <option value="sector">Sector Name</option>
              <option value="weakest">Weakest Sectors</option>
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
