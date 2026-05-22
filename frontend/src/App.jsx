import { useEffect, useState } from "react";
import { compareFiles } from "./api";
import UploadPanel from "./components/UploadPanel";
import KPICards from "./components/KPICards";
import IndexSelector from "./components/IndexSelector";
import RankingTable from "./components/RankingTable";
import ReportGenerator from "./components/ReportGenerator";
import IndexRpDashboard from "./components/IndexRpDashboard";
import SectorQuickGlance from "./components/SectorQuickGlance";

const APP_TITLE = "Daily Stock Ranking Movement Dashboard";

const sampleResponse = {
  indices: [],
  data: {},
};

const emptyIndexRpResponse = {
  kpis: null,
  rows: [],
};

const emptyQuickGlanceResponse = {
  summary: null,
  sectors: [],
};

export default function App() {
  const [activeView, setActiveView] = useState("stock");
  const [yesterdayFile, setYesterdayFile] = useState(null);
  const [todayFile, setTodayFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [response, setResponse] = useState(sampleResponse);
  const [selectedIndex, setSelectedIndex] = useState("");
  const [indexRpFile, setIndexRpFile] = useState(null);
  const [indexRpResponse, setIndexRpResponse] = useState(emptyIndexRpResponse);
  const [quickGlanceFile, setQuickGlanceFile] = useState(null);
  const [quickGlanceResponse, setQuickGlanceResponse] = useState(emptyQuickGlanceResponse);
  const [quickGlanceSearch, setQuickGlanceSearch] = useState("");
  const [quickGlanceSortMode, setQuickGlanceSortMode] = useState("performance");

  useEffect(() => {
    if (!selectedIndex && response.indices.length > 0) {
      setSelectedIndex(response.indices[0]);
    }
  }, [response, selectedIndex]);

  const currentIndexData = selectedIndex ? response.data[selectedIndex] : null;

  async function handleCompare() {
    if (!yesterdayFile || !todayFile) {
      setError("Please upload both Excel files.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const data = await compareFiles({ yesterdayFile, todayFile });
      setResponse(data);
      setSelectedIndex(data.indices[0] || "");
    } catch (requestError) {
      const detail =
        requestError?.response?.data?.detail || requestError?.message || "Comparison failed.";
      setError(detail);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="app-shell">
      <header className="app-header">
        <div className="brand-tag">ULJK Bloomberg Ranking System</div>
        <h1>{APP_TITLE}</h1>
        <p>
          Upload yesterday and today Bloomberg metadata files, compare ranking movement by
          index, and inspect Top 10 and Bottom 10 shifts in a terminal-style dashboard.
        </p>
      </header>

      <div className="dashboard-toggle" role="tablist" aria-label="Dashboard view">
        <button
          className={`dashboard-toggle-button ${activeView === "stock" ? "active" : ""}`}
          type="button"
          onClick={() => setActiveView("stock")}
        >
          Stock Dashboard
        </button>
        <button
          className={`dashboard-toggle-button ${activeView === "index-rp" ? "active" : ""}`}
          type="button"
          onClick={() => setActiveView("index-rp")}
        >
          Index RP Dashboard
        </button>
        <button
          className={`dashboard-toggle-button ${activeView === "sector-quick-glance" ? "active" : ""}`}
          type="button"
          onClick={() => setActiveView("sector-quick-glance")}
        >
          Sector Quick Glance
        </button>
      </div>

      {activeView === "stock" ? (
        <>
          <UploadPanel
            yesterdayFile={yesterdayFile}
            todayFile={todayFile}
            loading={loading}
            onYesterdayChange={setYesterdayFile}
            onTodayChange={setTodayFile}
            onCompare={handleCompare}
          />

          <ReportGenerator />

          {error ? <div className="error-banner">{error}</div> : null}

          {response.indices.length > 0 && currentIndexData ? (
            <>
              <div className="top-controls">
                <IndexSelector
                  indices={response.indices}
                  selectedIndex={selectedIndex}
                  onChange={setSelectedIndex}
                />
              </div>

              <KPICards summary={currentIndexData.summary} />

              <div className="tables-grid">
                <RankingTable title="Top 10 Stocks" rows={currentIndexData.top10} type="top" />
                <RankingTable
                  title="Bottom 10 Stocks"
                  rows={currentIndexData.bottom10}
                  type="bottom"
                />
              </div>
            </>
          ) : (
            <section className="empty-state panel">
              <h2>Dashboard Waiting For Comparison</h2>
              <p>
                Upload two Excel files and click <strong>Compare Files</strong> to populate the
                index dropdown, KPI cards, and ranking tables.
              </p>
            </section>
          )}
        </>
      ) : activeView === "index-rp" ? (
        <IndexRpDashboard
          file={indexRpFile}
          onFileChange={setIndexRpFile}
          response={indexRpResponse}
          onResponseChange={setIndexRpResponse}
        />
      ) : (
        <SectorQuickGlance
          file={quickGlanceFile}
          onFileChange={setQuickGlanceFile}
          initialFile={todayFile}
          response={quickGlanceResponse}
          onResponseChange={setQuickGlanceResponse}
          indexRpRows={indexRpResponse.rows}
          search={quickGlanceSearch}
          onSearchChange={setQuickGlanceSearch}
          sortMode={quickGlanceSortMode}
          onSortModeChange={setQuickGlanceSortMode}
        />
      )}
    </div>
  );
}
