import { useEffect, useState } from "react";
import { compareFiles } from "./api";
import UploadPanel from "./components/UploadPanel";
import KPICards from "./components/KPICards";
import IndexSelector from "./components/IndexSelector";
import RankingTable from "./components/RankingTable";
import ReportGenerator from "./components/ReportGenerator";

const APP_TITLE = "Daily Stock Ranking Movement Dashboard";

const sampleResponse = {
  indices: [],
  data: {},
};

export default function App() {
  const [yesterdayFile, setYesterdayFile] = useState(null);
  const [todayFile, setTodayFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [response, setResponse] = useState(sampleResponse);
  const [selectedIndex, setSelectedIndex] = useState("");

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
    </div>
  );
}
