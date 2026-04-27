function formatValue(value) {
  if (value === null || value === undefined || value === "") {
    return "--";
  }
  if (typeof value === "number") {
    return Number.isInteger(value) ? String(value) : value.toFixed(2);
  }
  return value;
}

function formatPercent(value) {
  if (value === null || value === undefined || value === "") {
    return "--";
  }
  if (typeof value === "number") {
    return `${(value * 100).toFixed(2)}%`;
  }
  return value;
}

function getChangeClass(change) {
  if (!change) {
    return "change-neutral";
  }
  if (change.includes("New Entry")) {
    return "change-new";
  }
  if (change.includes("Unchanged")) {
    return "change-neutral";
  }
  if (change.includes("Up")) {
    return "change-positive";
  }
  if (change.includes("Down")) {
    return "change-negative";
  }
  return "change-neutral";
}

export default function RankingTable({ title, rows, type }) {
  const rankLabel = type === "top" ? "Top Rank" : "Bottom Rank";
  const yesterdayRankLabel =
    type === "top" ? "Yesterday Top Rank" : "Yesterday Bottom Rank";

  return (
    <section className="panel table-panel">
      <div className="panel-header compact">
        <h2>{title}</h2>
      </div>
      <div className="table-wrapper">
        <table className="ranking-table">
          <thead>
            <tr>
              <th>{rankLabel}</th>
              <th>Company Name</th>
              <th>Today Total</th>
              <th>{yesterdayRankLabel}</th>
              <th>Change</th>
              <th>CMP</th>
              <th>FS</th>
              <th>TS</th>
              <th>RSI</th>
              <th>30WMA Value</th>
              <th>30WMA %</th>
              <th>ROCE</th>
              <th>ROE</th>
              <th>EPS Growth</th>
              <th>D/E</th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan="15" className="empty-cell">
                  No rows available.
                </td>
              </tr>
            ) : (
              rows.map((row) => {
                const rankValue = type === "top" ? row.top_rank : row.bottom_rank;
                const yesterdayRank =
                  type === "top" ? row.yesterday_top_rank : row.yesterday_bottom_rank;

                return (
                  <tr key={`${title}-${row.company_name}-${rankValue}`}>
                    <td>{formatValue(rankValue)}</td>
                    <td className="company-cell" title={row.company_name || ""}>
                      <span className="company-text">
                        {formatValue(row.company_name)}
                      </span>
                    </td>
                    <td>{formatValue(row.today_total)}</td>
                    <td>{formatValue(yesterdayRank)}</td>
                    <td>
                      <span className={`change-pill ${getChangeClass(row.change)}`}>
                        {formatValue(row.change)}
                      </span>
                    </td>
                    <td>{formatValue(row.cmp)}</td>
                    <td>{formatValue(row.fs)}</td>
                    <td>{formatValue(row.ts)}</td>
                    <td>{formatValue(row.rsi)}</td>
                    <td>{formatValue(row["30wma"])}</td>
                    <td>{formatPercent(row["30wma_pct"])}</td>
                    <td>{formatValue(row.roce)}</td>
                    <td>{formatValue(row.roe)}</td>
                    <td>{formatValue(row.eps_growth)}</td>
                    <td>{formatValue(row.d_e)}</td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
    </section>
  );
}
