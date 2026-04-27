export default function KPICards({ summary }) {
  const items = [
    ["Stocks in Index", summary?.stocks_in_index ?? 0],
    ["New Top 10 Entries", summary?.new_top10_entries ?? 0],
    ["New Bottom 10 Entries", summary?.new_bottom10_entries ?? 0],
    ["Top 10 Unchanged", summary?.top10_unchanged ?? 0],
    ["Bottom 10 Unchanged", summary?.bottom10_unchanged ?? 0],
  ];

  return (
    <section className="kpi-grid">
      {items.map(([label, value]) => (
        <div className="kpi-card" key={label}>
          <div className="kpi-label">{label}</div>
          <div className="kpi-value">{value}</div>
        </div>
      ))}
    </section>
  );
}
