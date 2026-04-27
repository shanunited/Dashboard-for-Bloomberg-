export default function IndexSelector({ indices, selectedIndex, onChange }) {
  return (
    <section className="panel selector-panel">
      <div className="panel-header compact">
        <h2>Index Filter</h2>
      </div>
      <select
        className="index-select"
        value={selectedIndex}
        onChange={(event) => onChange(event.target.value)}
      >
        {indices.map((indexName) => (
          <option key={indexName} value={indexName}>
            {indexName}
          </option>
        ))}
      </select>
    </section>
  );
}
