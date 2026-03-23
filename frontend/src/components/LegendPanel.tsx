export function LegendPanel() {
  return (
    <section className="panel legend-panel">
      <h3>Legend</h3>
      <ul>
        <li>
          <i className="legend-dot input" /> Input node
        </li>
        <li>
          <i className="legend-dot computed" /> Computed node
        </li>
        <li>
          <i className="legend-dot output" /> Output node
        </li>
        <li>
          <i className="legend-dot error" /> Error or cycle node
        </li>
        <li>
          <i className="legend-line same" /> Same-file dependency
        </li>
        <li>
          <i className="legend-line cross" /> Cross-file dependency
        </li>
      </ul>
    </section>
  );
}
