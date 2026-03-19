import { useWorkbookStore } from "../store/workbookStore";

export function ValidationPanel() {
  const workbook = useWorkbookStore((s) => s.workbook);

  const issues = workbook?.validationIssues ?? [];
  if (issues.length === 0) {
    return (
      <section className="panel validations">
        <h3>Validation</h3>
        <p>No issues detected.</p>
      </section>
    );
  }

  return (
    <section className="panel validations">
      <h3>Validation ({issues.length})</h3>
      <ul>
        {issues.map((issue, idx) => (
          <li key={`${issue.type}-${idx}`} className="issue-item">
            <strong>{issue.type}</strong>
            <span>{issue.message}</span>
          </li>
        ))}
      </ul>
    </section>
  );
}
