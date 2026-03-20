import { useWorkbookStore } from "../store/workbookStore";

export function ValidationPanel() {
  const workbook = useWorkbookStore((s) => s.workbook);

  const issues = workbook?.validationIssues ?? [];
  if (issues.length === 0) {
    return (
      <section className="panel validations">
        <details className="panel-collapsible" open>
          <summary>Validation</summary>
          <p>No issues detected.</p>
        </details>
      </section>
    );
  }

  return (
    <section className="panel validations">
      <details className="panel-collapsible" open>
        <summary>Validation ({issues.length})</summary>
        <ul>
          {issues.map((issue, idx) => (
            <li key={`${issue.type}-${idx}`} className="issue-item">
              <strong>{issue.type}</strong>
              <span>{issue.message}</span>
            </li>
          ))}
        </ul>
      </details>
    </section>
  );
}
