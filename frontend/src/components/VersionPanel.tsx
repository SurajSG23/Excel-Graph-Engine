import { useWorkbookStore } from "../store/workbookStore";

export function VersionPanel() {
  const versions = useWorkbookStore((s) => s.versions);

  return (
    <section className="panel versions">
      <details className="panel-collapsible versions-collapsible" open>
        <summary>Version Timeline</summary>
        <ul>
          {versions
            .slice()
            .reverse()
            .map((version) => (
              <li key={version.version}>
                <span>v{version.version}</span>
                <span>{version.label}</span>
                <time>{new Date(version.timestamp).toLocaleString()}</time>
              </li>
            ))}
        </ul>
      </details>
    </section>
  );
}
