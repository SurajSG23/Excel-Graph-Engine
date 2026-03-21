# Excel Graph Engine

Excel2Graph helps you understand and improve spreadsheet logic by turning an Excel workbook into a visual dependency map.

Instead of tracing formulas manually cell by cell, you can upload a workbook, see how everything connects, make changes, validate your logic, and export an updated Excel file.

## What This Project Does

Excel2Graph transforms spreadsheet logic into a guided workflow:

1. Upload a workbook.
2. View every relevant cell as a connected graph node.
3. Inspect dependencies and downstream impact.
4. Edit formulas and recompute results.
5. Detect formula and dependency issues.
6. Export back to `.xlsx`.

Think of it as a visual workspace for spreadsheet decision models.

## End-to-End Flow

### 1. Start the App

You launch the project and open the web app.

You get:

1. A graph canvas.
2. A control sidebar.
3. Panels for upload, filtering, inspection, validation, and version history.

### 2. Upload Workbook

Upload an `.xlsx` file.

The app will:

1. Read workbook sheets and cells.
2. Build a cell dependency map.
3. Compute initial values.
4. Create an initial version snapshot.

### 3. Explore the Graph

Use the graph to understand workbook behavior.

You can:

1. Filter by sheet.
2. Search by node/cell/formula text.
3. Select nodes to inspect details.
4. See overall stats such as node count, edges, and issue counts.

### 4. Inspect and Edit Nodes

From the node details panel, edit formulas and apply changes.

After applying:

1. The workbook is recomputed.
2. Impacted values update.
3. The graph refreshes.
4. A new version entry is added.

### 5. Validate Quality

The validation panel flags common workbook risks, including:

1. Missing references.
2. Invalid formulas.
3. Circular dependencies.

### 6. Track Changes

Each major action is recorded in a version timeline with:

1. Version number.
2. Label.
3. Timestamp.

### 7. Export Workbook

When satisfied, export to `.xlsx`.

The exported file reflects your latest graph state, formulas, and values.

## Typical Usage Scenario

1. Upload a planning workbook.
2. Focus on one sheet.
3. Trace key output cells back to their drivers.
4. Update formulas for a what-if scenario.
5. Check validation warnings.
6. Review timeline entries.
7. Export and share the updated workbook.

## Quick Start

```bash
npm install
npm run dev
```

- Frontend: http://localhost:5173
- Backend: http://localhost:4100

## Optional: Generate a Sample Workbook

```bash
npm run setup:sample --workspace backend
```
