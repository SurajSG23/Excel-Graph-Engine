# Excel Graph Engine

Excel Graph Engine is a production-ready full-stack application that turns Excel workbooks into a graph workflow you can inspect, edit, recompute, and export back to `.xlsx`.

## Stack

- Frontend: React + TypeScript + Vite + React Flow + Zustand + Axios
- Backend: Node.js + Express + TypeScript + Multer + xlsx + formulajs

## Quick Start

```bash
npm install
npm run dev
```

- Frontend: http://localhost:5173
- Backend: http://localhost:4100

## Features

- Upload `.xlsx` workbooks with multiple sheets
- Build a unified workbook graph (`Sheet!Cell` IDs)
- Visualize dependencies with React Flow
- Edit formulas and recalculate incrementally
- Detect cycles, invalid formulas, and missing references
- Export graph back to Excel while preserving sheets/formulas
- Version snapshots for scenario testing

## API

- `POST /api/upload` - Upload and parse workbook
- `POST /api/recompute` - Update formulas/values and recompute impacted nodes
- `POST /api/export` - Export current graph state to `.xlsx`

## Sample

A generated workbook is available under `samples/` after running:

```bash
npm run setup:sample --workspace backend
```
