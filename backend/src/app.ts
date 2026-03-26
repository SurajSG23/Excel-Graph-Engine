import express from "express";
import cors from "cors";
import workbookRoutes from "./routes/workbookRoutes";

export const app = express();

app.use(cors());
app.use(express.json({ limit: "5mb" }));

app.get("/api/health", (_req, res) => {
  res.status(200).json({ status: "ok", service: "excel2graph-pipeline-backend" });
});

app.use("/api", workbookRoutes);
