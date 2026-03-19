import { Router } from "express";
import multer from "multer";
import path from "node:path";
import fs from "node:fs";
import { workbookController } from "../controllers/workbookController";

const router = Router();

const uploadDir = path.resolve(process.cwd(), "uploads");
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, uploadDir),
  filename: (_req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`)
});

const upload = multer({
  storage,
  fileFilter: (_req, file, cb) => {
    const isExcel = file.mimetype.includes("sheet") || file.originalname.endsWith(".xlsx");
    if (!isExcel) {
      cb(new Error("Only .xlsx files are allowed."));
      return;
    }
    cb(null, true);
  }
});

router.post("/upload", upload.single("file"), (req, res) => workbookController.uploadWorkbook(req, res));
router.post("/recompute", (req, res) => workbookController.recomputeWorkbook(req, res));
router.post("/export", (req, res) => workbookController.exportWorkbook(req, res));

export default router;
