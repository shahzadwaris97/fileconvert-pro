
import express from "express";
import multer from "multer";
import cors from "cors";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";
import { randomUUID } from "node:crypto";

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors({
  origin: [
    "http://localhost:5173",
    "http://localhost:5174"
  ]
}));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 },
  fileFilter(req, file, cb) {
    const ext = path.extname(file.originalname).toLowerCase();
    cb(null, [".xlsx", ".xls", ".ods"].includes(ext));
  }
});


function convertWithLibreOffice(inputPath, outputDir) {
  return new Promise(async (resolve, reject) => {
    const executable = process.env.LIBREOFFICE_PATH ||
      "C:\\Program Files\\LibreOffice\\program\\soffice.exe";

    const profileDir = path.join(
      os.tmpdir(),
      "lo-profile-" + randomUUID()
    );

    try {
      await fs.mkdir(profileDir, { recursive: true });

      const profilePath =
        `file:///${profileDir.replace(/\\/g, "/")}`;

      const child = spawn(executable, [
        `-env:UserInstallation=${profilePath}`,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", outputDir,
        inputPath
      ]);

      let stderr = "";

      child.stderr.on("data", data => {
        stderr += data.toString();
      });

      child.on("error", reject);

      child.on("close", async code => {
        await fs.rm(profileDir, { recursive: true, force: true })
          .catch(() => {});

        if (code === 0) {
          resolve();
        } else {
          reject(new Error(
            stderr || `LibreOffice exited with code ${code}`
          ));
        }
      });
    } catch (error) {
      await fs.rm(profileDir, { recursive: true, force: true })
        .catch(() => {});
      reject(error);
    }
  });
}

app.get("/health", (req, res) => {
  res.json({ ok: true, service: "Umar Convert Excel PDF" });
});

app.post("/api/excel-to-pdf", upload.single("file"), async (req, res) => {
  let tempDir;

  try {
    if (!req.file) {
      return res.status(400).json({
        error: "Please upload an XLSX, XLS, or ODS file."
      });
    }

    tempDir = await fs.mkdtemp(
      path.join(os.tmpdir(), "umar-convert-")
    );

    const ext = path.extname(req.file.originalname).toLowerCase();
    const inputPath = path.join(tempDir, `input${ext}`);

    await fs.writeFile(inputPath, req.file.buffer);

    await convertWithLibreOffice(inputPath, tempDir);

    const outputPath = path.join(tempDir, "input.pdf");
    const pdf = await fs.readFile(outputPath);

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="converted.pdf"'
    );

    res.send(pdf);
  } catch (error) {
    console.error("Excel conversion failed:", error.message);
    if (!res.headersSent) {
      res.status(500).json({
        error: "Excel conversion failed. Please try another file."
      });
    }
  } finally {
    if (tempDir) {
      await fs.rm(tempDir, { recursive: true, force: true })
        .catch(() => {});
    }
  }
});

app.listen(PORT, () => {
  console.log(`Umar Convert backend running on port ${PORT}`);
});
