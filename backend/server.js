const express = require("express");
const multer = require("multer");
const { extractLayout } = require("./docx-layout-extractor");
const path = require("node:path");
const { generateWordDocument } = require("../generators/word-generator");

const app = express();
const host = process.env.BACKEND_HOST ?? "127.0.0.1";
const port = Number(process.env.BACKEND_PORT ?? 3000);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024,
    files: 1,
  },
  fileFilter: (request, file, callback) => {
    if (!file.originalname.toLowerCase().endsWith(".docx")) {
      callback(new Error("Please upload a .docx file."));
      return;
    }

    callback(null, true);
  },
});

app.use(express.json());
app.use(express.json({ limit: "10mb" }));

app.get("/api/health", (request, response) => {
  response.json({
    ok: true,
    service: "residential-valuation-report-backend",
    status: "ready",
  });
});

app.post(
  "/api/layout/extract",
  upload.single("file"),
  async (request, response) => {
    if (!request.file) {
      response.status(400).json({
        ok: false,
        error: "Choose a .docx file to upload.",
      });
      return;
    }

    try {
      const layout = await extractLayout(request.file.buffer);
      response.json(layout);
    } catch (error) {
      response.status(400).json({
        ok: false,
        error: error.message || "Could not extract the Word layout.",
      });
    }
  },
);

app.post("/api/layout/generate", async (req, res, next) => {
  try {
    const layout = req.body;

    if (!layout?.document?.page || !Array.isArray(layout.sections)) {
      return res.status(400).json({
        ok: false,
        error: "Please send a valid extracted layout JSON.",
      });
    }

    const outputPath = path.join(
      __dirname,
      "..",
      "output",
      "generated-layout.docx",
    );

    await generateWordDocument({
      reportData: {},
      reportLayout: layout,
      outputPath,
    });

    res.download(outputPath, "generated-layout.docx");
  } catch (error) {
    next(error);
  }
});

app.use((error, request, response, next) => {
  const statusCode =
    error instanceof multer.MulterError && error.code === "LIMIT_FILE_SIZE"
      ? 413
      : 400;

  response.status(statusCode).json({
    ok: false,
    error: error.message || "The request could not be processed.",
  });
});

app.use((request, response) => {
  response.status(404).json({
    ok: false,
    error: "Route not found",
  });
});

app.listen(port, host, () => {
  console.log(`Backend listening at http://${host}:${port}`);
});
