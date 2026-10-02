const express = require("express");
const multer = require("multer");
const fs = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const { extractLayout } = require("./docx-layout-extractor");
const { generateWordDocument } = require("../generators/word-generator");

const app = express();
const host = process.env.BACKEND_HOST ?? "127.0.0.1";
const port = Number(process.env.BACKEND_PORT ?? 3000);
const outputDir = path.join(__dirname, "..", "output");
const layoutsDir = path.join(outputDir, "layouts");

async function saveLayout(layout) {
  const layoutId = crypto.randomUUID();
  await fs.mkdir(layoutsDir, { recursive: true });
  await fs.writeFile(
    path.join(layoutsDir, `${layoutId}.json`),
    `${JSON.stringify({ ...layout, layoutId }, null, 2)}\n`,
    "utf8",
  );
  return layoutId;
}

function validLayoutId(layoutId) {
  return /^[\da-f]{8}-[\da-f]{4}-[\da-f]{4}-[\da-f]{4}-[\da-f]{12}$/i.test(
    layoutId,
  );
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024,
    files: 1,
  },
  fileFilter: (_request, file, callback) => {
    if (!file.originalname.toLowerCase().endsWith(".docx")) {
      callback(new Error("Please upload a .docx file."));
      return;
    }
    callback(null, true);
  },
});

app.use(express.json({ limit: "20mb" }));

app.get("/api/health", (_request, response) => {
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
      const layout = await extractLayout(
        request.file.buffer,
        request.file.originalname,
      );
      const layoutId = await saveLayout(layout);
      response.json({ ...layout, layoutId });
    } catch (error) {
      response.status(400).json({
        ok: false,
        error: error.message || "Could not extract the Word layout.",
      });
    }
  },
);

app.get("/api/layout/:layoutId", async (request, response, next) => {
  if (!validLayoutId(request.params.layoutId)) {
    response.status(400).json({ ok: false, error: "Invalid layout ID." });
    return;
  }

  try {
    const layoutPath = path.join(layoutsDir, `${request.params.layoutId}.json`);
    const layout = JSON.parse(await fs.readFile(layoutPath, "utf8"));
    response.json(layout);
  } catch (error) {
    if (error.code === "ENOENT") {
      response.status(404).json({ ok: false, error: "Layout JSON not found." });
      return;
    }
    next(error);
  }
});


app.post(
  "/api/document/process",
  upload.single("file"),
  async (request, response, next) => {
    if (!request.file) {
      response
        .status(400)
        .json({ ok: false, error: "Choose a .docx file to upload." });
      return;
    }

    let outputPath;
    try {
      const layout = await extractLayout(
        request.file.buffer,
        request.file.originalname,
      );
      const layoutId = await saveLayout(layout);
      outputPath = path.join(outputDir, `generated-${layoutId}.docx`);
      await generateWordDocument({
        reportData: {},
        reportLayout: layout,
        outputPath,
      });
      response.set("X-Layout-Id", layoutId);
      response.set("X-Layout-Json", `/api/layout/${layoutId}`);
      response.download(outputPath, "generated-layout.docx", async (error) => {
        await fs.rm(outputPath, { force: true }).catch(() => {});
        if (error && !response.headersSent) next(error);
      });
    } catch (error) {
      if (outputPath) await fs.rm(outputPath, { force: true }).catch(() => {});
      next(error);
    }
  },
);

app.post("/api/layout/generate", async (request, response, next) => {
  try {
    const layout = request.body;
    if (!layout?.document?.page || !Array.isArray(layout.sections)) {
      response.status(400).json({
        ok: false,
        error: "Please send a valid extracted layout JSON.",
      });
      return;
    }

    await fs.mkdir(outputDir, { recursive: true });
    const outputPath = path.join(
      outputDir,
      `generated-layout-${Date.now()}.docx`,
    );
    await generateWordDocument({
      reportData: {},
      reportLayout: layout,
      outputPath,
    });
    response.download(outputPath, "generated-layout.docx", async (error) => {
      await fs.rm(outputPath, { force: true }).catch(() => {});
      if (error && !response.headersSent) next(error);
    });
  } catch (error) {
    next(error);
  }
});

app.use((error, _request, response, _next) => {
  if (response.headersSent) return;
  const statusCode =
    error instanceof multer.MulterError && error.code === "LIMIT_FILE_SIZE"
      ? 413
      : (error.status ?? error.statusCode ?? 400);
  response.status(statusCode).json({
    ok: false,
    error: error.message || "The request could not be processed.",
  });
});

app.use((_request, response) => {
  response.status(404).json({
    ok: false,
    error: "Route not found",
  });
});

app.listen(port, host, () => {
  console.log(`Backend listening at http://${host}:${port}`);
});
