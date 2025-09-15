// server.js
import express from "express";
import morgan from "morgan";
import { spawn } from "child_process";
import { v4 as uuidv4 } from "uuid";
import path from "path";
import fs from "fs";

// Ensure paths resolve relative to project root
const __dirname = path.resolve();

const app = express();
const PORT = process.env.PORT || 3000;

// Parse JSON bodies (increase limit for large boards)
app.use(express.json({ limit: "20mb" }));
app.use(morgan("dev"));

// Health check
app.get("/healthz", (_req, res) => res.json({ ok: true }));

/**
 * POST /convert
 * Accepts a Miro JSON object in request body
 * Query params:
 *   ?download=true   -> streams back .pptx file
 */
app.post("/convert", async (req, res) => {
  try {
    const boardJson = req.body;
    if (!boardJson || typeof boardJson !== "object") {
      return res.status(400).json({ error: "Expected JSON body" });
    }

    // Unique id for this conversion
    const id = uuidv4();
    const inputPath = path.join(__dirname, "converter", `${id}.json`);
    const outputPath = path.join(__dirname, "output", `${id}.pptx`);

    // Save JSON input
    fs.writeFileSync(inputPath, JSON.stringify(boardJson, null, 2), "utf-8");

    // Call Python converter
    const py = spawn("python3", [
      "converter/convert.py",
      "--input",
      inputPath,
      "--output",
      outputPath,
    ]);

    let stdout = "";
    let stderr = "";

    py.stdout.on("data", (d) => (stdout += d.toString()));
    py.stderr.on("data", (d) => (stderr += d.toString()));

    py.on("close", (code) => {
      // Clean up input JSON after processing
      fs.unlinkSync(inputPath);

      if (code !== 0) {
        console.error("Python error:", stderr);
        return res
          .status(500)
          .json({ error: "Conversion failed", details: stderr.trim() });
      }

      if (req.query.download === "true") {
        res.setHeader(
          "Content-Type",
          "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        );
        res.setHeader(
          "Content-Disposition",
          `attachment; filename="${id}.pptx"`
        );
        fs.createReadStream(outputPath).pipe(res);
      } else {
        res.json({
          ok: true,
          id,
          output: outputPath,
          log: stdout.trim(),
        });
      }
    });
  } catch (err) {
    console.error("Server error:", err);
    res.status(500).json({ error: "Internal error", details: String(err) });
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Server ready on http://localhost:${PORT}`);
});
