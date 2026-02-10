import express from "express";
import { existsSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { msgToEml } from "./msg-to-eml.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT || 3000;

// Serve static files from client directory
const clientPath = join(__dirname, "..", "client");
app.use(express.static(clientPath));

// Parse raw binary data for file uploads
app.use("/api/convert", express.raw({ type: "*/*", limit: "50mb" }));

// Health check
app.get("/api/health", (_req, res) => {
  res.json({ status: "ok" });
});

// Convert single MSG file
app.post("/api/convert", async (req, res) => {
  try {
    const buffer = req.body as Buffer;
    if (!buffer || buffer.length === 0) {
      res.status(400).json({ error: "No file provided" });
      return;
    }

    const arrayBuffer = buffer.buffer.slice(
      buffer.byteOffset,
      buffer.byteOffset + buffer.byteLength
    ) as ArrayBuffer;

    const eml = msgToEml(arrayBuffer);

    res.setHeader("Content-Type", "message/rfc822");
    res.setHeader("Content-Disposition", 'attachment; filename="converted.eml"');
    res.send(eml);
  } catch (error) {
    console.error("Conversion error:", error);
    res.status(500).json({
      error: "Failed to convert file",
      details: error instanceof Error ? error.message : String(error)
    });
  }
});

// Fallback to index.html for SPA
app.get("*", (req, res) => {
  const indexPath = join(clientPath, "index.html");
  if (existsSync(indexPath)) {
    res.sendFile(indexPath);
  } else {
    res.status(404).send("Client not found. Run build first.");
  }
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
