import { existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import express from "express";
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
  const startTime = Date.now();
  try {
    const buffer = req.body as Buffer;
    if (!buffer || buffer.length === 0) {
      console.warn(`[convert] Bad request: empty body`);
      res.status(400).json({ error: "No file provided" });
      return;
    }

    console.log(`[convert] Received ${(buffer.length / 1024).toFixed(1)} KB`);

    const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer;

    const eml = msgToEml(arrayBuffer);

    const elapsed = Date.now() - startTime;
    console.log(`[convert] Success in ${elapsed}ms (output ${(eml.length / 1024).toFixed(1)} KB)`);

    res.setHeader("Content-Type", "message/rfc822");
    res.setHeader("Content-Disposition", 'attachment; filename="converted.eml"');
    res.send(eml);
  } catch (error) {
    const elapsed = Date.now() - startTime;
    console.error(`[convert] Failed after ${elapsed}ms:`, error instanceof Error ? error.stack : error);
    res.status(500).json({
      error: "Failed to convert file",
      details: error instanceof Error ? error.message : String(error),
    });
  }
});

// Fallback to index.html for SPA
app.get("*", (_req, res) => {
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
