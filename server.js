const express = require("express");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = process.env.PORT || 3000;
const LOANBOSS_BASE = "https://pensfordcalculators.loanboss.com";

// CORS headers for all responses
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type, Accept, X-Api-Key");
  if (req.method === "OPTIONS") {
    return res.sendStatus(200);
  }
  next();
});

// Parse JSON bodies for the proxy
app.use("/api/cap-pricer", express.json());

// Proxy: /api/cap-pricer/* -> LoanBoss /cap-pricer/*
app.all("/api/cap-pricer/*", async (req, res) => {
  const targetPath = req.path.replace(/^\/api\/cap-pricer/, "/cap-pricer");
  const targetUrl = LOANBOSS_BASE + targetPath;

  const headers = {
    "Content-Type": "application/json",
    "Accept": "application/json",
  };
  if (req.headers["x-api-key"]) {
    headers["X-Api-Key"] = req.headers["x-api-key"];
  }

  const fetchOpts = {
    method: req.method,
    headers: headers,
  };
  if (req.method === "POST" && req.body) {
    fetchOpts.body = JSON.stringify(req.body);
  }

  try {
    const upstream = await fetch(targetUrl, fetchOpts);
    const body = await upstream.text();
    res.status(upstream.status);
    res.set("Content-Type", upstream.headers.get("content-type") || "application/json");
    res.send(body);
  } catch (err) {
    console.error("Proxy error:", err.message);
    res.status(502).json({ error: "Proxy error: " + err.message });
  }
});

// Serve static files with .html extension support
app.use(express.static(path.join(__dirname), {
  extensions: ["html", "htm"]
}));

// Fallback: try appending .html
app.use((req, res, next) => {
  if (req.method === "GET" && !path.extname(req.path)) {
    const htmlPath = path.join(__dirname, req.path + ".html");
    if (fs.existsSync(htmlPath)) {
      return res.sendFile(htmlPath);
    }
  }
  next();
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Pensford Cap Pricer Add-in server running on port ${PORT}`);
});
