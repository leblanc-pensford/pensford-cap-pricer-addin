const express = require("express");
const { createProxyMiddleware } = require("http-proxy-middleware");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;

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

// Proxy /api/cap-pricer/* -> https://pensfordcalculators.loanboss.com/cap-pricer/*
app.use(
  "/api/cap-pricer",
  createProxyMiddleware({
    target: "https://pensfordcalculators.loanboss.com",
    changeOrigin: true,
    pathRewrite: { "^/api/cap-pricer": "/cap-pricer" },
  })
);

// Serve static files
app.use(express.static(path.join(__dirname)));

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Pensford Cap Pricer Add-in server running on port ${PORT}`);
});
