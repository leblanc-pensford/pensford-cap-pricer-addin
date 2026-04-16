const express = require("express");

const app = express();
const PORT = process.env.PORT || 3000;
const LOANBOSS_API_KEY = process.env.LOANBOSS_API_KEY || "";
const LOANBOSS_URL = "https://pensfordcalculators.loanboss.com/cap-pricer/calculate";
const DEFAULT_RATE_INDEX = "c6b1a7bb-4681-11e9-82f6-0242ac120002";

if (!LOANBOSS_API_KEY) {
  console.error("WARNING: LOANBOSS_API_KEY environment variable is not set. All /cap requests will fail.");
}

// CORS headers
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type, Accept");
  if (req.method === "OPTIONS") {
    return res.sendStatus(200);
  }
  next();
});

// GET /
app.get("/", (req, res) => {
  res.type("text/plain").send(
    "Pensford Cap Pricer Proxy - Use GET /cap?notional=...&strike=...&months=...&field=capCost"
  );
});

// GET /health
app.get("/health", (req, res) => {
  res.type("text/plain").send("ok");
});

// GET /cap
app.get("/cap", async (req, res) => {
  const start = Date.now();
  const { notional, strike, effectiveDate, terminationDate, months, rateIndexId, field } = req.query;

  // --- Validation ---
  if (!LOANBOSS_API_KEY) {
    log(req, 500, start);
    return res.status(500).type("text/plain").send("Error: LOANBOSS_API_KEY is not configured on the server");
  }

  if (!notional || isNaN(Number(notional))) {
    log(req, 400, start);
    return res.status(400).type("text/plain").send("Error: notional is required and must be a number");
  }

  if (!strike || isNaN(Number(strike))) {
    log(req, 400, start);
    return res.status(400).type("text/plain").send("Error: strike is required and must be a number");
  }

  if (!terminationDate && !months) {
    log(req, 400, start);
    return res.status(400).type("text/plain").send("Error: either terminationDate or months is required");
  }

  // --- Build JSON body ---
  const body = {
    notional: Number(notional),
    strike: Number(strike),
    effectiveDate: effectiveDate || todayISO(),
    rateIndexId: rateIndexId || DEFAULT_RATE_INDEX,
  };

  if (terminationDate) {
    body.terminationDate = terminationDate;
  } else {
    body.months = parseInt(months, 10);
  }

  // --- POST to LoanBoss ---
  try {
    const upstream = await fetch(LOANBOSS_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "X-Api-Key": LOANBOSS_API_KEY,
      },
      body: JSON.stringify(body),
      signal: AbortSignal.timeout(30000),
    });

    if (!upstream.ok) {
      const errText = await upstream.text();
      log(req, upstream.status, start);
      return res.status(upstream.status).type("text/plain").send("Error: LoanBoss API returned " + upstream.status + " - " + errText);
    }

    const data = await upstream.json();

    // --- Return single field or full JSON ---
    if (field) {
      const value = data[field];
      if (value === undefined) {
        log(req, 400, start);
        return res.status(400).type("text/plain").send("Error: field '" + field + "' not found in response");
      }
      log(req, 200, start);
      return res.type("text/plain").send(String(value));
    }

    log(req, 200, start);
    return res.json(data);

  } catch (err) {
    if (err.name === "TimeoutError") {
      log(req, 504, start);
      return res.status(504).type("text/plain").send("Error: LoanBoss API timed out after 30 seconds");
    }
    console.error("Proxy error:", err.message);
    log(req, 502, start);
    return res.status(502).type("text/plain").send("Error: " + err.message);
  }
});

// --- Helpers ---

function todayISO() {
  const d = new Date();
  return d.getFullYear() + "-" +
    String(d.getMonth() + 1).padStart(2, "0") + "-" +
    String(d.getDate()).padStart(2, "0");
}

function log(req, status, start) {
  const duration = Date.now() - start;
  console.log(`${req.method} ${req.path} ${JSON.stringify(req.query)} -> ${status} (${duration}ms)`);
}

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Pensford Cap Pricer Proxy running on port ${PORT}`);
});
