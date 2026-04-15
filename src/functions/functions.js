/**
 * =============================================================================
 * Pensford Cap Pricer — Excel Custom Function
 * =============================================================================
 *
 * Registers =PENSFORD.CAP() as a streaming custom function in Excel.
 * Calls the LoanBoss Cap Pricer API and returns pricing data.
 */

const API_URL = "https://pensfordcalculators.loanboss.com/cap-pricer/calculate";
const DEFAULT_INDEX_UUID = "c6b1a7bb-4681-11e9-82f6-0242ac120002"; // 1-Month Term SOFR

// Friendly index names → UUIDs
const INDEX_MAP = {
  "tsfr1m": "c6b1a7bb-4681-11e9-82f6-0242ac120002",
  "nyfed":  "d48732b4-6011-4ef5-8098-98691f75e7ec",
  "30d avg sofr": "d48732b4-6011-4ef5-8098-98691f75e7ec",
  "30 day avg sofr": "d48732b4-6011-4ef5-8098-98691f75e7ec"
};

// Output keys → API response fields
const OUTPUT_MAP = {
  "bankadj":    "bankAdjustment",
  "capcost":    "capCost",
  "mtm":        "markToMarketPrice",
  "mtmpercent": "markToMarketPercentOfNotional",
  "mid":        "midMarketPrice",
  "netbenefit": "netBenefit",
  "pv":         "presentValueOfTotalPayout"
};

// Header display labels
const HEADER_MAP = {
  "bankadj":    "Bank Adj",
  "capcost":    "Cap Cost",
  "mtm":        "MTM Price",
  "mtmpercent": "MTM % of Notional",
  "mid":        "Mid-Market Price",
  "netbenefit": "Net Benefit",
  "pv":         "PV Total Payout"
};

const ALL_OUTPUT_KEYS = ["bankadj", "capcost", "mtm", "mtmpercent", "mid", "netbenefit", "pv"];


/**
 * Prices an interest rate cap via the LoanBoss Cap Pricer API.
 * @customfunction
 * @param {number} notional Loan notional amount the cap is based on (e.g. 25000000 for $25M).
 * @param {number} strike Strike rate (rate ceiling). Enter as percent (e.g. 4.5 for 4.50%) or decimal (e.g. 0.045).
 * @param {string} [effective] Effective (start) date of the cap in YYYY-MM-DD format. Defaults to today if left blank. Example: "2026-06-01"
 * @param {string} [termination] End date OR term in months. Provide a date (e.g. "2029-06-01") or an integer for months (e.g. 36 for 3 years). Required.
 * @param {string} [index] Rate index the cap hedges against. "TSFR1M" (1-Month Term SOFR, default) or "NYFED" (NY Fed SOFR). Example: "TSFR1M"
 * @param {string} [output] Pricing field(s) to return as comma-separated list. Options: bankAdj, capCost, mtm, mtmPercent, mid, netBenefit, pv. Blank returns ALL fields.
 * @param {boolean} [headers] Include header labels. TRUE (default) or FALSE. Leave blank for TRUE.
 * @param {string} [direction] Spill direction. "H" (horizontal, default) or "V" (vertical). Leave blank for "H".
 * @returns {any[][]} Pricing result(s) as a spilled array.
 */
async function cap(notional, strike, effective, termination, index, output, headers, direction) {

  // ── Get API key from storage ──────────────────────────────────────────────
  const apiKey = await getStoredApiKey();
  if (!apiKey) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "API key not set. Open the Pensford Cap Pricer taskpane to configure it."
    );
  }

  // ── Validate required params ──────────────────────────────────────────────
  if (!notional) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "notional is required (e.g. 25000000)"
    );
  }
  if (strike === undefined || strike === null || strike === "") {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "strike is required (e.g. 4.5 for 4.50%)"
    );
  }

  // ── Normalize strike ──────────────────────────────────────────────────────
  const strikeDecimal = (strike > 0.20) ? strike / 100 : strike;

  // ── Effective date ────────────────────────────────────────────────────────
  let effectiveDate;
  if (!effective || effective === "") {
    effectiveDate = formatDate(new Date());
  } else if (typeof effective === "number") {
    // Excel serial date number
    effectiveDate = formatDate(excelSerialToDate(effective));
  } else {
    effectiveDate = String(effective);
  }

  // ── Termination / months ──────────────────────────────────────────────────
  if (!termination && termination !== 0) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "termination is required — provide a date (YYYY-MM-DD) or months (e.g. 36)"
    );
  }

  let terminationDate = null;
  let months = null;

  if (typeof termination === "number") {
    if (termination <= 120 && Number.isInteger(termination)) {
      months = termination;
    } else {
      // Excel serial date
      terminationDate = formatDate(excelSerialToDate(termination));
    }
  } else {
    const parsed = parseInt(String(termination), 10);
    if (!isNaN(parsed) && String(parsed) === String(termination).trim() && parsed <= 120) {
      months = parsed;
    } else {
      terminationDate = String(termination);
    }
  }

  // ── Rate index ────────────────────────────────────────────────────────────
  let rateIndexId;
  if (!index || index === "") {
    rateIndexId = DEFAULT_INDEX_UUID;
  } else {
    const indexLower = String(index).toLowerCase().trim();
    if (INDEX_MAP[indexLower]) {
      rateIndexId = INDEX_MAP[indexLower];
    } else if (index.length > 10 && index.indexOf("-") > -1) {
      rateIndexId = String(index);
    } else {
      throw new CustomFunctions.Error(
        CustomFunctions.ErrorCode.invalidValue,
        'Unknown index: "' + index + '". Use TSFR1M, NYFED, or a UUID.'
      );
    }
  }

  // ── Output fields ─────────────────────────────────────────────────────────
  let outputKeys;
  if (!output || output === "") {
    outputKeys = ALL_OUTPUT_KEYS;
  } else {
    outputKeys = String(output).toLowerCase().replace(/[\s"']/g, "").split(",").filter(Boolean);
    for (const key of outputKeys) {
      if (!OUTPUT_MAP[key]) {
        throw new CustomFunctions.Error(
          CustomFunctions.ErrorCode.invalidValue,
          'Unknown output: "' + key + '". Options: ' + ALL_OUTPUT_KEYS.join(", ")
        );
      }
    }
  }

  // ── Headers / direction ───────────────────────────────────────────────────
  const showHeaders = (headers === false || headers === "FALSE" || headers === "false" || headers === 0) ? false : true;
  const dir = (!direction || direction === "") ? "H" : String(direction).toUpperCase().charAt(0);

  // ── Build request body ────────────────────────────────────────────────────
  const body = {
    notional: notional,
    strike: strikeDecimal,
    effectiveDate: effectiveDate,
    rateIndexId: rateIndexId
  };
  if (terminationDate) {
    body.terminationDate = terminationDate;
  } else if (months) {
    body.months = months;
  }

  // ── Call the API ──────────────────────────────────────────────────────────
  let response;
  try {
    response = await fetch(API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "X-Api-Key": apiKey
      },
      body: JSON.stringify(body)
    });
  } catch (e) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "Network error: " + e.message
    );
  }

  if (response.status === 401 || response.status === 403) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "API authentication failed (HTTP " + response.status + "). Check your API key."
    );
  }
  if (!response.ok) {
    const errText = await response.text();
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "API error (HTTP " + response.status + "): " + errText
    );
  }

  const data = await response.json();

  // ── Extract fields ────────────────────────────────────────────────────────
  const headerLabels = [];
  const values = [];
  for (const key of outputKeys) {
    const apiField = OUTPUT_MAP[key];
    let val = data[apiField];
    if (val === undefined) val = "N/A";
    headerLabels.push(HEADER_MAP[key] || key);
    values.push(val);
  }

  // ── Build 2D output array ─────────────────────────────────────────────────
  if (dir === "V") {
    // Vertical: each field is a row
    if (showHeaders) {
      return outputKeys.map((_, i) => [headerLabels[i], values[i]]);
    } else {
      return outputKeys.map((_, i) => [values[i]]);
    }
  } else {
    // Horizontal: each field is a column
    if (showHeaders) {
      return [headerLabels, values];
    } else {
      return [values];
    }
  }
}


// ─── Helper Functions ────────────────────────────────────────────────────────

function formatDate(d) {
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function excelSerialToDate(serial) {
  // Excel serial date: days since 1900-01-01 (with the 1900 leap year bug)
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400 * 1000;
  return new Date(utcValue);
}

async function getStoredApiKey() {
  try {
    return localStorage.getItem("pensford_cap_api_key") || null;
  } catch (e) {
    // localStorage may not be available in all contexts
    return null;
  }
}


// ─── Register the function ───────────────────────────────────────────────────

CustomFunctions.associate("CAP", cap);
