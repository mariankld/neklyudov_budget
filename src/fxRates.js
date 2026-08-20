const axios = require("axios");

/**
 * Historical currency -> HKD conversion.
 *
 * Every expense's "Sum (HKD)" and "Курс" must be frozen to the rate on the day the
 * expense happened, never recalculated later even if today's rate changes. So instead
 * of a live VLOOKUP against a single "today's rate" table, the bot fetches the
 * *historical* rate for that specific date the first time it's needed and writes it
 * as a plain number.
 *
 * Sources:
 * - Frankfurter (api.frankfurter.dev) — free, keyless, ECB-sourced, historical since
 *   1948. Used for every tracked currency except RUB.
 * - Central Bank of Russia daily XML feed (cbr.ru) — used for RUB specifically, because
 *   the ECB suspended publishing a EUR/RUB reference rate on 2022-03-01, so Frankfurter
 *   has no RUB data after that date.
 */

const TRACKED_CURRENCIES = ["HKD", "USD", "EUR", "GBP", "THB", "CNY", "RUB"];

const FRANKFURTER_BASE = "https://api.frankfurter.dev/v1";
const CBR_DAILY_URL = "https://www.cbr.ru/scripts/XML_daily.asp";

/** key: "CUR|dd/mm/yyyy" -> rate (HKD per 1 unit of CUR) */
const rateCache = new Map();
/** key: "CUR|dd/mm/yyyy" -> in-flight Promise, to dedupe concurrent lookups for the same day */
const inflight = new Map();

function cacheKey(currency, ddmmyyyy) {
  return `${String(currency || "").trim().toUpperCase()}|${ddmmyyyy}`;
}

/** dd/mm/yyyy -> yyyy-mm-dd (Frankfurter's date format) */
function toIsoDate(ddmmyyyy) {
  const m = String(ddmmyyyy || "").match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) throw new Error(`Invalid date for FX lookup: ${ddmmyyyy}`);
  const [, d, mo, y] = m;
  return `${y}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
}

/** Never request a future date — Frankfurter/CBR only have data up to "today". Clamped, not rejected. */
function clampToToday(ddmmyyyy) {
  const m = String(ddmmyyyy || "").match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return ddmmyyyy;
  const [, d, mo, y] = m;
  const requested = new Date(Number(y), Number(mo) - 1, Number(d));
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return requested > today ? formatDdMmYyyy(today) : ddmmyyyy;
}

function formatDdMmYyyy(date) {
  const d = date.getDate().toString().padStart(2, "0");
  const m = (date.getMonth() + 1).toString().padStart(2, "0");
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
}

async function fetchFrankfurterRateToHkd(currency, ddmmyyyy) {
  const isoDate = toIsoDate(ddmmyyyy);
  const url = `${FRANKFURTER_BASE}/${isoDate}`;
  const { data } = await axios.get(url, {
    params: { base: currency, symbols: "HKD" },
    timeout: 15000,
  });
  const rate = data?.rates?.HKD;
  if (!Number.isFinite(rate) || rate <= 0) {
    throw new Error(`Frankfurter returned no usable HKD rate for ${currency} on ${ddmmyyyy}`);
  }
  return rate;
}

/**
 * Parses the CBR daily XML feed with a small regex scan (no XML dependency needed for
 * this fixed, well-known structure) and returns HKD per 1 RUB for that date.
 */
async function fetchCbrRateToHkd(ddmmyyyy) {
  const { data } = await axios.get(CBR_DAILY_URL, {
    params: { date_req: ddmmyyyy },
    timeout: 15000,
    responseType: "text",
    // CBR serves windows-1251 by default but also accepts/returns UTF-8 reasonably for digits/tags we need.
  });
  const xml = typeof data === "string" ? data : String(data);

  const valuteBlocks = xml.match(/<Valute[^>]*>[\s\S]*?<\/Valute>/g) || [];
  const hkdBlock = valuteBlocks.find((b) => /<CharCode>\s*HKD\s*<\/CharCode>/i.test(b));
  if (!hkdBlock) {
    throw new Error(`CBR daily feed has no HKD entry for ${ddmmyyyy}`);
  }
  const nominalMatch = hkdBlock.match(/<Nominal>\s*(\d+)\s*<\/Nominal>/i);
  const valueMatch = hkdBlock.match(/<Value>\s*([\d.,]+)\s*<\/Value>/i);
  if (!nominalMatch || !valueMatch) {
    throw new Error(`CBR daily feed HKD entry could not be parsed for ${ddmmyyyy}`);
  }
  const nominal = Number(nominalMatch[1]);
  const rubPerNominalHkd = Number(valueMatch[1].replace(",", "."));
  if (!Number.isFinite(nominal) || nominal <= 0 || !Number.isFinite(rubPerNominalHkd) || rubPerNominalHkd <= 0) {
    throw new Error(`CBR daily feed HKD entry had unusable numbers for ${ddmmyyyy}`);
  }
  const rubPerHkd = rubPerNominalHkd / nominal;
  // We want HKD per 1 RUB (RUB -> HKD conversion), the inverse of RUB per 1 HKD.
  return 1 / rubPerHkd;
}

/**
 * Returns HKD per 1 unit of `currency` on `ddmmyyyy` (dd/mm/yyyy). Cached per process —
 * a given (currency, date) pair only ever hits the network once. `onNewRate`, if given,
 * fires exactly once per newly-fetched (currency, date) pair — used by callers to append
 * an audit-log row without duplicating work here.
 */
async function getRateToHkd(currency, ddmmyyyy, onNewRate) {
  const cur = String(currency || "").trim().toUpperCase();
  if (!cur) throw new Error("getRateToHkd: currency is required");
  if (cur === "HKD") return 1;

  const safeDate = clampToToday(ddmmyyyy);
  const key = cacheKey(cur, safeDate);

  if (rateCache.has(key)) {
    return rateCache.get(key);
  }
  if (inflight.has(key)) {
    return inflight.get(key);
  }

  const promise = (async () => {
    const rate = cur === "RUB" ? await fetchCbrRateToHkd(safeDate) : await fetchFrankfurterRateToHkd(cur, safeDate);
    rateCache.set(key, rate);
    if (typeof onNewRate === "function") {
      try {
        await onNewRate({ currency: cur, date: safeDate, rate });
      } catch (err) {
        // Audit-log failures must never block the actual expense write.
        console.error(`fxRates onNewRate callback failed for ${cur} ${safeDate}:`, err.message);
      }
    }
    return rate;
  })();

  inflight.set(key, promise);
  try {
    return await promise;
  } finally {
    inflight.delete(key);
  }
}

/** Seeds the cache without hitting the network — used to hydrate from Excel's ExchangeRateHistory at boot. */
function primeCache(currency, ddmmyyyy, rate) {
  if (!Number.isFinite(rate) || rate <= 0) return;
  rateCache.set(cacheKey(currency, ddmmyyyy), rate);
}

module.exports = {
  TRACKED_CURRENCIES,
  getRateToHkd,
  primeCache,
  clampToToday,
};
