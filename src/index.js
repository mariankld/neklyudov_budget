require("dotenv").config();
const express = require("express");
const axios = require("axios");
const OpenAI = require("openai");
const { google } = require("googleapis");

const app = express();
app.use(express.json());

const PORT = Number(process.env.PORT || 3000);

const INCOME_SHEET = (process.env.INCOME_SHEET_NAME || "Income").trim();
const RAW_SHEET = (process.env.RAW_SHEET_NAME || "Expenses_RAW").trim();
/** Dashboard tab: never used as a category target or written by the logger. */
const SUMMARY_TAB = (process.env.GOOGLE_SUMMARY_TAB || "Summary").trim();
const DEFAULT_CURRENCY = (process.env.DEFAULT_EXPENSE_CURRENCY || "HKD").trim();

/** Filled at startup from the spreadsheet (tab titles minus skip list). */
let allowedCategories = [];

/** Lowercase alias → exact tab title; from CATEGORY_MAPPING JSON in .env */
let categoryAliasMap = new Map();

const TELEGRAM_API_BASE = `https://api.telegram.org/bot${process.env.TELEGRAM_BOT_TOKEN}`;

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

function getIsoParts(date) {
  const isoDate = date.toISOString().slice(0, 10);
  const time = date.toTimeString().slice(0, 8);
  return { isoDate, time };
}

function escapeGoogleSheetRangeTitle(sheetName) {
  return `'${sheetName.replace(/'/g, "''")}'`;
}

function createGoogleSheetsClient() {
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_OAUTH_REDIRECT_URI || "http://127.0.0.1:3001/oauth/callback"
  );
  oauth2Client.setCredentials({
    refresh_token: process.env.GOOGLE_REFRESH_TOKEN,
  });
  return google.sheets({ version: "v4", auth: oauth2Client });
}

function parseSkipTabNames() {
  const extra = (process.env.GOOGLE_SKIP_TABS || "")
    .split(/[|,\n]/)
    .map((s) => s.trim())
    .filter(Boolean);
  return new Set(extra);
}

function parseCategoryMappingEnv() {
  const raw = (process.env.CATEGORY_MAPPING || "").trim();
  if (!raw) return new Map();
  try {
    const obj = JSON.parse(raw);
    if (!obj || typeof obj !== "object" || Array.isArray(obj)) return new Map();
    const map = new Map();
    for (const [key, val] of Object.entries(obj)) {
      if (typeof key !== "string" || typeof val !== "string") continue;
      const k = key.trim().toLowerCase();
      const v = val.trim();
      if (k && v) map.set(k, v);
    }
    return map;
  } catch {
    return new Map();
  }
}

/**
 * Map model output to an allowed sheet tab: exact match, alias (CATEGORY_MAPPING), or case-insensitive tab match.
 */
function resolveCategoryToSheet(name, categories) {
  const trimmed = String(name || "").trim();
  if (!trimmed) {
    throw new Error("Empty category");
  }
  if (categories.includes(trimmed)) {
    return trimmed;
  }
  const aliasTarget = categoryAliasMap.get(trimmed.toLowerCase());
  if (aliasTarget && categories.includes(aliasTarget)) {
    return aliasTarget;
  }
  const lower = trimmed.toLowerCase();
  const ci = categories.find((c) => c.toLowerCase() === lower);
  if (ci) {
    return ci;
  }
  throw new Error(`Unsupported category: ${trimmed}`);
}

function formatDateDdMmYyyy(date) {
  const d = date.getDate().toString().padStart(2, "0");
  const m = (date.getMonth() + 1).toString().padStart(2, "0");
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
}

async function fetchSpreadsheetTabTitles(sheets, spreadsheetId) {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "sheets(properties(title))",
  });
  const sheetsList = meta.data.sheets || [];
  return sheetsList
    .map((s) => s.properties && s.properties.title)
    .filter(Boolean);
}

/**
 * Tabs the model may choose as the expense/income *category* (real sheet to log under).
 * RAW is omitted on purpose: it is not a category—every transaction is always appended to
 * RAW_SHEET separately as the full audit trail (see webhook). Summary is the dashboard.
 */
function buildAllowedCategoriesFromTabs(allTitles) {
  const skip = parseSkipTabNames();
  skip.add(RAW_SHEET);
  if (SUMMARY_TAB) skip.add(SUMMARY_TAB);
  return allTitles.filter((t) => !skip.has(t));
}

async function logFirstRow(sheets, spreadsheetId, sheetLabel, sheetName) {
  try {
    const range = `${escapeGoogleSheetRangeTitle(sheetName)}!1:1`;
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });
    const row = res.data.values && res.data.values[0];
    console.log(
      `[structure] ${sheetLabel} "${sheetName}" row1:`,
      row && row.length ? row.join(" | ") : "(empty)"
    );
  } catch (err) {
    console.warn(`[structure] Could not read row1 of "${sheetName}": ${err.message}`);
  }
}

async function appendRowToGoogleSheet(sheets, spreadsheetId, sheetName, rowValues) {
  const range = `${escapeGoogleSheetRangeTitle(sheetName)}!A:Z`;
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [rowValues],
    },
  });
}

function buildPrompt(messageText) {
  const categoryList = allowedCategories.join(", ");
  return [
    `Extract from this message: (1) numeric amount, (2) description, (3) category (use the exact sheet name from: ${categoryList}, or a close synonym—synonyms are mapped automatically), and (4) type (expense or income).`,
    `Never use "${RAW_SHEET}" or any universal/log tab as category—the app logs every message there automatically with full detail.`,
    'Return ONLY a JSON object: {"amount": number, "description": string, "category": string, "type": string}.',
    'Example: {"amount": 250, "description": "groceries", "category": "Shopping", "type": "expense"}',
    "",
    `Message: ${messageText}`,
  ].join("\n");
}

async function sendTelegramMessage(chatId, text, replyToMessageId) {
  await axios.post(`${TELEGRAM_API_BASE}/sendMessage`, {
    chat_id: chatId,
    text,
    reply_to_message_id: replyToMessageId,
  });
}

async function parseWithOpenAI(messageText) {
  const today = new Date().toISOString().slice(0, 10);
  const completion = await openai.responses.create({
    model: process.env.OPENAI_MODEL || "gpt-5.5-medium",
    input: [
      {
        role: "system",
        content: `You are a financial parsing assistant. Extract expense/income amounts and determine the most appropriate category. For incoming payments/transfers, set type to 'income' and category to '${INCOME_SHEET}'. For outgoing expenses, set type to 'expense' and use one of the expense category tabs provided. Today's date is ${today}. Output strict JSON only.`,
      },
      {
        role: "user",
        content: buildPrompt(messageText),
      },
    ],
    text: {
      format: {
        type: "json_schema",
        name: "transaction",
        schema: {
          type: "object",
          properties: {
            amount: { type: "number" },
            description: { type: "string" },
            category: { type: "string" },
            type: { type: "string", enum: ["expense", "income"] },
          },
          required: ["amount", "description", "category", "type"],
          additionalProperties: false,
        },
      },
    },
  });

  const raw = completion.output_text;
  const parsed = JSON.parse(raw);
  return parsed;
}

function normalizeResult(result) {
  const normalized = {
    amount: Number(result.amount),
    description: String(result.description || "").trim(),
    category: String(result.category || "").trim(),
    type: String(result.type || "").trim().toLowerCase(),
  };

  if (!Number.isFinite(normalized.amount) || normalized.amount <= 0) {
    throw new Error("No valid amount extracted");
  }

  if (normalized.type === "income") {
    normalized.category = INCOME_SHEET;
  } else if (normalized.type !== "expense") {
    throw new Error(`Unsupported type: ${normalized.type}`);
  } else {
    normalized.category = resolveCategoryToSheet(normalized.category, allowedCategories);
  }

  return normalized;
}

function requireEnv(name) {
  const v = process.env[name];
  if (!v || !String(v).trim()) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
}

/**
 * Telegram only POSTs updates if setWebhook points at this server. Long polling is not used.
 * On Railway, RAILWAY_PUBLIC_DOMAIN is set automatically; override with TELEGRAM_WEBHOOK_BASE_URL.
 */
async function ensureTelegramWebhook() {
  if (!process.env.TELEGRAM_BOT_TOKEN || !String(process.env.TELEGRAM_BOT_TOKEN).trim()) {
    console.warn("TELEGRAM_BOT_TOKEN is missing.");
    return;
  }

  const explicit = (process.env.TELEGRAM_WEBHOOK_BASE_URL || "").trim().replace(/\/$/, "");
  const staticUrl = (process.env.RAILWAY_STATIC_URL || "").trim().replace(/\/$/, "");
  const railwayHost = (process.env.RAILWAY_PUBLIC_DOMAIN || "").trim();
  const railwayFromHost = railwayHost
    ? `https://${railwayHost.replace(/^https?:\/\//, "")}`
    : "";
  const publicUrl = (process.env.PUBLIC_URL || "").trim().replace(/\/$/, "");

  const base = explicit || staticUrl || railwayFromHost || publicUrl;
  if (!base) {
    console.warn(
      "Telegram webhook not set automatically (no TELEGRAM_WEBHOOK_BASE_URL, RAILWAY_PUBLIC_DOMAIN, RAILWAY_STATIC_URL, or PUBLIC_URL). Register manually: POST https://api.telegram.org/bot<token>/setWebhook with url=https://<host>/webhook/telegram"
    );
    return;
  }

  const webhookUrl = `${base}/webhook/telegram`;
  try {
    const setRes = await axios.post(`${TELEGRAM_API_BASE}/setWebhook`, { url: webhookUrl });
    if (!setRes.data?.ok) {
      console.error("Telegram setWebhook failed:", setRes.data);
      return;
    }
    const infoRes = await axios.get(`${TELEGRAM_API_BASE}/getWebhookInfo`);
    const info = infoRes.data?.result;
    console.log(
      `Telegram webhook → ${info?.url || webhookUrl} (pending updates: ${info?.pending_update_count ?? 0})`
    );
  } catch (err) {
    console.error("Telegram setWebhook error:", err.response?.data || err.message);
  }
}

async function bootstrap() {
  requireEnv("TELEGRAM_BOT_TOKEN");
  requireEnv("OPENAI_API_KEY");
  requireEnv("GOOGLE_SPREADSHEET_ID");
  requireEnv("GOOGLE_CLIENT_ID");
  requireEnv("GOOGLE_CLIENT_SECRET");
  requireEnv("GOOGLE_REFRESH_TOKEN");

  const sheets = createGoogleSheetsClient();
  const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;

  const tabTitles = await fetchSpreadsheetTabTitles(sheets, spreadsheetId);
  if (!tabTitles.length) {
    throw new Error("Spreadsheet has no sheet tabs.");
  }

  if (!tabTitles.includes(RAW_SHEET)) {
    throw new Error(
      `RAW tab "${RAW_SHEET}" not found. Available tabs: ${tabTitles.join(", ")}. Set RAW_SHEET_NAME in .env.`
    );
  }

  if (!tabTitles.includes(INCOME_SHEET)) {
    throw new Error(
      `Income tab "${INCOME_SHEET}" not found. Available tabs: ${tabTitles.join(", ")}. Set INCOME_SHEET_NAME in .env.`
    );
  }

  categoryAliasMap = parseCategoryMappingEnv();
  allowedCategories = buildAllowedCategoriesFromTabs(tabTitles);

  if (!allowedCategories.length) {
    throw new Error(
      "No category tabs after exclusions. Check GOOGLE_SKIP_TABS and RAW_SHEET_NAME."
    );
  }

  if (!allowedCategories.includes(INCOME_SHEET)) {
    throw new Error(
      `Income tab "${INCOME_SHEET}" was excluded. Remove it from GOOGLE_SKIP_TABS if listed there.`
    );
  }

  console.log(
    `Budget logger: ${allowedCategories.length} category targets (sheet tabs the model may choose) — ${allowedCategories.join(", ")}`
  );
  console.log(
    `Universal RAW log (every transaction is appended here; not a category option): "${RAW_SHEET}"`
  );
  if (categoryAliasMap.size) {
    console.log(`Category aliases loaded: ${categoryAliasMap.size}`);
  }

  await logFirstRow(sheets, spreadsheetId, "RAW log", RAW_SHEET);
  const sampleCategoryTab = allowedCategories.find((t) => t !== INCOME_SHEET) || INCOME_SHEET;
  await logFirstRow(sheets, spreadsheetId, "Sample category", sampleCategoryTab);
}

app.post("/webhook/telegram", async (req, res) => {
  res.status(200).json({ ok: true });

  const message = req.body?.message;
  const chatId = message?.chat?.id;
  const originalMessage = message?.text;
  const sender =
    message?.from?.first_name ||
    message?.from?.username ||
    `${message?.from?.id || "Unknown"}`;
  const messageId = message?.message_id;

  if (!chatId || !originalMessage) {
    return;
  }

  try {
    await sendTelegramMessage(chatId, "⏳ Logging your transaction...", messageId);

    const aiResult = normalizeResult(await parseWithOpenAI(originalMessage));
    const now = new Date();
    const { isoDate, time } = getIsoParts(now);

    const categorySheetName =
      aiResult.type === "income" ? INCOME_SHEET : aiResult.category;

    const categoryRow = [
      isoDate,
      time,
      aiResult.amount,
      aiResult.type,
      categorySheetName,
      aiResult.description,
    ];

    const rawRow = [
      formatDateDdMmYyyy(now),
      aiResult.type === "income" ? "Income" : "Expense",
      categorySheetName,
      "",
      aiResult.description,
      "",
      aiResult.amount,
      DEFAULT_CURRENCY,
      aiResult.amount,
      `Telegram — ${sender}: ${originalMessage}`,
    ];

    const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
    const sheets = createGoogleSheetsClient();
    await appendRowToGoogleSheet(sheets, spreadsheetId, RAW_SHEET, rawRow);
    await appendRowToGoogleSheet(sheets, spreadsheetId, categorySheetName, categoryRow);

    await sendTelegramMessage(
      chatId,
      `✅ Logged: ${aiResult.amount} ${aiResult.type} under ${categorySheetName} on ${isoDate}`,
      messageId
    );
  } catch (error) {
    try {
      await sendTelegramMessage(
        chatId,
        "❌ Error saving to the spreadsheet. Please check access and try again.",
        messageId
      );
    } catch (notifyError) {
      console.error("Failed to send Telegram error:", notifyError.message);
    }

    console.error("Workflow error:", error.message);
  }
});

app.get("/health", (_req, res) => {
  res.status(200).json({ status: "ok" });
});

bootstrap()
  .then(() => {
    app.listen(PORT, async () => {
      console.log(`Budget logger is running on port ${PORT}`);
      await ensureTelegramWebhook();
    });
  })
  .catch((err) => {
    console.error("Startup failed:", err.message);
    process.exit(1);
  });
