require("dotenv").config();
const crypto = require("crypto");
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
const DEFAULT_LOCATION = (process.env.DEFAULT_EXPENSE_LOCATION || "Hong Kong").trim();
const FAMILY_GREETING_NAME = (process.env.BUDGET_FAMILY_NAME || "Neklyudov").trim();

/** Filled at startup from the spreadsheet (tab titles minus skip list). */
let allowedCategories = [];

const PENDING_TTL_MS = 24 * 60 * 60 * 1000;

/** token -> { draft, originalMessage, sender, sourceMessageId, chatId, createdAt } */
const pendingByToken = new Map();

/** Successfully logged tokens (idempotency); same TTL as pending. token -> completedAt ms */
const completedLogByToken = new Map();

/** Tokens currently being written (guards parallel webhook deliveries / double-tap races). */
const inflightLogTokens = new Set();

/** chatId -> { token } — user clicked Edit and should reply with changes */
const awaitingEditByChat = new Map();

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

function cellsRoughlyEqual(expected, actual) {
  if (expected === actual) return true;
  const ex = expected === null || expected === undefined ? "" : String(expected).trim();
  const ac = actual === null || actual === undefined ? "" : String(actual).trim();
  if (ex === ac) return true;
  const nEx = Number(ex.replace(",", "."));
  const nAc = Number(String(ac).replace(",", "."));
  if (Number.isFinite(nEx) && Number.isFinite(nAc) && Math.abs(nEx - nAc) < 1e-9) {
    return true;
  }
  return false;
}

/**
 * Triple-check: (1) append API says exactly one row written, (2) read-back row exists,
 * (3) critical fields match (amount + category), tolerating Sheets date/number display formatting.
 */
async function appendRowAndVerify(
  sheets,
  spreadsheetId,
  sheetName,
  rowValues,
  label,
  criticalChecks
) {
  const range = `${escapeGoogleSheetRangeTitle(sheetName)}!A:Z`;
  const appendRes = await sheets.spreadsheets.values.append({
    spreadsheetId,
    range,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [rowValues],
    },
  });

  const updates = appendRes.data?.updates;
  const updatedRows = updates?.updatedRows;
  const updatedRange = updates?.updatedRange;
  if (updatedRows !== 1 || !updatedRange) {
    throw new Error(
      `${label}: Sheets append did not report a single new row (updatedRows=${updatedRows}).`
    );
  }

  const readRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: updatedRange,
    valueRenderOption: "UNFORMATTED_VALUE",
    dateTimeRenderOption: "SERIAL_NUMBER",
  });
  const gotRow = readRes.data?.values?.[0];
  if (!gotRow || !gotRow.length) {
    throw new Error(`${label}: Read-back after append returned empty.`);
  }

  for (const { index, name, expected } of criticalChecks) {
    const act = index < gotRow.length ? gotRow[index] : "";
    if (!cellsRoughlyEqual(expected, act)) {
      throw new Error(
        `${label}: Verification failed for ${name} (expected "${expected}", got "${act}").`
      );
    }
  }

  return appendRes;
}

const transactionJsonSchema = {
  type: "object",
  properties: {
    amount: { type: "number" },
    description: { type: "string" },
    category: { type: "string" },
    subcategory: { type: "string" },
    type: { type: "string", enum: ["expense", "income"] },
    location: { type: "string" },
    currency: { type: "string" },
    notes: { type: "string" },
  },
  required: ["amount", "description", "category", "subcategory", "type", "location", "currency", "notes"],
  additionalProperties: false,
};

function buildAnalysisPrompt(messageText, contextDateLabel) {
  const categoryList = allowedCategories.join(", ");
  return [
    `The user's message date context: ${contextDateLabel}.`,
    `Available category tabs (pick exactly one sheet name or a close synonym; synonyms map via env): ${categoryList}.`,
    `Never use "${RAW_SHEET}" as category—that tab receives every row automatically.`,
    `For salary, incoming transfers, or money received, use type "income" and category "${INCOME_SHEET}".`,
    "",
    "Fill fields as follows:",
    "- amount: primary numeric sum from the message.",
    "- description: short label; you may use the user's wording.",
    `- subcategory: specific line item (e.g. Transport → Highway toll).`,
    `- location: default "${DEFAULT_LOCATION}" unless the message states another place.`,
    `- currency: default "${DEFAULT_CURRENCY}" unless the message states another currency.`,
    `- notes: empty string unless the user explicitly adds an extra note beyond amount/description; do not duplicate the description.`,
    "",
    `User message:\n${messageText}`,
  ].join("\n");
}

function buildEditPrompt(currentDraft, editInstruction, originalMessage) {
  const categoryList = allowedCategories.join(", ");
  return [
    "Apply the user's edit instructions to this draft. Keep other fields unchanged unless the edit implies them.",
    `Allowed categories: ${categoryList}. Never use "${RAW_SHEET}" as category.`,
    "",
    "Current draft (JSON):",
    JSON.stringify(currentDraft),
    "",
    "Original user message:",
    originalMessage,
    "",
    "Edit instruction:",
    editInstruction,
  ].join("\n");
}

async function sendTelegramMessage(chatId, text, replyToMessageId, replyMarkup) {
  const body = {
    chat_id: chatId,
    text,
  };
  if (replyToMessageId != null) {
    body.reply_to_message_id = replyToMessageId;
  }
  if (replyMarkup) {
    body.reply_markup = replyMarkup;
  }
  await axios.post(`${TELEGRAM_API_BASE}/sendMessage`, body);
}

async function answerCallbackQuery(callbackQueryId, text, showAlert) {
  await axios.post(`${TELEGRAM_API_BASE}/answerCallbackQuery`, {
    callback_query_id: callbackQueryId,
    text: text || undefined,
    show_alert: Boolean(showAlert),
  });
}

async function editTelegramMessage(chatId, messageId, text, replyMarkup) {
  const body = {
    chat_id: chatId,
    message_id: messageId,
    text,
  };
  if (replyMarkup !== undefined) {
    body.reply_markup = replyMarkup;
  }
  await axios.post(`${TELEGRAM_API_BASE}/editMessageText`, body);
}

function formatContextDateLabel(messageDate) {
  const d = messageDate instanceof Date ? messageDate : new Date();
  return d.toISOString().slice(0, 10);
}

async function parseTransactionWithOpenAI(messageText, messageDate) {
  const contextDateLabel = formatContextDateLabel(messageDate);
  const completion = await openai.responses.create({
    model: process.env.OPENAI_MODEL || "gpt-4o-mini",
    input: [
      {
        role: "system",
        content: `You analyze Telegram budget messages for a family spreadsheet. Output strict JSON only. Map spending to the closest category tab. Today for context is ${contextDateLabel}.`,
      },
      {
        role: "user",
        content: buildAnalysisPrompt(messageText, contextDateLabel),
      },
    ],
    text: {
      format: {
        type: "json_schema",
        name: "transaction",
        schema: transactionJsonSchema,
      },
    },
  });

  const raw = completion.output_text;
  return JSON.parse(raw);
}

async function parseEditWithOpenAI(currentDraft, editInstruction, originalMessage) {
  const completion = await openai.responses.create({
    model: process.env.OPENAI_MODEL || "gpt-4o-mini",
    input: [
      {
        role: "system",
        content:
          "You update a budget transaction draft from user edit instructions. Preserve unspecified fields. Output strict JSON only.",
      },
      {
        role: "user",
        content: buildEditPrompt(currentDraft, editInstruction, originalMessage),
      },
    ],
    text: {
      format: {
        type: "json_schema",
        name: "transaction_edit",
        schema: transactionJsonSchema,
      },
    },
  });

  const raw = completion.output_text;
  return JSON.parse(raw);
}

function normalizeDraft(result, messageDate) {
  const md = messageDate instanceof Date ? messageDate : new Date();
  const normalized = {
    amount: Number(result.amount),
    description: String(result.description || "").trim(),
    category: String(result.category || "").trim(),
    subcategory: String(result.subcategory || "").trim(),
    type: String(result.type || "").trim().toLowerCase(),
    location: String(result.location || "").trim() || DEFAULT_LOCATION,
    currency: String(result.currency || "").trim() || DEFAULT_CURRENCY,
    notes: String(result.notes || "").trim(),
    /** dd/mm/yyyy for RAW sheet (from message time) */
    dateDdMmYyyy: formatDateDdMmYyyy(md),
  };

  if (!Number.isFinite(normalized.amount) || normalized.amount <= 0) {
    throw new Error("No valid amount found. Send a positive number with a short description.");
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

function prunePending() {
  const now = Date.now();
  for (const [token, entry] of pendingByToken.entries()) {
    if (now - entry.createdAt > PENDING_TTL_MS) {
      pendingByToken.delete(token);
    }
  }
  for (const [token, completedAt] of completedLogByToken.entries()) {
    if (now - completedAt > PENDING_TTL_MS) {
      completedLogByToken.delete(token);
    }
  }
}

function newPendingToken() {
  return crypto.randomBytes(8).toString("hex");
}

function formatDraftPreview(draft, categorySheetName) {
  const typeLabel = draft.type === "income" ? "Income" : "Expense";
  const notesLine =
    draft.notes && draft.notes.length
      ? `\nNotes: ${draft.notes}`
      : "\nNotes: —";

  return [
    "I will log this as:",
    "",
    `Date: ${draft.dateDdMmYyyy}`,
    `Type: ${typeLabel}`,
    `Category: ${categorySheetName} (sheet tab)`,
    `Subcategory: ${draft.subcategory || "—"}`,
    `Description: ${draft.description}`,
    `Location: ${draft.location}`,
    `Sum: ${draft.amount}`,
    `Currency: ${draft.currency}`,
    "Sum (HKD): (leave blank — calculated in sheet)",
    notesLine,
    "",
    "Does this look good?",
  ].join("\n");
}

function categoryTabRowDescription(draft) {
  if (draft.subcategory) {
    return `${draft.subcategory}: ${draft.description}`;
  }
  return draft.description;
}

function buildRawRowValues(draft, categorySheetName) {
  const notesCell = draft.notes && draft.notes.length ? draft.notes : "";

  return [
    draft.dateDdMmYyyy,
    draft.type === "income" ? "Income" : "Expense",
    categorySheetName,
    draft.subcategory,
    draft.description,
    draft.location,
    draft.amount,
    draft.currency,
    "",
    notesCell,
  ];
}

function buildCategoryRowValues(draft, categorySheetName, messageDate) {
  const md = messageDate instanceof Date ? messageDate : new Date();
  const { isoDate, time } = getIsoParts(md);
  return [
    isoDate,
    time,
    draft.amount,
    draft.type,
    categorySheetName,
    categoryTabRowDescription(draft),
  ];
}

async function appendTransactionToSheets(draft, categorySheetName, messageDate) {
  const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
  const sheets = createGoogleSheetsClient();
  const rawRow = buildRawRowValues(draft, categorySheetName);
  const categoryRow = buildCategoryRowValues(draft, categorySheetName, messageDate);

  await appendRowAndVerify(sheets, spreadsheetId, RAW_SHEET, rawRow, `RAW "${RAW_SHEET}"`, [
    { index: 1, name: "type", expected: draft.type === "income" ? "Income" : "Expense" },
    { index: 2, name: "category tab", expected: categorySheetName },
    { index: 6, name: "amount", expected: draft.amount },
  ]);
  await appendRowAndVerify(
    sheets,
    spreadsheetId,
    categorySheetName,
    categoryRow,
    `Category "${categorySheetName}"`,
    [
      { index: 2, name: "amount", expected: draft.amount },
      { index: 4, name: "category tab", expected: categorySheetName },
    ]
  );
}

function isStartCommand(text) {
  if (!text || typeof text !== "string") return false;
  const t = text.trim();
  return t === "/start" || t.startsWith("/start ");
}

function getWelcomeMessage() {
  return [
    `👋 Hello, ${FAMILY_GREETING_NAME} family! The Budget Bot is now active and ready to track your expenses!`,
    "",
    "📝 To log an expense, simply send a message like:",
    '• "Spent 500 on groceries"',
    '• "250 taxi"',
    '• "1200 dinner at restaurant"',
    '• "Купил продукты на 800"',
    "",
    "I'll analyze the amount and description, show a preview (category tab, RAW fields), and wait for you to tap Yes, log it or Edit before anything is saved. ✅",
    "",
    "If you're editing and change your mind, send /cancel.",
  ].join("\n");
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

function isCancelCommand(text) {
  const t = (text || "").trim();
  return t === "/cancel" || t.startsWith("/cancel ");
}

async function processExpenseFlow(chatId, originalMessage, sender, messageId, messageDate) {
  awaitingEditByChat.delete(chatId);
  await sendTelegramMessage(chatId, "⏳ Analyzing your message...", messageId);

  const parsed = await parseTransactionWithOpenAI(originalMessage, messageDate);
  const draft = normalizeDraft(parsed, messageDate);
  const categorySheetName = draft.category;

  prunePending();
  const token = newPendingToken();
  pendingByToken.set(token, {
    draft,
    originalMessage,
    sender,
    sourceMessageId: messageId,
    chatId,
    createdAt: Date.now(),
    messageDate,
  });

  const preview = formatDraftPreview(draft, categorySheetName);
  const keyboard = {
    inline_keyboard: [
      [
        { text: "Yes, log it", callback_data: `log:${token}` },
        { text: "Edit", callback_data: `edit:${token}` },
      ],
    ],
  };

  await sendTelegramMessage(chatId, preview, messageId, keyboard);
}

async function processEditInstruction(chatId, editText, messageId) {
  const wait = awaitingEditByChat.get(chatId);
  if (!wait) {
    return false;
  }

  const entry = pendingByToken.get(wait.token);
  if (!entry) {
    awaitingEditByChat.delete(chatId);
    await sendTelegramMessage(
      chatId,
      "That edit session expired. Send the expense again.",
      messageId
    );
    return true;
  }

  prunePending();

  const flatDraft = {
    amount: entry.draft.amount,
    description: entry.draft.description,
    category: entry.draft.category,
    subcategory: entry.draft.subcategory,
    type: entry.draft.type,
    location: entry.draft.location,
    currency: entry.draft.currency,
    notes: entry.draft.notes,
  };

  try {
    await sendTelegramMessage(chatId, "⏳ Updating your draft...", messageId);
    const parsed = await parseEditWithOpenAI(flatDraft, editText, entry.originalMessage);
    const draft = normalizeDraft(parsed, entry.messageDate);
    const categorySheetName = draft.category;

    pendingByToken.delete(wait.token);
    const newToken = newPendingToken();
    pendingByToken.set(newToken, {
      draft,
      originalMessage: entry.originalMessage,
      sender: entry.sender,
      sourceMessageId: entry.sourceMessageId,
      chatId,
      createdAt: Date.now(),
      messageDate: entry.messageDate,
    });
    awaitingEditByChat.delete(chatId);

    const preview = formatDraftPreview(draft, categorySheetName);
    const keyboard = {
      inline_keyboard: [
        [
          { text: "Yes, log it", callback_data: `log:${newToken}` },
          { text: "Edit", callback_data: `edit:${newToken}` },
        ],
      ],
    };

    await sendTelegramMessage(chatId, preview, entry.sourceMessageId, keyboard);
  } catch (err) {
    await sendTelegramMessage(
      chatId,
      `❌ Could not apply changes: ${err.message}. Try again or send /cancel.`,
      messageId
    );
  }
  return true;
}

async function handleCallbackQuery(callbackQuery) {
  const data = callbackQuery.data || "";
  const chatId = callbackQuery.message?.chat?.id;
  const messageId = callbackQuery.message?.message_id;
  const callbackQueryId = callbackQuery.id;

  if (!chatId || !messageId) {
    await answerCallbackQuery(callbackQueryId, "Missing chat message.", true);
    return;
  }

  const logPrefix = "log:";
  const editPrefix = "edit:";

  if (data.startsWith(logPrefix)) {
    const token = data.slice(logPrefix.length);
    prunePending();

    if (completedLogByToken.has(token)) {
      await answerCallbackQuery(
        callbackQueryId,
        "Already logged — this entry was saved. Check RAW and the category tab.",
        true
      );
      return;
    }

    const entry = pendingByToken.get(token);
    if (!entry) {
      if (inflightLogTokens.has(token)) {
        await answerCallbackQuery(
          callbackQueryId,
          "Already saving this expense — please wait for the result.",
          true
        );
        return;
      }
      await answerCallbackQuery(callbackQueryId, "This confirmation expired. Send the expense again.", true);
      return;
    }
    if (entry.chatId !== chatId) {
      await answerCallbackQuery(callbackQueryId, "Not allowed.", true);
      return;
    }

    if (!pendingByToken.delete(token)) {
      await answerCallbackQuery(
        callbackQueryId,
        "Already saving this expense — please wait for the result.",
        true
      );
      return;
    }
    inflightLogTokens.add(token);

    try {
      await appendTransactionToSheets(entry.draft, entry.draft.category, entry.messageDate);
      completedLogByToken.set(token, Date.now());

      const cat = entry.draft.category;
      const dLabel = entry.draft.dateDdMmYyyy;
      const successText =
        `✅ Logged successfully.\n` +
        `Verified in Google Sheets: "${RAW_SHEET}" and "${cat}".\n` +
        `${entry.draft.amount} ${entry.draft.currency} · ${dLabel}`;

      try {
        await answerCallbackQuery(callbackQueryId, "Logged successfully.");
      } catch (_) {
        /* callback may already be answered on duplicate delivery */
      }
      try {
        await editTelegramMessage(chatId, messageId, successText, { inline_keyboard: [] });
      } catch {
        await sendTelegramMessage(chatId, successText, null);
      }
      awaitingEditByChat.delete(chatId);
    } catch (err) {
      console.error("Log callback error:", err.message);
      pendingByToken.set(token, entry);

      const failDetail = (err.message || "Unknown error").slice(0, 400);
      try {
        await answerCallbackQuery(
          callbackQueryId,
          `Log failed: ${failDetail}`.slice(0, 190),
          true
        );
      } catch (_) {
        /* ignore */
      }
      try {
        await sendTelegramMessage(
          chatId,
          `❌ Log failed — nothing new was saved to the sheet.\n${failDetail}`,
          null
        );
      } catch (notifyError) {
        console.error("Failed to send Telegram error:", notifyError.message);
      }
    } finally {
      inflightLogTokens.delete(token);
    }
    return;
  }

  if (data.startsWith(editPrefix)) {
    const token = data.slice(editPrefix.length);
    prunePending();
    const entry = pendingByToken.get(token);
    if (!entry) {
      await answerCallbackQuery(callbackQueryId, "This confirmation expired. Send the expense again.", true);
      return;
    }
    if (entry.chatId !== chatId) {
      await answerCallbackQuery(callbackQueryId, "Not allowed.", true);
      return;
    }

    awaitingEditByChat.set(chatId, { token });
    await answerCallbackQuery(callbackQueryId);
    await sendTelegramMessage(
      chatId,
      "✏️ What would you like to change? For example: category to Shopping, amount 120, currency USD, or location Tokyo.",
      messageId
    );
    return;
  }

  await answerCallbackQuery(callbackQueryId);
}

app.post("/webhook/telegram", async (req, res) => {
  const ok = () => res.status(200).json({ ok: true });

  try {
    const callbackQuery = req.body?.callback_query;
    if (callbackQuery) {
      await handleCallbackQuery(callbackQuery);
      return ok();
    }

    const message = req.body?.message;
    const chatId = message?.chat?.id;
    const originalMessage = message?.text;
    const sender =
      message?.from?.first_name ||
      message?.from?.username ||
      `${message?.from?.id || "Unknown"}`;
    const messageId = message?.message_id;

    if (!chatId || !originalMessage) {
      return ok();
    }

    const messageDate = new Date((message.date || Math.floor(Date.now() / 1000)) * 1000);

    if (isStartCommand(originalMessage)) {
      awaitingEditByChat.delete(chatId);
      try {
        await sendTelegramMessage(chatId, getWelcomeMessage(), messageId);
      } catch (err) {
        console.error("Telegram send error:", err.message);
      }
      return ok();
    }

    if (isCancelCommand(originalMessage)) {
      awaitingEditByChat.delete(chatId);
      try {
        await sendTelegramMessage(
          chatId,
          "Edit cancelled. Send an expense when you're ready.",
          messageId
        );
      } catch (err) {
        console.error("Telegram send error:", err.message);
      }
      return ok();
    }

    try {
      if (awaitingEditByChat.has(chatId)) {
        const handled = await processEditInstruction(chatId, originalMessage.trim(), messageId);
        if (handled) {
          return ok();
        }
      }

      await processExpenseFlow(chatId, originalMessage, sender, messageId, messageDate);
    } catch (error) {
      try {
        await sendTelegramMessage(
          chatId,
          `❌ ${error.message || "Something went wrong. Try again."}`,
          messageId
        );
      } catch (notifyError) {
        console.error("Failed to send Telegram error:", notifyError.message);
      }

      console.error("Workflow error:", error.message);
    }

    return ok();
  } catch (err) {
    console.error("Webhook handler error:", err.message);
    return ok();
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
