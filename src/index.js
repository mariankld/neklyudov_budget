require("dotenv").config();
const crypto = require("crypto");
const express = require("express");
const axios = require("axios");
const OpenAI = require("openai");
const { listWorkbookTables, appendTableRow } = require("./graphExcel");

const app = express();
app.use(express.json());

const PORT = Number(process.env.PORT || 3000);

/** Logical income category label stored in RAW (column 3); Income has no dedicated Excel table, RAW-only. */
const INCOME_SHEET = (process.env.INCOME_SHEET_NAME || "Income").trim();
const RAW_SHEET = (process.env.RAW_SHEET_NAME || "RAW").trim();
const DEFAULT_CURRENCY = (process.env.DEFAULT_EXPENSE_CURRENCY || "HKD").trim();
const DEFAULT_LOCATION = (process.env.DEFAULT_EXPENSE_LOCATION || "Hong Kong").trim();
const FAMILY_GREETING_NAME = (process.env.BUDGET_FAMILY_NAME || "Neklyudov").trim();

/** Category label (as used in RAW / chosen by OpenAI) -> real Excel Table name in the workbook. */
const CATEGORY_TABLE_MAP = {
  "Credit Cards": "CreditCards",
  "Shopping": "Shopping",
  "Transportation": "TransportTable",
  "Utilities": "Utilities",
  "Entertainment": "Entertainment",
  "Restaurants": "Restaurants",
  "Family and Staff": "FamilyStaff",
  "Personal Spending": "StaffExpenses",
  "Other": "Other",
  "Subscriptions": "TelecomSubscriptions",
  "Travel": "Travel",
  "Health": "Health",
  "Education": "Education",
  "Rent": "Rent",
  "Insurance": "MedInsurance",
  "CAPEX": "CAPEX",
};

/** Category labels the model may choose from: every expense category table plus Income (RAW-only). */
let allowedCategories = [];

const PENDING_TTL_MS = 24 * 60 * 60 * 1000;

/** token -> { draft, originalMessage, sender, sourceMessageId, chatId, createdAt, messageThreadId? } */
const pendingByToken = new Map();

/** Successfully logged tokens (idempotency); same TTL as pending. token -> completedAt ms */
const completedLogByToken = new Map();

/** Tokens currently being written (guards parallel webhook deliveries / double-tap races). */
const inflightLogTokens = new Set();

/** chatId -> { token } — user clicked Edit and should reply with changes */
const awaitingEditByChat = new Map();

/** Lowercase alias → exact tab title; from CATEGORY_MAPPING JSON in .env */
let categoryAliasMap = new Map();

function getTelegramApiBase() {
  const t = process.env.TELEGRAM_BOT_TOKEN;
  if (!t || !String(t).trim()) {
    throw new Error("TELEGRAM_BOT_TOKEN is not set");
  }
  return `https://api.telegram.org/bot${String(t).trim()}`;
}

const telegramHttpTimeoutMs = () =>
  Number(process.env.TELEGRAM_HTTP_TIMEOUT_MS || 25000);

/** Telegram Bot API returns HTTP 200 with { ok: false } for many errors; axios does not throw. */
async function callTelegramApi(method, body) {
  try {
    const res = await axios.post(`${getTelegramApiBase()}/${method}`, body, {
      timeout: telegramHttpTimeoutMs(),
    });
    const data = res.data;
    if (!data?.ok) {
      throw new Error(data?.description || `Telegram ${method}: ${JSON.stringify(data)}`);
    }
    return data;
  } catch (err) {
    const desc = err.response?.data?.description;
    if (desc) {
      throw new Error(`Telegram ${method}: ${desc}`);
    }
    throw err;
  }
}

/** Fetches the Telegram file_path for a file_id (needed to build the download URL). */
async function getTelegramFile(fileId) {
  const res = await axios.get(`${getTelegramApiBase()}/getFile`, {
    params: { file_id: fileId },
    timeout: telegramHttpTimeoutMs(),
  });
  if (!res.data?.ok) {
    throw new Error(res.data?.description || "Telegram getFile failed");
  }
  const filePath = res.data.result?.file_path;
  if (!filePath) {
    throw new Error("Telegram getFile returned no file_path");
  }
  return filePath;
}

function getTelegramFileDownloadUrl(filePath) {
  const t = process.env.TELEGRAM_BOT_TOKEN;
  return `https://api.telegram.org/file/bot${String(t).trim()}/${filePath}`;
}

/** Downloads a Telegram-hosted file and returns it as base64 (for inline image input to OpenAI). */
async function downloadTelegramFileAsBase64(filePath) {
  const url = getTelegramFileDownloadUrl(filePath);
  const res = await axios.get(url, {
    responseType: "arraybuffer",
    timeout: Math.max(30000, telegramHttpTimeoutMs()),
  });
  return Buffer.from(res.data).toString("base64");
}

/** Best-effort mime type from the Telegram file path extension; Telegram photos are usually .jpg. */
function guessImageMimeType(filePath) {
  const lower = String(filePath || "").toLowerCase();
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".webp")) return "image/webp";
  if (lower.endsWith(".heic") || lower.endsWith(".heif")) return "image/heic";
  return "image/jpeg";
}

/** Telegram sends multiple resolutions for a photo; pick the highest-resolution one for OCR quality. */
function pickLargestPhoto(photoSizes) {
  if (!Array.isArray(photoSizes) || !photoSizes.length) return null;
  return photoSizes.reduce((best, p) => {
    if (!best) return p;
    const bestArea = (best.width || 0) * (best.height || 0);
    const area = (p.width || 0) * (p.height || 0);
    return area > bestArea ? p : best;
  }, null);
}

/**
 * Finds an image file_id on the message, whether sent as a compressed photo or as an
 * uncompressed image document (Telegram "Send without compression").
 */
function extractImageFileId(message) {
  const photo = pickLargestPhoto(message?.photo);
  if (photo?.file_id) {
    return { fileId: photo.file_id };
  }
  const doc = message?.document;
  if (doc?.file_id && typeof doc.mime_type === "string" && doc.mime_type.startsWith("image/")) {
    return { fileId: doc.file_id };
  }
  return null;
}

/**
 * Clears the client's loading state on inline buttons. Must not throw — runs before heavy work.
 * Uses a short timeout so the webhook can finish quickly even if Telegram is slow.
 */
async function safeAnswerCallbackQuery(callbackQueryId) {
  if (!callbackQueryId) return;
  try {
    const res = await axios.post(
      `${getTelegramApiBase()}/answerCallbackQuery`,
      { callback_query_id: callbackQueryId },
      { timeout: Math.min(15000, telegramHttpTimeoutMs()) }
    );
    if (!res.data?.ok) {
      console.error("answerCallbackQuery:", res.data?.description || res.data);
    }
  } catch (err) {
    console.error("answerCallbackQuery failed:", err.response?.data || err.message);
  }
}

/** Inline button callback_data must be 1–64 bytes (UTF-8). */
function assertCallbackDataLength(callbackData) {
  const n = Buffer.byteLength(String(callbackData), "utf8");
  if (n < 1 || n > 64) {
    throw new Error(`callback_data must be 1–64 bytes (got ${n}).`);
  }
}

function buildConfirmKeyboard(token) {
  const logData = `log:${token}`;
  const editData = `edit:${token}`;
  assertCallbackDataLength(logData);
  assertCallbackDataLength(editData);
  return {
    inline_keyboard: [
      [
        { text: "Yes, log it", callback_data: logData },
        { text: "Edit", callback_data: editData },
      ],
    ],
  };
}

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

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
    paymentMethod: { type: "string" },
    recipient: { type: "string" },
  },
  required: [
    "amount",
    "description",
    "category",
    "subcategory",
    "type",
    "location",
    "currency",
    "notes",
    "paymentMethod",
    "recipient",
  ],
  additionalProperties: false,
};

function buildAnalysisPrompt(messageText, contextDateLabel) {
  const categoryList = allowedCategories.join(", ");
  return [
    `The user's message date context: ${contextDateLabel}.`,
    `Available categories (pick exactly one name or a close synonym; synonyms map via env): ${categoryList}.`,
    `Never use "${RAW_SHEET}" as category—all rows are logged only there.`,
    `For salary, incoming transfers, or money received, use type "income" and category "${INCOME_SHEET}".`,
    "",
    "Fill fields as follows:",
    "- amount: primary numeric sum from the message.",
    "- description: short label; you may use the user's wording.",
    `- subcategory: specific line item (e.g. Transport → Highway toll).`,
    `- location: default "${DEFAULT_LOCATION}" unless the message states another place.`,
    `- currency: default "${DEFAULT_CURRENCY}" unless the message states another currency.`,
    `- notes: empty string unless the user explicitly adds an extra note beyond amount/description; do not duplicate the description.`,
    `- paymentMethod: if the message names how it was paid (e.g. cash, a specific card, or bank), be as precise as the message allows (e.g. "BOC Visa", "HSBC transfer", "Cash"). If nothing is mentioned, use "Unknown". Never invent a payment method.`,
    `- recipient: empty string unless the message names a specific person the payment was made to or received from (e.g. a staff member, family member, or vendor contact) — e.g. "Получатель: Yaya" or "for Maria". Never invent a name.`,
    "",
    `User message:\n${messageText}`,
  ].join("\n");
}

function buildEditPrompt(currentDraft, editInstruction, originalMessage) {
  const categoryList = allowedCategories.join(", ");
  return [
    "Apply the user's edit instructions to this draft. Keep other fields unchanged unless the edit implies them.",
    "This includes paymentMethod: keep its current value unless the edit instruction explicitly changes the payment method.",
    `Allowed categories: ${categoryList}. Never use "${RAW_SHEET}" as category name.`,
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

async function sendTelegramMessage(chatId, text, replyToMessageId, replyMarkup, extras = {}) {
  const body = {
    chat_id: chatId,
    text,
    ...extras,
  };
  if (replyToMessageId != null) {
    body.reply_to_message_id = replyToMessageId;
  }
  if (replyMarkup) {
    body.reply_markup = replyMarkup;
  }
  await callTelegramApi("sendMessage", body);
}

async function editTelegramMessage(chatId, messageId, text, replyMarkup, extras = {}) {
  const body = {
    chat_id: chatId,
    message_id: messageId,
    text,
    ...extras,
  };
  if (replyMarkup !== undefined) {
    body.reply_markup = replyMarkup;
  }
  await callTelegramApi("editMessageText", body);
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
        content: `You analyze Telegram budget messages for a family spreadsheet. Output strict JSON only. Map spending to the closest allowed category name. Today for context is ${contextDateLabel}.`,
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

function buildImageAnalysisPrompt(captionText, contextDateLabel) {
  const categoryList = allowedCategories.join(", ");
  const captionLine =
    captionText && captionText.trim()
      ? `The user also added this caption: "${captionText.trim()}".`
      : "No caption was provided.";
  return [
    `The user's message date context: ${contextDateLabel}.`,
    "The attached image is a receipt or a screenshot of a payment (for example an iPhone Wallet payment confirmation).",
    "Read the amount, merchant/payee, and any other details you can from the image itself.",
    captionLine,
    `Available categories (pick exactly one name or a close synonym; synonyms map via env): ${categoryList}.`,
    `Never use "${RAW_SHEET}" as category—all rows are logged only there.`,
    `For salary, incoming transfers, or money received, use type "income" and category "${INCOME_SHEET}".`,
    "",
    "Fill fields as follows:",
    "- amount: the primary total amount charged or paid, as shown in the image. Look carefully, even if the text is small, blurry, or partially compressed, and report your best-effort reading of that amount.",
    "- Only set amount to 0 if there is truly no numeric amount visible anywhere in the image. Never guess or invent a number, and never default to 0 just because the text is slightly hard to read—read it as carefully as you can first.",
    "- description: merchant or payee name, or a short label for what the image shows.",
    "- subcategory: a specific line item if visible (e.g. Transport → Highway toll).",
    `- location: default "${DEFAULT_LOCATION}" unless the image or caption states another place.`,
    `- currency: default "${DEFAULT_CURRENCY}" unless the image clearly shows another currency symbol or code.`,
    "- notes: empty string unless there is a clear extra detail worth keeping (e.g. last 4 card digits, transaction id); do not duplicate the description.",
    "- paymentMethod: read this as precisely as the image allows—bank name plus card type/tier if shown (e.g. \"BOC VISA Infinite\", \"DBS Mastercard\"), or \"Cash\" for cash receipts. If a card is shown but the bank/type isn't legible, use \"Card (unspecified)\". If there is no payment method visible at all, use \"Unknown\". Never invent a bank or card name that isn't actually visible.",
    "- recipient: empty string unless the image or caption names a specific person the payment was made to or received from (e.g. a staff member, family member, or vendor contact). Never invent a name.",
  ].join("\n");
}

async function parseTransactionFromImage(base64Image, mimeType, captionText, messageDate) {
  const contextDateLabel = formatContextDateLabel(messageDate);
  const completion = await openai.responses.create({
    model: process.env.OPENAI_VISION_MODEL || "gpt-4o",
    input: [
      {
        role: "system",
        content: `You analyze photos of receipts or payment screenshots for a family budget spreadsheet. Output strict JSON only. Today for context is ${contextDateLabel}.`,
      },
      {
        role: "user",
        content: [
          { type: "input_text", text: buildImageAnalysisPrompt(captionText, contextDateLabel) },
          { type: "input_image", image_url: `data:${mimeType};base64,${base64Image}` },
        ],
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
    paymentMethod: String(result.paymentMethod || "").trim() || "Unknown",
    recipient: String(result.recipient || "").trim(),
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
    `Category: ${categorySheetName}`,
    `Subcategory: ${draft.subcategory || "—"}`,
    `Description: ${draft.description}`,
    `Location: ${draft.location}`,
    `Sum: ${draft.amount}`,
    `Currency: ${draft.currency}`,
    "Sum (HKD): (leave blank — calculated in sheet)",
    `Payment method: ${draft.paymentMethod || "Unknown"}`,
    `Recipient: ${draft.recipient && draft.recipient.length ? draft.recipient : "—"}`,
    `Logged by: ${draft.sender || "Unknown"}`,
    notesLine,
    "",
    "Does this look good?",
  ].join("\n");
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
    draft.sender || "Unknown",
    draft.paymentMethod || "Unknown",
  ];
}

/**
 * Column order for every expense category Excel Table (17 columns), confirmed live via
 * inspect-excel-structure.js: Дата, Категория, Описание, Локация, Сумма, Валюта, Курс,
 * Сумма (HKD), Примечание, Метод оплаты, Получатель/Сотрудник, Страховка, Статус выплаты,
 * Пользователь, Страховая компания, Период покрытия, Карта.
 *
 * Категория here is RAW's Subcategory (precision line item), not RAW's Category — confirmed
 * with Mariya. Курс and Сумма (HKD) are Excel formulas (exchange rate lookup) and must never
 * be written by the bot. Страховка / Статус выплаты / Страховая компания / Период покрытия /
 * Карта are not populated by the Telegram flow and are left blank.
 */
function buildCategoryRowValues(draft) {
  const notesCell = draft.notes && draft.notes.length ? draft.notes : "";

  return [
    draft.dateDdMmYyyy, // Дата
    draft.subcategory, // Категория
    draft.description, // Описание
    draft.location, // Локация
    draft.amount, // Сумма
    draft.currency, // Валюта
    "", // Курс — Excel formula, never written here
    "", // Сумма (HKD) — Excel formula, never written here
    notesCell, // Примечание
    draft.paymentMethod || "Unknown", // Метод оплаты
    draft.recipient || "", // Получатель/Сотрудник
    "", // Страховка — not tracked by the bot
    "", // Статус выплаты — not tracked by the bot
    draft.sender || "Unknown", // Пользователь
    "", // Страховая компания — not tracked by the bot
    "", // Период покрытия — not tracked by the bot
    "", // Карта — not tracked by the bot
  ];
}

/**
 * Writes RAW first (always), then the matching category Excel Table for expenses only
 * (income has no dedicated table — RAW is the sole record). These are two independent
 * Graph API calls: if RAW succeeds but the category write fails, the thrown error is tagged
 * with rawSucceeded=true so the caller never retries (which would duplicate the RAW row).
 */
async function appendTransactionToExcel(draft) {
  const driveId = getExcelDriveIdFromEnv();
  const itemId = getExcelItemIdFromEnv();
  const rawRow = buildRawRowValues(draft, draft.category);

  await appendTableRow(driveId, itemId, RAW_SHEET, rawRow);

  if (draft.type !== "expense") {
    return;
  }

  const tableName = CATEGORY_TABLE_MAP[draft.category];
  if (!tableName) {
    const err = new Error(
      `RAW was logged, but "${draft.category}" has no matching Excel Table configured (CATEGORY_TABLE_MAP).`
    );
    err.rawSucceeded = true;
    throw err;
  }

  try {
    const categoryRow = buildCategoryRowValues(draft);
    await appendTableRow(driveId, itemId, tableName, categoryRow);
  } catch (err) {
    err.rawSucceeded = true;
    throw err;
  }
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
    "I'll analyze the amount and description, show a preview (category + RAW columns), and wait for you to tap Yes, log it or Edit before anything is saved. ✅",
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

function getExcelDriveIdFromEnv() {
  const v = process.env.EXCEL_DRIVE_ID;
  return v ? String(v).trim() : "";
}

function getExcelItemIdFromEnv() {
  const v = process.env.EXCEL_ITEM_ID;
  return v ? String(v).trim() : "";
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
  /** Must include callback_query or inline-button taps are never POSTed to this URL. */
  const allowedUpdates = ["message", "callback_query", "edited_message"];

  try {
    await callTelegramApi("setWebhook", {
      url: webhookUrl,
      allowed_updates: allowedUpdates,
      drop_pending_updates: false,
    });

    const infoRes = await axios.get(`${getTelegramApiBase()}/getWebhookInfo`, {
      timeout: 15000,
    });
    const info = infoRes.data?.result;
    if (!infoRes.data?.ok) {
      console.error("getWebhookInfo failed:", infoRes.data);
      return;
    }
    console.log(
      `Telegram webhook → ${info?.url || webhookUrl} (pending: ${info?.pending_update_count ?? 0}) allowed_updates=${JSON.stringify(info?.allowed_updates)}`
    );
    const au = info?.allowed_updates;
    if (Array.isArray(au) && au.length > 0 && !au.includes("callback_query")) {
      console.error(
        "WARNING: Telegram webhook is not subscribed to callback_query — inline buttons will not reach this server. Redeploy or call setWebhook with allowed_updates including callback_query."
      );
    }
  } catch (err) {
    console.error("Telegram setWebhook error:", err.response?.data || err.message);
  }
}

async function bootstrap() {
  requireEnv("TELEGRAM_BOT_TOKEN");
  requireEnv("OPENAI_API_KEY");
  requireEnv("AZURE_TENANT_ID");
  requireEnv("AZURE_CLIENT_ID");
  requireEnv("AZURE_CLIENT_SECRET");
  requireEnv("EXCEL_DRIVE_ID");
  requireEnv("EXCEL_ITEM_ID");

  const driveId = getExcelDriveIdFromEnv();
  const itemId = getExcelItemIdFromEnv();

  categoryAliasMap = parseCategoryMappingEnv();
  allowedCategories = [...Object.keys(CATEGORY_TABLE_MAP), INCOME_SHEET];

  console.log(
    `Budget logger: ${allowedCategories.length} categories — ${allowedCategories.join(", ")}`
  );
  console.log(
    `Every transaction is appended to "${RAW_SHEET}"; expenses are also appended to their category Excel Table.`
  );
  if (categoryAliasMap.size) {
    console.log(`Category aliases loaded: ${categoryAliasMap.size}`);
  }

  /** Startup sanity check: confirm RAW and every mapped category table actually exist in the live workbook. */
  const tables = await listWorkbookTables(driveId, itemId);
  const liveTableNames = new Set((tables || []).map((t) => t.name || t.id));
  const requiredTableNames = [RAW_SHEET, ...Object.values(CATEGORY_TABLE_MAP)];
  const missing = requiredTableNames.filter((name) => !liveTableNames.has(name));
  if (missing.length) {
    throw new Error(
      `Excel workbook is missing expected table(s): ${missing.join(", ")}. ` +
        `Found tables: ${[...liveTableNames].join(", ")}. Check CATEGORY_TABLE_MAP / EXCEL_DRIVE_ID / EXCEL_ITEM_ID.`
    );
  }
  console.log(`Excel workbook check OK: ${requiredTableNames.length} required tables found.`);
}

function isCancelCommand(text) {
  const t = (text || "").trim();
  return t === "/cancel" || t.startsWith("/cancel ");
}

async function presentDraftForConfirmation(parsed, context) {
  const { chatId, originalMessage, sender, messageId, messageDate, messageThreadId } = context;
  const draft = normalizeDraft(parsed, messageDate);
  draft.sender = sender || "Unknown";
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
    messageThreadId,
  });

  const threadExtras = messageThreadId != null ? { message_thread_id: messageThreadId } : {};
  const preview = formatDraftPreview(draft, categorySheetName);
  const keyboard = buildConfirmKeyboard(token);

  await sendTelegramMessage(chatId, preview, messageId, keyboard, threadExtras);
}

async function processExpenseFlow(chatId, originalMessage, sender, messageId, messageDate, messageThreadId) {
  awaitingEditByChat.delete(chatId);
  const threadExtras = messageThreadId != null ? { message_thread_id: messageThreadId } : {};
  await sendTelegramMessage(chatId, "⏳ Analyzing your message...", messageId, undefined, threadExtras);

  const parsed = await parseTransactionWithOpenAI(originalMessage, messageDate);
  await presentDraftForConfirmation(parsed, {
    chatId,
    originalMessage,
    sender,
    messageId,
    messageDate,
    messageThreadId,
  });
}

async function processImageExpenseFlow(chatId, fileId, captionText, sender, messageId, messageDate, messageThreadId) {
  awaitingEditByChat.delete(chatId);
  const threadExtras = messageThreadId != null ? { message_thread_id: messageThreadId } : {};
  await sendTelegramMessage(chatId, "⏳ Reading your photo...", messageId, undefined, threadExtras);

  const filePath = await getTelegramFile(fileId);
  const base64Image = await downloadTelegramFileAsBase64(filePath);
  const mimeType = guessImageMimeType(filePath);

  const parsed = await parseTransactionFromImage(base64Image, mimeType, captionText, messageDate);
  const originalMessageLabel =
    captionText && captionText.trim() ? `[photo] ${captionText.trim()}` : "[photo]";

  await presentDraftForConfirmation(parsed, {
    chatId,
    originalMessage: originalMessageLabel,
    sender,
    messageId,
    messageDate,
    messageThreadId,
  });
}

async function processEditInstruction(chatId, editText, messageId, messageThreadId) {
  const wait = awaitingEditByChat.get(chatId);
  if (!wait) {
    return false;
  }

  const entry = pendingByToken.get(wait.token);
  if (!entry) {
    awaitingEditByChat.delete(chatId);
    const threadExtras = messageThreadId != null ? { message_thread_id: messageThreadId } : {};
    await sendTelegramMessage(
      chatId,
      "That edit session expired. Send the expense again.",
      messageId,
      undefined,
      threadExtras
    );
    return true;
  }

  prunePending();

  const threadExtras =
    messageThreadId != null
      ? { message_thread_id: messageThreadId }
      : entry.messageThreadId != null
        ? { message_thread_id: entry.messageThreadId }
        : {};

  const flatDraft = {
    amount: entry.draft.amount,
    description: entry.draft.description,
    category: entry.draft.category,
    subcategory: entry.draft.subcategory,
    type: entry.draft.type,
    location: entry.draft.location,
    currency: entry.draft.currency,
    notes: entry.draft.notes,
    paymentMethod: entry.draft.paymentMethod,
    recipient: entry.draft.recipient,
  };

  try {
    await sendTelegramMessage(chatId, "⏳ Updating your draft...", messageId, undefined, threadExtras);
    const parsed = await parseEditWithOpenAI(flatDraft, editText, entry.originalMessage);
    const draft = normalizeDraft(parsed, entry.messageDate);
    draft.sender = entry.sender || "Unknown";
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
      messageThreadId: entry.messageThreadId,
    });
    awaitingEditByChat.delete(chatId);

    const preview = formatDraftPreview(draft, categorySheetName);
    const keyboard = buildConfirmKeyboard(newToken);

    await sendTelegramMessage(chatId, preview, entry.sourceMessageId, keyboard, threadExtras);
  } catch (err) {
    await sendTelegramMessage(
      chatId,
      `❌ Could not apply changes: ${err.message}. Try again or send /cancel.`,
      messageId,
      undefined,
      threadExtras
    );
  }
  return true;
}

async function handleCallbackQuery(callbackQuery) {
  const data = String(callbackQuery.data || "");
  const msg = callbackQuery.message;
  const chatId = msg?.chat?.id;
  const messageId = msg?.message_id;
  const callbackThreadId = msg?.message_thread_id;

  const callbackExtras =
    callbackThreadId != null ? { message_thread_id: callbackThreadId } : {};

  async function tell(text) {
    if (!chatId) return;
    try {
      await sendTelegramMessage(chatId, text, null, undefined, callbackExtras);
    } catch (e) {
      console.error("callback notice send:", e.message);
    }
  }

  if (!chatId || messageId == null) {
    console.error(
      "callback_query missing message (cannot notify user); data=%s",
      data.slice(0, 60)
    );
    return;
  }

  const logPrefix = "log:";
  const editPrefix = "edit:";

  if (data.startsWith(logPrefix)) {
    const token = data.slice(logPrefix.length);
    prunePending();

    if (completedLogByToken.has(token)) {
      await tell(`ℹ️ This line was already logged. Check "${RAW_SHEET}".`);
      return;
    }

    const entry = pendingByToken.get(token);
    if (!entry) {
      if (inflightLogTokens.has(token)) {
        await tell("⏳ Still saving that expense — watch for an update in this chat.");
        return;
      }
      await tell(
        "⚠️ This confirmation is not on this server anymore — usually a restart or a second app instance. Send the expense again. On Railway/hosting, use exactly one replica unless you add shared storage."
      );
      return;
    }
    if (entry.chatId !== chatId) {
      await tell("Not allowed.");
      return;
    }

    if (!pendingByToken.delete(token)) {
      await tell("⏳ Already saving this expense — watch for the next message.");
      return;
    }
    inflightLogTokens.add(token);

    const entryThreadExtras =
      entry.messageThreadId != null
        ? { message_thread_id: entry.messageThreadId }
        : callbackExtras;

    try {
      await appendTransactionToExcel(entry.draft);
      completedLogByToken.set(token, Date.now());

      const cat = entry.draft.category;
      const dLabel = entry.draft.dateDdMmYyyy;
      const tableNote = entry.draft.type === "expense" ? ` and "${cat}"` : "";
      const successText =
        `✅ Logged successfully.\n` +
        `Saved to "${RAW_SHEET}"${tableNote}.\n` +
        `${entry.draft.amount} ${entry.draft.currency} · ${dLabel}`;

      try {
        await editTelegramMessage(chatId, messageId, successText, { inline_keyboard: [] }, entryThreadExtras);
      } catch {
        await sendTelegramMessage(
          chatId,
          successText,
          null,
          undefined,
          entryThreadExtras
        );
      }
      awaitingEditByChat.delete(chatId);
    } catch (err) {
      console.error("Log callback error:", err.message);

      const failDetail = (err.message || "Unknown error").slice(0, 400);
      if (err.rawSucceeded) {
        /** RAW already has this row — do NOT restore to pendingByToken, a retry would duplicate it. */
        completedLogByToken.set(token, Date.now());
        try {
          await sendTelegramMessage(
            chatId,
            `⚠️ Logged to "${RAW_SHEET}" but the "${entry.draft.category}" table write failed — check the workbook.\n${failDetail}`,
            null,
            undefined,
            entryThreadExtras
          );
        } catch (notifyError) {
          console.error("Failed to send Telegram error:", notifyError.message);
        }
      } else {
        pendingByToken.set(token, entry);
        try {
          await sendTelegramMessage(
            chatId,
            `❌ Log failed — nothing was saved.\n${failDetail}`,
            null,
            undefined,
            entryThreadExtras
          );
        } catch (notifyError) {
          console.error("Failed to send Telegram error:", notifyError.message);
        }
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
      await tell(
        "⚠️ This confirmation is not on this server anymore. Send the expense again. With multiple replicas, use only one instance."
      );
      return;
    }
    if (entry.chatId !== chatId) {
      await tell("Not allowed.");
      return;
    }

    awaitingEditByChat.set(chatId, { token });
    const editThreadExtras =
      entry.messageThreadId != null
        ? { message_thread_id: entry.messageThreadId }
        : callbackExtras;

    await sendTelegramMessage(
      chatId,
      "✏️ What would you like to change? For example: category to Shopping, amount 120, currency USD, or location Tokyo.",
      messageId,
      undefined,
      editThreadExtras
    );
    return;
  }

  await tell("Unknown button.");
}

app.post("/webhook/telegram", async (req, res) => {
  const ok = () => res.status(200).json({ ok: true });

  try {
    const callbackQuery = req.body?.callback_query;
    if (callbackQuery) {
      console.log(
        "Telegram callback:",
        String(callbackQuery.data || "").slice(0, 48),
        "chat",
        callbackQuery.message?.chat?.id
      );
      await safeAnswerCallbackQuery(callbackQuery.id);
      res.status(200).json({ ok: true });
      handleCallbackQuery(callbackQuery).catch((err) =>
        console.error("handleCallbackQuery:", err.stack || err.message)
      );
      return;
    }

    const message = req.body?.message;
    if (!message && req.body && Object.keys(req.body).length) {
      console.log(
        "Telegram webhook (no message/callback):",
        Object.keys(req.body).join(", ")
      );
    }
    const chatId = message?.chat?.id;
    const originalMessage = message?.text;
    const imageInput = extractImageFileId(message);
    // No leading "@" here: this workbook has Excel's Lotus transition formula
    // entry compatibility active, which silently reinterprets any cell value
    // starting with "@" as a formula (e.g. "@name" -> "=name"), producing
    // #NAME? errors and — because it's a Table calculated column — smearing
    // that single bad result across every row in the column.
    const senderUsername = message?.from?.username || null;
    const senderFallback = message?.from?.first_name || `${message?.from?.id || "Unknown"}`;
    const sender = senderUsername || senderFallback;
    const messageId = message?.message_id;
    const messageThreadId = message?.message_thread_id;
    const threadExtras = messageThreadId != null ? { message_thread_id: messageThreadId } : {};

    if (!chatId || (!originalMessage && !imageInput)) {
      return ok();
    }

    const messageDate = new Date((message.date || Math.floor(Date.now() / 1000)) * 1000);

    if (isStartCommand(originalMessage)) {
      awaitingEditByChat.delete(chatId);
      try {
        await sendTelegramMessage(chatId, getWelcomeMessage(), messageId, undefined, threadExtras);
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
          messageId,
          undefined,
          threadExtras
        );
      } catch (err) {
        console.error("Telegram send error:", err.message);
      }
      return ok();
    }

    try {
      if (imageInput) {
        /** A photo always starts a fresh draft; it is never treated as a reply to a pending edit. */
        awaitingEditByChat.delete(chatId);
        await processImageExpenseFlow(
          chatId,
          imageInput.fileId,
          message?.caption,
          sender,
          messageId,
          messageDate,
          messageThreadId
        );
        return ok();
      }

      if (awaitingEditByChat.has(chatId)) {
        const handled = await processEditInstruction(
          chatId,
          originalMessage.trim(),
          messageId,
          messageThreadId
        );
        if (handled) {
          return ok();
        }
      }

      await processExpenseFlow(
        chatId,
        originalMessage,
        sender,
        messageId,
        messageDate,
        messageThreadId
      );
    } catch (error) {
      try {
        await sendTelegramMessage(
          chatId,
          `❌ ${error.message || "Something went wrong. Try again."}`,
          messageId,
          undefined,
          threadExtras
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
