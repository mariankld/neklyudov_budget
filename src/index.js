require("dotenv").config();
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const express = require("express");
const axios = require("axios");
const OpenAI = require("openai");
const { listWorkbookTables, appendTableRow } = require("./graphExcel");
const fxRates = require("./fxRates");
const { computeSyncHash, runSync, formatSyncReport } = require("./syncJob");
const { refreshCurrencyRatesTable, formatDailyJobsReport } = require("./dailyJobs");

const app = express();
app.use(express.json());

const PORT = Number(process.env.PORT || 3000);

/** Logical income category label stored in RAW (column 3); Income has no dedicated Excel table, RAW-only. */
const INCOME_SHEET = (process.env.INCOME_SHEET_NAME || "Income").trim();
const RAW_SHEET = (process.env.RAW_SHEET_NAME || "RAW").trim();
/** Optional audit-log table for every newly-fetched historical FX rate. Bot logs to it best-effort; missing table never blocks an expense write. */
const EXCHANGE_RATE_HISTORY_TABLE = (process.env.EXCHANGE_RATE_HISTORY_TABLE || "ExchangeRateHistory").trim();
/** "Today's rate" lookup table refreshed daily for anyone manually typing rows into Excel — see dailyJobs.js. */
const CURRENCY_RATES_TABLE_NAME = (process.env.CURRENCY_RATES_TABLE_NAME || "CurrencyRates").trim();
/** Optional chat id to receive daily-maintenance-job reports (CurrencyRates refresh + RAW<->category sync). If unset, reports are just logged to the console. */
const REPORT_CHAT_ID = process.env.TELEGRAM_REPORT_CHAT_ID ? String(process.env.TELEGRAM_REPORT_CHAT_ID).trim() : "";
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

/**
 * Categories the bot must never log to itself. Credit Cards was removed 2026-08-20: a live
 * header check (node scripts/migrate-workbook.js) confirmed CreditCards is currently a 17-column
 * table with the exact same shape as Health/MedInsurance, so it now shares that write branch
 * (see INSURANCE_CATEGORY_TABLES below) instead of the generic 14-column one — no risk of
 * landing data in the wrong columns. Once RowID/SyncHash are added (migrate-workbook.js --apply),
 * it's a normal pickable category like any other; the sync job also now scans it.
 */
const MANUAL_ONLY_CATEGORIES = [];

/**
 * Enumerated set of Payment Method values. Everything written to the workbook must be exactly
 * one of these strings — no free text — so Summary tab SUMIF/SUMIFS formulas that match on
 * Payment Method (e.g. "*Visa BOC*") never silently miss a row because of a syntax variant like
 * "VISA" or "Boc Visa". When OpenAI can't tell which one applies from the message/receipt, it
 * returns PAYMENT_METHOD_UNCLEAR and the user is asked to pick via Telegram buttons before the
 * transaction can be confirmed.
 *
 * CREDIT_CARD_PAYMENT_METHODS is the ONLY list the Summary!H20 "spent on credit cards" formula
 * matches against (see scripts/fix-h20-credit-card-total.js) — Octopus, WeChat Pay, and any
 * learned custom method are payment methods but never count as a credit card, so they must never
 * be added to this array.
 */
const CREDIT_CARD_PAYMENT_METHODS = ["Master SC", "Visa BOC", "Master BEA", "Master Citic"];
const OTHER_PAYMENT_METHODS = ["Cash", "Bank Transfer", "Octopus", "WeChat Pay"];
const BASE_PAYMENT_METHODS = [...CREDIT_CARD_PAYMENT_METHODS, ...OTHER_PAYMENT_METHODS];
const PAYMENT_METHOD_UNCLEAR = "Unclear";
/**
 * Written into Payment Method when the user declines to permanently save a genuinely-new method
 * string (the `newpm:...:n` button) — per Mariya (2026-08-20), declining must NOT discard the
 * expense or force a pick from the fixed list; the transaction still logs, just labeled as an
 * unfamiliar payment method so it's easy to find and fix by hand later. Distinct from
 * PAYMENT_METHOD_UNCLEAR (which still gates the draft behind a picker) — this value resolves
 * keyboardForDraft straight to buildConfirmKeyboard.
 */
const PAYMENT_METHOD_UNFAMILIAR = "Unfamiliar Payment Method";

/**
 * Learned payment methods (e.g. "PayMe", "AliPay") the user has confirmed via the Telegram
 * "add this as a payment method?" prompt — see maybeAddCustomPaymentMethod / the `newpm:`
 * callback branch. Never credit cards (that requires a code change + explicit review, since it
 * changes the Summary credit-card total). Persisted to CUSTOM_PAYMENT_METHODS_FILE so they
 * survive a restart; loaded once at bootstrap.
 */
const CUSTOM_PAYMENT_METHODS_FILE = path.join(__dirname, "..", "data", "customPaymentMethods.json");
let customPaymentMethods = [];

function loadCustomPaymentMethods() {
  try {
    const raw = fs.readFileSync(CUSTOM_PAYMENT_METHODS_FILE, "utf8");
    const parsed = JSON.parse(raw);
    customPaymentMethods = Array.isArray(parsed)
      ? parsed.filter((v) => typeof v === "string" && v.trim()).map((v) => v.trim())
      : [];
  } catch (err) {
    if (err.code !== "ENOENT") {
      console.error("loadCustomPaymentMethods: failed to read/parse, starting empty:", err.message);
    }
    customPaymentMethods = [];
  }
}

function saveCustomPaymentMethods() {
  try {
    fs.mkdirSync(path.dirname(CUSTOM_PAYMENT_METHODS_FILE), { recursive: true });
    fs.writeFileSync(CUSTOM_PAYMENT_METHODS_FILE, JSON.stringify(customPaymentMethods, null, 2));
  } catch (err) {
    // Best-effort: the new method still works for the rest of this process's lifetime even if
    // the write fails (e.g. read-only filesystem) — it just won't survive a restart.
    console.error("saveCustomPaymentMethods: failed to persist:", err.message);
  }
}

/** Full current list of valid Payment Method values — base (fixed) + any learned custom ones. */
function getPaymentMethods() {
  return [...BASE_PAYMENT_METHODS, ...customPaymentMethods];
}

/** Case-insensitive exact match against the current known list; returns the canonical stored casing, or null. */
function findKnownPaymentMethod(raw) {
  const needle = String(raw || "").trim().toLowerCase();
  if (!needle) return null;
  return getPaymentMethods().find((m) => m.toLowerCase() === needle) || null;
}

/**
 * Records a brand-new payment method the user just confirmed via the `newpm:...:y` button.
 * Never a credit card — adding a card requires a deliberate code change since it changes the
 * Summary!H20 credit-card total, which the owner reviews by hand (see fix-h20-credit-card-total.js).
 */
function addCustomPaymentMethod(name) {
  const clean = String(name || "").trim();
  if (!clean) return null;
  const existing = findKnownPaymentMethod(clean);
  if (existing) return existing;
  customPaymentMethods.push(clean);
  saveCustomPaymentMethods();
  console.log(`Payment method learned: "${clean}" (now ${getPaymentMethods().length} total).`);
  return clean;
}

/** Category labels the model may choose from: every expense category table plus Income (RAW-only). */
let allowedCategories = [];

const PENDING_TTL_MS = 24 * 60 * 60 * 1000;

/** token -> { draft, originalMessage, sender, sourceMessageId, chatId, createdAt, messageThreadId? } */
const pendingByToken = new Map();

/** Successfully logged tokens (idempotency); same TTL as pending. token -> completedAt ms */
const completedLogByToken = new Map();

/** Tokens currently being written (guards parallel webhook deliveries / double-tap races). */
const inflightLogTokens = new Set();

/**
 * token -> Promise<{ok:boolean, detail?:string}> for the in-flight write. Telegram sometimes
 * redelivers the same callback_query (e.g. if our ack was slow), so a second tap of the same
 * button can arrive while the first is still writing to Excel. Rather than replying with a
 * static "still saving" message — which can land in the chat *after* the real "✅ Logged
 * successfully" message once the original finishes, looking like a stale/confusing bug — the
 * duplicate awaits this same promise and reports the actual outcome.
 */
const inflightLogPromises = new Map();

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

/**
 * Shown instead of buildConfirmKeyboard whenever a draft's paymentMethod is PAYMENT_METHOD_UNCLEAR
 * (OpenAI couldn't tell how the transaction was paid). One button per known getPaymentMethods()
 * entry (base + any learned custom methods); callback_data encodes the token + the entry's index
 * into that same list so we never have to fit arbitrary card names into the 64-byte callback_data
 * limit. The list is read fresh each call so a just-learned method shows up immediately.
 */
function buildPaymentMethodKeyboard(token) {
  const methods = getPaymentMethods();
  const rows = [];
  for (let i = 0; i < methods.length; i += 2) {
    const row = [];
    for (const idx of [i, i + 1]) {
      if (idx >= methods.length) continue;
      const data = `pm:${token}:${idx}`;
      assertCallbackDataLength(data);
      row.push({ text: methods[idx], callback_data: data });
    }
    rows.push(row);
  }
  return { inline_keyboard: rows };
}

/**
 * Shown when a draft's paymentMethod is still PAYMENT_METHOD_UNCLEAR but paymentMethodRaw names
 * something that doesn't match any known payment method (findKnownPaymentMethod returns null) —
 * e.g. "PayMe" or "AliPay". Asks whether to save it permanently (addCustomPaymentMethod, which
 * updates the OpenAI-facing list immediately) or, on decline, log the expense anyway with
 * Payment Method set to PAYMENT_METHOD_UNFAMILIAR (see that constant — never discards the draft).
 */
function buildNewPaymentMethodKeyboard(token, rawPaymentMethod) {
  const yesData = `newpm:${token}:y`;
  const noData = `newpm:${token}:n`;
  assertCallbackDataLength(yesData);
  assertCallbackDataLength(noData);
  return {
    inline_keyboard: [
      [{ text: `✅ Yes, save "${rawPaymentMethod}"`.slice(0, 64), callback_data: yesData }],
      [{ text: `❌ No, log as "${PAYMENT_METHOD_UNFAMILIAR}"`.slice(0, 64), callback_data: noData }],
    ],
  };
}

/**
 * Gates the normal Yes/Edit keyboard behind a resolved payment method. Three states:
 *   1. paymentMethod already resolved -> buildConfirmKeyboard.
 *   2. paymentMethod unclear but paymentMethodRaw names something genuinely new (not a known
 *      method under any casing) -> buildNewPaymentMethodKeyboard, offering to learn it.
 *   3. paymentMethod unclear and nothing new to learn (raw empty, or it already matches a known
 *      method) -> buildPaymentMethodKeyboard, the plain picker.
 */
function keyboardForDraft(token, draft) {
  if (draft.paymentMethod !== PAYMENT_METHOD_UNCLEAR) {
    return buildConfirmKeyboard(token);
  }
  const raw = draft.paymentMethodRaw;
  if (raw && !findKnownPaymentMethod(raw)) {
    return buildNewPaymentMethodKeyboard(token, raw);
  }
  return buildPaymentMethodKeyboard(token);
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

/**
 * Built fresh on every OpenAI call (rather than a static const) so a payment method learned
 * mid-session via addCustomPaymentMethod() is immediately available to the very next parse —
 * no restart needed.
 *
 * paymentMethodRaw is intentionally NOT enum-constrained: it's how we detect an unfamiliar
 * payment method. paymentMethod itself must still be one of the known values (or "Unclear");
 * paymentMethodRaw carries the verbatim wording so the bot can offer to learn it.
 */
function buildTransactionJsonSchema() {
  return {
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
      paymentMethod: { type: "string", enum: [...getPaymentMethods(), PAYMENT_METHOD_UNCLEAR] },
      paymentMethodRaw: { type: "string" },
      recipient: { type: "string" },
      /** dd/mm/yyyy if an explicit or relative date is stated/legible, else null — never guessed. */
      date: { type: ["string", "null"] },
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
      "paymentMethodRaw",
      "recipient",
      "date",
    ],
    additionalProperties: false,
  };
}

/** Parses a strict dd/mm/yyyy (or dd-mm-yyyy / dd.mm.yyyy) string into a real, calendar-valid Date. Returns null if invalid/absent. */
function parseDateDdMmYyyy(str) {
  if (typeof str !== "string") return null;
  const m = str.trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (!m) return null;
  const day = Number(m[1]);
  const month = Number(m[2]);
  const year = Number(m[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const d = new Date(year, month - 1, day);
  if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) return null;
  return d;
}

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
    `- paymentMethod: MUST be exactly one of these strings — no other value is valid: ${getPaymentMethods().map((m) => `"${m}"`).join(", ")}. Match the message to the closest of these (e.g. "BOC Visa" or "the Visa" → "Visa BOC"; "bank" or "transfer" → "Bank Transfer"; "cash" → "Cash"; "octopus" → "Octopus"; "wechat" or "wechat pay" → "WeChat Pay"). If the message does not indicate how it was paid clearly enough to pick one of these with confidence, use "${PAYMENT_METHOD_UNCLEAR}" — never guess a specific card, and never invent a value outside this list.`,
    `- paymentMethodRaw: the exact word(s) the message used for how it was paid, verbatim (e.g. "PayMe", "Alipay", "octopus card"), even if you couldn't map it to one of the fixed paymentMethod options above. Empty string if the message says nothing at all about how it was paid. This is used to detect a payment method the user hasn't saved yet — never invent something the message doesn't actually say.`,
    `- recipient: empty string unless the message names a specific person the payment was made to or received from (e.g. a staff member, family member, or vendor contact) — e.g. "Получатель: Yaya" or "for Maria". Never invent a name.`,
    `- date: if the message states an explicit or relative date for the transaction (e.g. "yesterday", "on 15/03", "last Friday"), resolve it relative to the context date above and output as dd/mm/yyyy. If no date is stated, output null — never guess or default to today; the caller falls back to the message's own timestamp.`,
    "",
    `User message:\n${messageText}`,
  ].join("\n");
}

function buildEditPrompt(currentDraft, editInstruction, originalMessage) {
  const categoryList = allowedCategories.join(", ");
  return [
    "Apply the user's edit instructions to this draft. Keep other fields unchanged unless the edit implies them.",
    `This includes paymentMethod: keep its current value unless the edit instruction explicitly changes the payment method. paymentMethod MUST be exactly one of these strings — no other value is valid: ${getPaymentMethods().map((m) => `"${m}"`).join(", ")}. If the edit instruction requests a payment method that doesn't clearly match one of these, use "${PAYMENT_METHOD_UNCLEAR}" instead of inventing a new value.`,
    "This also includes date (currently dd/mm/yyyy in the draft): keep it unchanged unless the edit instruction explicitly gives a new date, e.g. \"date to 15/03/2026\", \"change date to yesterday\", \"it was actually last Friday\" — resolve relative dates the same way. Always output date as dd/mm/yyyy, never null, when editing an existing draft.",
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
        schema: buildTransactionJsonSchema(),
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
    `- paymentMethod: MUST be exactly one of these strings — no other value is valid: ${getPaymentMethods().map((m) => `"${m}"`).join(", ")}. Match what's visible in the image to the closest of these (e.g. a BOC-branded Visa card → "Visa BOC"; a cash receipt → "Cash"; a bank transfer confirmation → "Bank Transfer"; an Octopus card/app screen → "Octopus"; a WeChat Pay screen → "WeChat Pay"). If a card or payment app is shown but you can't confidently match it to one of these, or if no payment method is visible at all, use "${PAYMENT_METHOD_UNCLEAR}" — never guess a specific card, and never invent a value outside this list.`,
    `- paymentMethodRaw: the exact text or branding visible for how it was paid, verbatim (e.g. "PayMe", "Alipay", card issuer name printed on a receipt), even if you couldn't map it to one of the fixed paymentMethod options above. Empty string if nothing about payment method is visible. This is used to detect a payment method the user hasn't saved yet — never invent something not actually visible.`,
    "- recipient: empty string unless the image or caption names a specific person the payment was made to or received from (e.g. a staff member, family member, or vendor contact). Never invent a name.",
    "- date: read the transaction date printed on the receipt or screenshot if it is legible (e.g. a printed receipt date, or an iPhone Wallet payment timestamp), and output it as dd/mm/yyyy. Receipts are very often not from today. If no date is visible anywhere in the image, output null — never guess or default to today.",
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
        schema: buildTransactionJsonSchema(),
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
        schema: buildTransactionJsonSchema(),
      },
    },
  });

  const raw = completion.output_text;
  return JSON.parse(raw);
}

function normalizeDraft(result, messageDate) {
  const md = messageDate instanceof Date ? messageDate : new Date();
  /** Prefer a date read off the receipt/message or set via an edit; fall back to the Telegram message timestamp. */
  const extractedDate = parseDateDdMmYyyy(result.date);
  const effectiveDate = extractedDate || md;
  const normalized = {
    amount: Number(result.amount),
    description: String(result.description || "").trim(),
    category: String(result.category || "").trim(),
    subcategory: String(result.subcategory || "").trim(),
    type: String(result.type || "").trim().toLowerCase(),
    location: String(result.location || "").trim() || DEFAULT_LOCATION,
    currency: String(result.currency || "").trim() || DEFAULT_CURRENCY,
    notes: String(result.notes || "").trim(),
    paymentMethod: String(result.paymentMethod || "").trim() || PAYMENT_METHOD_UNCLEAR,
    /** Verbatim payment-method wording from the message/receipt, even when unmapped to a known enum value — used to detect and offer to learn a new payment method. */
    paymentMethodRaw: String(result.paymentMethodRaw || "").trim(),
    recipient: String(result.recipient || "").trim(),
    /** dd/mm/yyyy — from the receipt/message/edit when present, else the Telegram message time. */
    dateDdMmYyyy: formatDateDdMmYyyy(effectiveDate),
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
    "Sum (HKD): (calculated automatically from that day's exchange rate)",
    `Payment method: ${
      draft.paymentMethod === PAYMENT_METHOD_UNCLEAR || !draft.paymentMethod
        ? draft.paymentMethodRaw && !findKnownPaymentMethod(draft.paymentMethodRaw)
          ? `⚠️ New payment method detected: "${draft.paymentMethodRaw}" — save it? (see buttons below)`
          : "⚠️ Not yet chosen — please pick below"
        : draft.paymentMethod
    }`,
    `Recipient: ${draft.recipient && draft.recipient.length ? draft.recipient : "—"}`,
    `Logged by: ${draft.sender || "Unknown"}`,
    notesLine,
    "",
    "Does this look good?",
  ].join("\n");
}

function buildRawRowValues(draft, categorySheetName, sumHkd, rowId, syncHash) {
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
    sumHkd,
    notesCell,
    draft.sender || "Unknown",
    draft.paymentMethod || PAYMENT_METHOD_UNCLEAR,
    rowId || "",
    syncHash || "",
  ];
}

/**
 * Column order for every expense category Excel Table (14 columns as of the RowID/SyncHash
 * migration): Дата, Категория, Описание, Локация, Сумма, Валюта, Курс, Сумма (HKD), Примечание,
 * Метод оплаты, Получатель/Сотрудник, Пользователь, RowID, SyncHash.
 *
 * Mariya removed the 5 unused insurance columns (Страховка, Статус выплаты, Страховая
 * компания, Период покрытия, Карта) from every expense category table on 2026-08-19,
 * shrinking each table from 17 to 12 columns. The bot's old fixed 17-value write then
 * failed with Graph's InvalidArgument "number of rows or columns doesn't match". Пользователь
 * is now the last visible column instead of being sandwiched among the insurance fields — do
 * not reorder it back without re-confirming the live table layout. RowID/SyncHash were then
 * appended as two new technical columns (12, 13) — see scripts/migrate-workbook.js and
 * syncJob.js's STANDARD_CAT_COLS (this table's shape) / catColsFor() (per-table dispatch).
 * Both must be present on every row for the sync job to match/diff it; run the migration
 * script before deploying this code.
 *
 * Категория here is RAW's Subcategory (precision line item), not RAW's Category — confirmed
 * with Mariya. Курс and Сумма (HKD) used to be live Excel formulas (VLOOKUP against the single
 * current-rate CurrencyRates table). As of this change they are bot-computed static numbers
 * frozen to the transaction's own date, so a later rate update never retroactively changes a
 * past expense. See fxRates.js.
 */
/**
 * Health and MedInsurance (Insurance category) were NOT part of the 2026-08-19 insurance-column
 * removal — Mariya confirmed (2026-08-20) they must keep all 5 insurance columns (Страховка,
 * Статус выплаты, Страховая компания, Период покрытия, Карта), unlike every other category
 * table. The bot doesn't collect insurance-specific data via Telegram, so those 5 cells are
 * always written blank; RowID/SyncHash are still appended at the end so the sync job can match
 * these rows too. Keyed by Excel Table name (not category label) since that's what's passed in.
 * Keep in sync with scripts/migrate-workbook.js's INSURANCE_TABLES and syncJob.js's per-table
 * column maps — all three must agree on which tables keep the insurance columns.
 *
 * CreditCards added 2026-08-20: `node scripts/migrate-workbook.js` (dry run) confirmed the live
 * CreditCards table is currently sitting at the exact same 17-column layout as Health/
 * MedInsurance (Дата, Категория, Описание, Локация, Сумма, Валюта, Курс, Сумма (HKD),
 * Примечание, Метод оплаты, Получатель/Сотрудник, Страховка, Статус выплаты, Пользователь,
 * Страховая компания, Период покрытия, Карта) — it never got the 2026-08-19 cleanup either.
 * Nothing to do with insurance, it just happens to match column-for-column, so it reuses this
 * same branch/write shape. Summary!H21 filters CreditCards[Категория] (= Subcategory here, same
 * as everywhere else) for "*Выплата*" — log credit card debt repayments with subcategory
 * "Выплата" so that formula keeps picking them up.
 */
const INSURANCE_CATEGORY_TABLES = ["Health", "MedInsurance", "CreditCards"];

function buildCategoryRowValues(draft, rate, sumHkd, rowId, syncHash, tableName) {
  const notesCell = draft.notes && draft.notes.length ? draft.notes : "";

  if (INSURANCE_CATEGORY_TABLES.includes(tableName)) {
    // Old, pre-insurance-cleanup 17-column layout, preserved on purpose for this table, plus
    // RowID/SyncHash appended at the end (19 total). Column order confirmed 2026-08-20 via
    // scripts/inspect-excel-structure.js: Дата, Категория, Описание, Локация, Сумма, Валюта,
    // Курс, Сумма (HKD), Примечание, Метод оплаты, Получатель/Сотрудник, Страховка, Статус
    // выплаты, Пользователь, Страховая компания, Период покрытия, Карта.
    return [
      draft.dateDdMmYyyy, // Дата
      draft.subcategory, // Категория
      draft.description, // Описание
      draft.location, // Локация
      draft.amount, // Сумма
      draft.currency, // Валюта
      rate, // Курс
      sumHkd, // Сумма (HKD)
      notesCell, // Примечание
      draft.paymentMethod || PAYMENT_METHOD_UNCLEAR, // Метод оплаты
      draft.recipient || "", // Получатель/Сотрудник
      "", // Страховка — not collected via Telegram, left blank
      "", // Статус выплаты — not collected via Telegram, left blank
      draft.sender || "Unknown", // Пользователь
      "", // Страховая компания — not collected via Telegram, left blank
      "", // Период покрытия — not collected via Telegram, left blank
      "", // Карта — not collected via Telegram, left blank
      rowId || "", // RowID
      syncHash || "", // SyncHash
    ];
  }

  return [
    draft.dateDdMmYyyy, // Дата
    draft.subcategory, // Категория
    draft.description, // Описание
    draft.location, // Локация
    draft.amount, // Сумма
    draft.currency, // Валюта
    rate, // Курс — frozen to this transaction's date, bot-computed
    sumHkd, // Сумма (HKD) — frozen, bot-computed
    notesCell, // Примечание
    draft.paymentMethod || PAYMENT_METHOD_UNCLEAR, // Метод оплаты
    draft.recipient || "", // Получатель/Сотрудник
    draft.sender || "Unknown", // Пользователь
    rowId || "", // RowID — shared with the matching RAW row, written once at creation
    syncHash || "", // SyncHash — fingerprint of the mergeable fields, see syncJob.js
  ];
}

function roundMoney(n) {
  return Math.round(n * 100) / 100;
}

/** Appends one row to the (optional) FX audit-log table. Best-effort — never blocks an expense write. */
async function recordRateHistory({ currency, date, rate }) {
  try {
    const driveId = getExcelDriveIdFromEnv();
    const itemId = getExcelItemIdFromEnv();
    const source = currency === "RUB" ? "CBR (cbr.ru)" : "Frankfurter (ECB)";
    await appendTableRow(driveId, itemId, EXCHANGE_RATE_HISTORY_TABLE, [
      date,
      currency,
      rate,
      source,
      new Date().toISOString(),
    ]);
  } catch (err) {
    console.error(
      `Could not write to "${EXCHANGE_RATE_HISTORY_TABLE}" (run scripts/migrate-workbook.js if it doesn't exist yet):`,
      err.message
    );
  }
}

/**
 * Looks up (and caches) the historical rate for this transaction's currency+date and freezes
 * it into a static number. Never throws — if every FX source fails (network down, unsupported
 * currency, etc.) it falls back to blank cells so the expense still logs; the failure is
 * surfaced via `ok: false` so the Telegram success message can warn that the rate needs a
 * manual fill-in.
 */
async function computeFrozenHkdConversion(draft) {
  try {
    const rate = await fxRates.getRateToHkd(draft.currency, draft.dateDdMmYyyy, recordRateHistory);
    return { rate, sumHkd: roundMoney(draft.amount * rate), ok: true };
  } catch (err) {
    console.error(`FX lookup failed for ${draft.currency} ${draft.dateDdMmYyyy}:`, err.message);
    return { rate: "", sumHkd: "", ok: false, error: err.message };
  }
}

/**
 * Writes RAW first (always), then the matching category Excel Table for expenses only
 * (income has no dedicated table — RAW is the sole record). These are two independent
 * Graph API calls: if RAW succeeds but the category write fails, the thrown error is tagged
 * with rawSucceeded=true so the caller never retries (which would duplicate the RAW row).
 *
 * Every newly-created expense also gets a fresh RowID (shared by both the RAW and category
 * copies) and a SyncHash fingerprint of the fields that are supposed to mirror between them —
 * see syncJob.js. Both sides are written from the exact same draft in the same request, so
 * their fields necessarily agree right now; computing one hash and reusing it for both rows
 * means the very first sync pass sees them as already in sync (no spurious propagation).
 * Income rows get a RowID too (harmless, just never matched against anything, since income
 * has no category table) so every RAW row has a consistent shape.
 */
async function appendTransactionToExcel(draft) {
  const driveId = getExcelDriveIdFromEnv();
  const itemId = getExcelItemIdFromEnv();

  const fx = await computeFrozenHkdConversion(draft);

  const rowId = crypto.randomBytes(8).toString("hex");
  const syncHash = computeSyncHash({
    date: draft.dateDdMmYyyy,
    description: draft.description,
    location: draft.location,
    amount: Number(draft.amount),
    currency: String(draft.currency || "").trim().toUpperCase(),
    notes: draft.notes && draft.notes.length ? draft.notes : "",
    paymentMethod: draft.paymentMethod || PAYMENT_METHOD_UNCLEAR,
    subcategory: draft.subcategory,
  });

  const rawRow = buildRawRowValues(draft, draft.category, fx.sumHkd, rowId, syncHash);

  await appendTableRow(driveId, itemId, RAW_SHEET, rawRow);

  if (draft.type !== "expense") {
    return fx;
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
    const categoryRow = buildCategoryRowValues(draft, fx.rate, fx.sumHkd, rowId, syncHash, tableName);
    await appendTableRow(driveId, itemId, tableName, categoryRow);
  } catch (err) {
    err.rawSucceeded = true;
    throw err;
  }

  return fx;
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
    "",
    "RAW and the category tabs sync automatically every day. If you edited a row directly in Excel and want it reflected right away, send /sync.",
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
  allowedCategories = [
    ...Object.keys(CATEGORY_TABLE_MAP).filter((c) => !MANUAL_ONLY_CATEGORIES.includes(c)),
    INCOME_SHEET,
  ];

  loadCustomPaymentMethods();
  console.log(
    `Payment methods: ${getPaymentMethods().length} known (${BASE_PAYMENT_METHODS.length} base + ${customPaymentMethods.length} learned)${
      customPaymentMethods.length ? ` — learned: ${customPaymentMethods.join(", ")}` : ""
    }.`
  );

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

function isSyncCommand(text) {
  const t = (text || "").trim();
  return t === "/sync" || t.startsWith("/sync ");
}

/**
 * Runs both daily maintenance jobs (CurrencyRates refresh, RAW<->category sync) back to
 * back and returns a combined human-readable report. Shared by the daily scheduler and the
 * on-demand `/sync` Telegram command — `/sync` always includes both, since a stale
 * CurrencyRates table would otherwise silently persist until the next scheduled run.
 */
async function runDailyMaintenanceJobs() {
  const driveId = getExcelDriveIdFromEnv();
  const itemId = getExcelItemIdFromEnv();

  const currencyRates = await refreshCurrencyRatesTable({
    driveId,
    itemId,
    tableName: CURRENCY_RATES_TABLE_NAME,
  });

  const syncableCategoryTableMap = Object.fromEntries(
    Object.entries(CATEGORY_TABLE_MAP).filter(([category]) => !MANUAL_ONLY_CATEGORIES.includes(category))
  );

  const sync = await runSync({
    driveId,
    itemId,
    rawSheetName: RAW_SHEET,
    categoryTableMap: syncableCategoryTableMap,
  });

  const report = [formatDailyJobsReport({ currencyRates }), "", formatSyncReport(sync)].join("\n");
  return { currencyRates, sync, report };
}

let lastDailyJobsRunUtcDate = null;

/** Runs the daily jobs immediately and records today's UTC date so the hourly checker below doesn't re-run them again today. */
async function runDailyMaintenanceJobsNow(reason) {
  lastDailyJobsRunUtcDate = new Date().toISOString().slice(0, 10);
  try {
    console.log(`Running daily maintenance jobs (${reason})...`);
    const { report } = await runDailyMaintenanceJobs();
    console.log(`Daily maintenance jobs (${reason}) done:\n${report}`);
    if (REPORT_CHAT_ID) {
      try {
        await sendTelegramMessage(REPORT_CHAT_ID, `🗓️ Daily maintenance (${reason}):\n\n${report}`);
      } catch (err) {
        console.error("Failed to send daily maintenance report to Telegram:", err.message);
      }
    }
  } catch (err) {
    console.error(`Daily maintenance jobs (${reason}) failed:`, err.stack || err.message);
  }
}

/**
 * No native scheduler dependency (avoids adding node-cron): an hourly interval just checks
 * whether the UTC calendar date has changed since the last run, and if so runs once. Checked
 * hourly rather than e.g. daily-at-midnight so a Railway restart near midnight can't cause the
 * day to be skipped entirely.
 */
function scheduleDailyMaintenanceJobs() {
  setInterval(() => {
    const todayUtc = new Date().toISOString().slice(0, 10);
    if (todayUtc !== lastDailyJobsRunUtcDate) {
      runDailyMaintenanceJobsNow("daily schedule").catch((err) =>
        console.error("scheduleDailyMaintenanceJobs:", err.stack || err.message)
      );
    }
  }, 60 * 60 * 1000);
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
  const keyboard = keyboardForDraft(token, draft);

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
    date: entry.draft.dateDdMmYyyy,
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
    const keyboard = keyboardForDraft(newToken, draft);

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
        /** A duplicate/racing redelivery of the same tap — Telegram can resend a callback_query
         * if our ack was slow. Wait for the write that's already in progress and report its real
         * outcome instead of a static "still saving" message, which could otherwise land in the
         * chat after the "✅ Logged successfully" message once the original finishes. */
        const result = await (inflightLogPromises.get(token) || Promise.resolve({ ok: true }));
        await tell(
          result.ok
            ? "✅ Already logged, no action needed."
            : `❌ Log failed — nothing was saved.\n${result.detail || ""}`.trim()
        );
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

    const logPromise = (async () => {
      try {
        const fx = await appendTransactionToExcel(entry.draft);
        completedLogByToken.set(token, Date.now());

        const cat = entry.draft.category;
        const dLabel = entry.draft.dateDdMmYyyy;
        const tableNote = entry.draft.type === "expense" ? ` and "${cat}"` : "";
        const fxWarning =
          fx && fx.ok === false
            ? `\n⚠️ Exchange rate lookup failed — Курс/Sum (HKD) left blank, please fill in manually.`
            : "";
        const successText =
          `✅ Logged successfully.\n` +
          `Saved to "${RAW_SHEET}"${tableNote}.\n` +
          `${entry.draft.amount} ${entry.draft.currency} · ${dLabel}${fxWarning}`;

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
        return { ok: true };
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
          return { ok: true, detail: failDetail };
        }

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
        return { ok: false, detail: failDetail };
      } finally {
        inflightLogTokens.delete(token);
        inflightLogPromises.delete(token);
      }
    })();

    inflightLogPromises.set(token, logPromise);
    await logPromise;
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

  const pmPrefix = "pm:";
  if (data.startsWith(pmPrefix)) {
    const rest = data.slice(pmPrefix.length);
    const sepIdx = rest.lastIndexOf(":");
    const token = sepIdx === -1 ? rest : rest.slice(0, sepIdx);
    const idx = sepIdx === -1 ? NaN : Number(rest.slice(sepIdx + 1));

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
    const pmMethods = getPaymentMethods();
    if (!Number.isInteger(idx) || idx < 0 || idx >= pmMethods.length) {
      await tell("⚠️ Unrecognized payment method option — send the expense again.");
      return;
    }

    entry.draft.paymentMethod = pmMethods[idx];

    const pmThreadExtras =
      entry.messageThreadId != null
        ? { message_thread_id: entry.messageThreadId }
        : callbackExtras;
    const preview = formatDraftPreview(entry.draft, entry.draft.category);
    const keyboard = buildConfirmKeyboard(token);

    try {
      await editTelegramMessage(chatId, messageId, preview, keyboard, pmThreadExtras);
    } catch (e) {
      console.error("editTelegramMessage (pm select):", e.message);
      await sendTelegramMessage(chatId, preview, entry.sourceMessageId, keyboard, pmThreadExtras);
    }
    return;
  }

  const newpmPrefix = "newpm:";
  if (data.startsWith(newpmPrefix)) {
    const rest = data.slice(newpmPrefix.length);
    const sepIdx = rest.lastIndexOf(":");
    const token = sepIdx === -1 ? rest : rest.slice(0, sepIdx);
    const choice = sepIdx === -1 ? "" : rest.slice(sepIdx + 1);

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

    if (choice === "y") {
      // Learn it: persists to data/customPaymentMethods.json and is picked up by
      // getPaymentMethods() immediately, so the very next OpenAI parse call already knows it.
      const saved = addCustomPaymentMethod(entry.draft.paymentMethodRaw);
      entry.draft.paymentMethod = saved || entry.draft.paymentMethodRaw;
    } else {
      // Declined: per Mariya (2026-08-20), do NOT discard the expense and do NOT force a pick
      // from the fixed list — resolve paymentMethod straight to PAYMENT_METHOD_UNFAMILIAR so
      // keyboardForDraft falls through to buildConfirmKeyboard and the draft can be logged as-is,
      // just flagged for later cleanup. paymentMethodRaw is kept (not cleared) so the original
      // guess is still visible in the draft preview/notes if useful.
      entry.draft.paymentMethod = PAYMENT_METHOD_UNFAMILIAR;
    }

    const newpmThreadExtras =
      entry.messageThreadId != null
        ? { message_thread_id: entry.messageThreadId }
        : callbackExtras;
    const preview = formatDraftPreview(entry.draft, entry.draft.category);
    const keyboard = keyboardForDraft(token, entry.draft);

    try {
      await editTelegramMessage(chatId, messageId, preview, keyboard, newpmThreadExtras);
    } catch (e) {
      console.error("editTelegramMessage (newpm select):", e.message);
      await sendTelegramMessage(chatId, preview, entry.sourceMessageId, keyboard, newpmThreadExtras);
    }
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

    if (isSyncCommand(originalMessage)) {
      res.status(200).json({ ok: true });
      (async () => {
        try {
          await sendTelegramMessage(chatId, "⏳ Running sync (CurrencyRates + RAW ↔ category tables)...", messageId, undefined, threadExtras);
          const { report } = await runDailyMaintenanceJobs();
          lastDailyJobsRunUtcDate = new Date().toISOString().slice(0, 10);
          await sendTelegramMessage(chatId, `✅ Sync done.\n\n${report}`, messageId, undefined, threadExtras);
        } catch (err) {
          console.error("/sync failed:", err.stack || err.message);
          try {
            await sendTelegramMessage(chatId, `❌ Sync failed: ${err.message}`, messageId, undefined, threadExtras);
          } catch (notifyErr) {
            console.error("Failed to send /sync failure message:", notifyErr.message);
          }
        }
      })();
      return;
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
      runDailyMaintenanceJobsNow("startup").catch((err) =>
        console.error("Startup daily maintenance run failed:", err.stack || err.message)
      );
      scheduleDailyMaintenanceJobs();
    });
  })
  .catch((err) => {
    console.error("Startup failed:", err.message);
    process.exit(1);
  });
