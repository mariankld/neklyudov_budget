require("dotenv").config();
const express = require("express");
const axios = require("axios");
const OpenAI = require("openai");

const app = express();
app.use(express.json());

const PORT = Number(process.env.PORT || 3000);

const DEFAULT_ALLOWED_CATEGORIES = [
  "Shopping",
  "Транспорт",
  "Utilities",
  "Развлечения",
  "Рестораны",
  "Семья и персонал",
  "Расходы Персонал",
  "Прочее",
  "Связь и подписки",
  "Путешествия",
  "Здоровье",
  "Образование",
  "Аренда",
  "Мед. страховка",
  "CAPEX",
  "Credit Cards",
  "Доходы",
];

const INCOME_SHEET = "Доходы";
const RAW_SHEET = "Expenses_RAW";
const ALLOWED_CATEGORIES = (
  process.env.ALLOWED_CATEGORIES || DEFAULT_ALLOWED_CATEGORIES.join("|")
)
  .split("|")
  .map((value) => value.trim())
  .filter(Boolean);

if (!ALLOWED_CATEGORIES.includes(INCOME_SHEET)) {
  ALLOWED_CATEGORIES.push(INCOME_SHEET);
}

const TELEGRAM_API_BASE = `https://api.telegram.org/bot${process.env.TELEGRAM_BOT_TOKEN}`;
const GRAPH_API_BASE = "https://graph.microsoft.com/v1.0";

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

function getIsoParts(date) {
  const isoDate = date.toISOString().slice(0, 10);
  const time = date.toTimeString().slice(0, 8);
  return { isoDate, time };
}

function escapeSheetName(sheetName) {
  return sheetName.replace(/'/g, "''");
}

function buildPrompt(messageText) {
  const categoryList = ALLOWED_CATEGORIES.join(", ");
  return [
    `Extract from this message: (1) numeric amount, (2) description, (3) category (choose ONE from: ${categoryList}), and (4) type (expense or income).`,
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

async function getMicrosoftAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    client_id: process.env.MS_CLIENT_ID,
    client_secret: process.env.MS_CLIENT_SECRET,
    grant_type: "client_credentials",
    scope: "https://graph.microsoft.com/.default",
  });

  const response = await axios.post(tokenUrl, params, {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });
  return response.data.access_token;
}

async function appendRowToSheet(accessToken, sheetName, rowValues) {
  const url = `${GRAPH_API_BASE}/drives/${process.env.ONEDRIVE_DRIVE_ID}/items/${process.env.EXCEL_FILE_ID}/workbook/worksheets('${escapeSheetName(sheetName)}')/tables('${process.env.EXCEL_TABLE_NAME_PREFIX}${sheetName}')/rows/add`;
  await axios.post(
    url,
    { values: [rowValues] },
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
}

async function parseWithOpenAI(messageText) {
  const completion = await openai.responses.create({
    model: process.env.OPENAI_MODEL || "gpt-5.5-medium",
    input: [
      {
        role: "system",
        content:
          "You are a financial parsing assistant. Extract expense/income amounts and determine the most appropriate category. For incoming payments/transfers, set type to 'income' and category to 'Доходы'. For outgoing expenses, set type to 'expense' and use the expense categories provided. Today's date is 2026-05-01. Output strict JSON only.",
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
            category: { type: "string", enum: ALLOWED_CATEGORIES },
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

  if (!ALLOWED_CATEGORIES.includes(normalized.category)) {
    throw new Error(`Unsupported category: ${normalized.category}`);
  }

  if (normalized.type === "income") {
    normalized.category = INCOME_SHEET;
  } else if (normalized.type !== "expense") {
    throw new Error(`Unsupported type: ${normalized.type}`);
  }

  return normalized;
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
      now.toISOString().replace("T", " ").slice(0, 19),
      sender,
      originalMessage,
      aiResult.amount,
      aiResult.type,
      categorySheetName,
      aiResult.description,
      isoDate,
      time,
    ];

    const token = await getMicrosoftAccessToken();

    await appendRowToSheet(token, categorySheetName, categoryRow);
    await appendRowToSheet(token, RAW_SHEET, rawRow);

    await sendTelegramMessage(
      chatId,
      `✅ Logged: ${aiResult.amount} ${aiResult.type} under ${categorySheetName} on ${isoDate}`,
      messageId
    );
  } catch (error) {
    try {
      await sendTelegramMessage(
        chatId,
        "❌ Error saving to Excel. Please check the workbook access and try again.",
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

app.listen(PORT, () => {
  console.log(`Budget logger is running on port ${PORT}`);
});
