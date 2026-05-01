# Neklyudov Budget Logger

Telegram -> OpenAI -> Excel workflow that replaces the Gemini/Zapier parser with OpenAI and logs transactions to both:

- category sheet (or `Доходы` for income)
- `Expenses_RAW` (always)

Success message is sent only if both Excel writes succeed.

## 1) Install

```bash
npm install
```

## 2) Configure environment

```bash
cp .env.example .env
```

Fill in:

- `TELEGRAM_BOT_TOKEN`
- `OPENAI_API_KEY`
- `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`
- `ONEDRIVE_DRIVE_ID`
- `EXCEL_FILE_ID`

## 3) Excel workbook requirement

This code appends rows through Microsoft Graph table API, so each destination sheet must contain a table.

Expected table naming rule:

- `<EXCEL_TABLE_NAME_PREFIX><sheetName>`

With default `EXCEL_TABLE_NAME_PREFIX=tbl_`, create tables named:

- `tbl_Доходы`
- `tbl_Shopping`
- `tbl_Транспорт`
- `tbl_Utilities`
- `tbl_Развлечения`
- `tbl_Рестораны`
- `tbl_Семья и персонал`
- `tbl_Расходы Персонал`
- `tbl_Прочее`
- `tbl_Связь и подписки`
- `tbl_Путешествия`
- `tbl_Здоровье`
- `tbl_Образование`
- `tbl_Аренда`
- `tbl_Мед. страховка`
- `tbl_CAPEX`
- `tbl_Credit Cards`
- `tbl_Expenses_RAW`

## 4) Run

```bash
npm start
```

Server endpoints:

- `POST /webhook/telegram` (Telegram webhook target)
- `GET /health`

## 5) Telegram webhook setup

Expose local server (for example with ngrok), then set webhook:

```bash
curl -X POST "https://api.telegram.org/bot<TELEGRAM_BOT_TOKEN>/setWebhook" \
  -d "url=https://<your-public-url>/webhook/telegram"
```

## Behavior implemented

- Sends immediate placeholder: `⏳ Logging your transaction...`
- Parses message with OpenAI into strict JSON:
  - `amount`, `description`, `category`, `type`
- Routes:
  - `type=income` -> `Доходы`
  - otherwise -> selected expense category sheet
- Always logs to `Expenses_RAW`
- Sends:
  - success: `✅ Logged: ...` only after both writes succeed
  - failure: `❌ Error saving to Excel...` on parse/Excel errors
