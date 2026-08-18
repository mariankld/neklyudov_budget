# Neklyudov Budget Logger

Telegram → OpenAI → **Excel (SharePoint, via Microsoft Graph)**. Parsed transactions are appended to:

- `RAW` (always) — the full log of every transaction, income or expense
- the matching category Excel Table (expenses only; income has no dedicated table)

Categories are hard-coded in `CATEGORY_TABLE_MAP` (`src/index.js`), mapping each category label to its real Excel Table name in the workbook. `Income` is RAW-only.

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
- `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` (Entra ID app registration, app-only Graph access)
- `EXCEL_SHARE_URL` (SharePoint share link to the workbook)
- `EXCEL_DRIVE_ID`, `EXCEL_ITEM_ID` (resolved once via `npm run resolve-excel-share`, see below)

Set `INCOME_SHEET_NAME` and `RAW_SHEET_NAME` to match the workbook's `RAW` table name and income label.

## 3) Resolve the Excel share link (one-time)

```bash
npm run resolve-excel-share
```

Resolves `EXCEL_SHARE_URL` to a `driveId`/`itemId` pair and lists every Excel Table found in the workbook — paste the printed `EXCEL_DRIVE_ID` / `EXCEL_ITEM_ID` into `.env`.

Requires the Entra ID app registration to have admin-consented `Files.ReadWrite.All` (or `Sites.ReadWrite.All`) application permission.

## 4) Inspect the workbook structure (optional)

```bash
npm run inspect-excel-structure
```

Lists every worksheet, table, and column — useful for confirming `CATEGORY_TABLE_MAP` and column order in `src/index.js` still match the live workbook.

## 5) Run

```bash
npm start
```

On startup the server validates required env vars, then calls `listWorkbookTables` to confirm `RAW` and every table in `CATEGORY_TABLE_MAP` actually exist in the live workbook — it refuses to accept Telegram traffic if any are missing.

Endpoints:

- `POST /webhook/telegram` (Telegram webhook target)
- `GET /health`

## 6) Telegram webhook

On Railway, `RAILWAY_PUBLIC_DOMAIN` is picked up automatically. Otherwise expose the server (e.g. with ngrok) and either let `ensureTelegramWebhook()` register it via `TELEGRAM_WEBHOOK_BASE_URL` / `PUBLIC_URL`, or set it manually:

```bash
curl -X POST "https://api.telegram.org/bot<TELEGRAM_BOT_TOKEN>/setWebhook" \
  -d "url=https://<your-public-url>/webhook/telegram"
```

## Photos (receipts / payment screenshots)

Send a photo instead of text (with or without a caption) and the bot reads it the same way:

- Compressed Telegram photos and uncompressed image documents (image/* mime type) are both accepted.
- The highest-resolution version Telegram provides is downloaded and sent to OpenAI as an image input — no separate OCR step.
- Works for printed/handwritten receipts and screenshots such as iPhone Wallet payment confirmations.
- The bot never guesses the amount. It reads as carefully as it can, and only if there is truly no readable number does it reply asking you to enter the amount and description manually — it will not default to 0 or invent a figure.
- Optional `OPENAI_VISION_MODEL` env var to use a different model for images; defaults to `gpt-4o` (better than `gpt-4o-mini` at reading small/blurry text).
- The same Yes/Edit confirmation flow applies before anything is written to Excel.

## Row shape appended by the app

**`RAW` table (12 columns, always written):** Date (dd/mm/yyyy), Type (Income/Expense), Category, Subcategory, Description, Location, Amount, Currency, Sum (HKD) (blank — Excel formula), Notes, Sender, Payment Method.

**Category table (17 columns, expenses only):** Дата, Категория, Описание, Локация, Сумма, Валюта, Курс (blank — Excel formula, exchange-rate lookup), Сумма (HKD) (blank — Excel formula), Примечание, Метод оплаты, Получатель/Сотрудник, Страховка (blank, not tracked by the bot), Статус выплаты (blank), Пользователь, Страховая компания (blank), Период покрытия (blank), Карта (blank).

Category tables' **Категория** column is filled from RAW's **Subcategory** (the specific line item), not RAW's Category — confirmed with Mariya, since it gives more precision inside a tab that's already scoped to one category.

RAW and category-table column orders are independent: `appendTableRow` is called once per table with a values array built specifically for that table's real column order (named-field mapping, not positional copy-paste), so the two never need to line up.

**Sender:** the Telegram `@username` of whoever sent the message/photo (falls back to first name, then Telegram user id). Shown in the confirmation preview as "Logged by", stored in RAW's Sender column and the category table's Пользователь column.

**Recipient:** if the message names a specific person the payment was made to or received from (e.g. a staff member), OpenAI extracts it into the category table's Получатель/Сотрудник column. Empty string if not mentioned — never invented.

**Payment method:** how the transaction was paid, read from the message text or the photo (e.g. "BOC VISA Infinite", "Cash", "HSBC transfer"). The model is instructed to only report what it can actually read/infer — it uses "Unknown" rather than guessing a bank or card it can't confirm.

## Behavior

- Placeholder: `⏳ Analyzing your message...` / `⏳ Reading your photo...`
- OpenAI returns strict JSON: `amount`, `description`, `category`, `subcategory`, `type`, `location`, `currency`, `notes`, `paymentMethod`, `recipient`
- `type=income` → RAW only (category forced to `INCOME_SHEET_NAME`)
- `type=expense` → RAW + the mapped category Excel Table
- RAW is written first. If RAW succeeds but the category-table write fails, the error is tagged `rawSucceeded=true`: the bot does **not** retry (that would duplicate the RAW row) and instead warns in Telegram that the category table needs a manual check.
- Fully successful only after both writes succeed (income: after the one RAW write succeeds).

## Category → Excel Table mapping

Defined in `CATEGORY_TABLE_MAP` in `src/index.js`:

| Category (RAW / OpenAI label) | Excel Table |
| --- | --- |
| Credit Cards | CreditCards |
| Shopping | Shopping |
| Transportation | TransportTable |
| Utilities | Utilities |
| Entertainment | Entertainment |
| Restaurants | Restaurants |
| Family and Staff | FamilyStaff |
| Personal Spending | StaffExpenses |
| Other | Other |
| Subscriptions | TelecomSubscriptions |
| Travel | Travel |
| Health | Health |
| Education | Education |
| Rent | Rent |
| Insurance | MedInsurance |
| CAPEX | CAPEX |

`Income` has no entry — it's RAW-only.
