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

**`RAW` table (13 columns, always written):** Date (dd/mm/yyyy), Type (Income/Expense), Category, Subcategory, Description, Location, Amount, Currency, Exchange Rate, Sum (HKD), Notes, Sender, Payment Method.

**Category table (17 columns, expenses only):** Дата, Категория, Описание, Локация, Сумма, Валюта, Курс, Сумма (HKD), Примечание, Метод оплаты, Получатель/Сотрудник, Страховка (blank, not tracked by the bot), Статус выплаты (blank), Пользователь, Страховая компания (blank), Период покрытия (blank), Карта (blank).

**Exchange Rate / Курс and Sum (HKD):** both are bot-computed static numbers, not Excel formulas — frozen to the transaction's own date via `fxRates.getRateToHkd(currency, date)` (historical rate, not "today's" rate) at write time, so they never silently change if a later day's rate updates. `RAW`'s Exchange Rate column sits immediately after Currency; `syncJob.js`'s bidirectional sync keeps it in step with the category table's Курс whenever either side is edited.

> This 17-column shape is now the exception, not the rule — see `src/syncJob.js`'s `STANDARD_CAT_COLS` / `INSURANCE_CAT_COLS` / `CREDITCARDS_CAT_COLS` for each table's actual current layout: most tables are the standard 12 columns (+ RowID/SyncHash), Health/MedInsurance kept all 5 insurance columns (19 total), and CreditCards keeps only Карта (15 total).

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
| Staff Spending | StaffExpenses |
| Staff Advances | StaffAdvances |
| Other | Other |
| Subscriptions | TelecomSubscriptions |
| Travel | Travel |
| Health | Health |
| Education | Education |
| Rent | Rent |
| Insurance | MedInsurance |
| CAPEX | CAPEX |

`Income` has no entry — it's RAW-only.

## Commands

Sent as plain Telegram messages, checked before the normal expense-parsing flow:

| Command | Effect |
| --- | --- |
| `/cancel` | Cancels a pending edit. |
| `/ignore <text>` | Silent no-op — lets you leave a comment in the chat without it being parsed as an expense. |
| `/sync` | Runs the daily maintenance jobs (CurrencyRates refresh + RAW ↔ category-table sync) immediately, useful right after editing a row directly in Excel. |
| `/staff` | Replies with each helper's (Marietta, Yaya) remaining balance, and the amount/date of their most recent advance. See below. |

## Staff Advances / Staff Spending recipient

Every row logged under **Staff Advances** or **Staff Spending** must have `Получатель/Сотрудник` set to exactly one of **Marietta** or **Yaya** — this is enforced in `src/index.js` (`STAFF_MEMBERS`, `findKnownStaffMember`), so the balance math below is never thrown off by a typo or a missing name.

- If OpenAI can't confidently extract one of the two names from the message, the bot does **not** guess — the normal Yes/Edit confirmation is replaced with a forced 2-button picker (`buildStaffRecipientKeyboard`) and the entry can't be logged until one is tapped.
- Recipient casing is normalized on every draft (e.g. "yaya" → "Yaya") so downstream `SUMIF`/`SUMIFS`/`MAXIFS` matching, in both the `/staff` command and the live Excel formulas below, never misses a row.

## Staff balance summary (`/staff` + live Excel formulas)

The remaining balance for each helper is available two ways, and both use the same logic so they always agree:

1. **The `/staff` Telegram command** — computed on demand in `getStaffBalancesReport()` (`src/index.js`) from the live `StaffAdvances`/`StaffExpenses` tables.
2. **A live formula block** in columns **Q:V** of the "Staff Advances" worksheet, one row per helper, so the numbers are visible directly in Excel without needing the bot. Set it up once with:

   ```
   npm run add-staff-balance-formulas -- --apply
   ```

   (run without `--apply` first for a dry run). This writes:

   | Col | Header | Formula |
   | --- | --- | --- |
   | Q | Staff | `Marietta` / `Yaya` (plain values) |
   | R | Total Advances (HKD) | `SUMIF` over `StaffAdvances` |
   | S | Total Spent (HKD) | `SUMIF` over `StaffExpenses` |
   | T | Remaining (HKD) | `R - S` |
   | U | Last Advance Amount (HKD) | `SUMIFS` — every `StaffAdvances` row dated exactly the date in V |
   | V | Last Advance Date | `MAXIFS` — the latest date among that helper's `StaffAdvances` rows |

Since an advance can arrive as several receipts, **each distinct date is treated as one advance**: rows for the same helper logged on the same day are summed together (U), but different dates are separate advances — only the most recent date's total is shown as "last advance". `StaffAdvances[Дата]` stores real Excel date serials (see `scripts/fix-staff-advances-dates.js`), so `MAXIFS` works on it directly; the script also copies the worksheet's existing date `numberFormat` onto V2:V3 so it displays as a date rather than a raw serial number.
