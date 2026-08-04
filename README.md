# Neklyudov Budget Logger

Telegram → OpenAI → **Google Sheets**. Parsed transactions are appended to:

- the tab for the chosen category (or the income tab for `type=income`)
- the RAW log tab (always)

Categories are **not** hard-coded: on startup the app reads all sheet tab titles from your spreadsheet, skips the RAW tab and any names in `GOOGLE_SKIP_TABS`, and uses the rest for OpenAI’s category enum.

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
- `GOOGLE_SPREADSHEET_ID` (from the spreadsheet URL)
- `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`
- `GOOGLE_REFRESH_TOKEN` (see below)

Set `INCOME_SHEET_NAME` and `RAW_SHEET_NAME` to match your workbook tab names exactly.

## 3) Google OAuth (refresh token)

```bash
npm run google-oauth-setup
```

Open the printed URL, approve access, then copy the printed `GOOGLE_REFRESH_TOKEN` into `.env`.

The redirect URI in Google Cloud Console must match `GOOGLE_OAUTH_REDIRECT_URI` (default `http://127.0.0.1:3001/oauth/callback`).

## 4) List tabs (optional)

```bash
npm run list-sheet-tabs
```

Use this to confirm tab names before setting `INCOME_SHEET_NAME`, `RAW_SHEET_NAME`, and `GOOGLE_SKIP_TABS`.

## 5) Run

```bash
npm start
```

On startup the server logs the resolved category list and prints row 1 of the RAW tab and one sample category tab (for column alignment checks).

Endpoints:

- `POST /webhook/telegram` (Telegram webhook target)
- `GET /health`

## 6) Telegram webhook

Expose the server (for example with ngrok), then:

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
- The same Yes/Edit confirmation flow applies before anything is written to the sheet.

## Row shape appended by the app

**Category tab:** date (ISO), time, amount, type, category name, description.

**RAW tab:** date (dd/mm/yyyy), type (Income/Expense), category, subcategory, description, location, amount, currency, sum in HKD (blank — calculated by a sheet formula), notes, sender (Telegram @tag or name), payment method.

Align headers in the sheet with these columns or adjust the code if your layout differs.

**Sender:** the Telegram `@username` of whoever sent the message/photo (falls back to first name, then Telegram user id, if no username is set). Shown in the confirmation preview as "Logged by" and stored in every RAW row.

**Payment method:** how the transaction was paid, read from the message text or the photo (e.g. "BOC VISA Infinite", "Cash", "HSBC transfer"). The model is instructed to only report what it can actually read/infer — it uses "Unknown" rather than guessing a bank or card it can't confirm.

## Behavior

- Placeholder: `⏳ Logging your transaction...`
- OpenAI returns strict JSON: `amount`, `description`, `category`, `type`
- `type=income` → row goes to `INCOME_SHEET_NAME`
- `type=expense` → row goes to the selected category tab
- Always appends to `RAW_SHEET_NAME`
- Success only after both writes succeed
