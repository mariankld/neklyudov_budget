// Kept in sync by hand with the BASE_PAYMENT_METHODS constants in src/index.js. Duplicated here
// (rather than required from src/index.js) so the audit/fix scripts don't have to boot the
// Express app / Telegram webhook wiring just to read a fixed list.
//
// Fixed, enumerated set of Payment Method values. Everything written to the workbook must be
// exactly one of these strings — no free text — so Summary tab SUMIF/SUMIFS formulas that match
// on Payment Method (e.g. "*Visa BOC*") never silently miss a row because of a syntax variant
// like "VISA" or "Boc Visa". OpenAI is constrained to this list at parse time; when it can't tell
// which one applies, the bot asks the user to pick via Telegram buttons before logging.
// Octopus, PayMe, FPS, QR code, and WeChat Pay are valid payment methods but are intentionally
// excluded from CREDIT_CARD_PAYMENT_METHODS — they must never count toward the Summary!H20
// credit-card total. src/index.js also layers a runtime-learned custom list on top of this base
// set (see getPaymentMethods() there); this file only mirrors the fixed base set since
// fix-h20-credit-card-total.js only ever needs CREDIT_CARD_PAYMENT_METHODS, which never changes
// at runtime.
//
// "МИР Сбер" added 2026-08-21 (Mariya) — also the same card list used by the forced Карта picker
// for the Credit Cards category (see CREDIT_CARDS_CATEGORY / buildCardKeyboard in src/index.js).
const CREDIT_CARD_PAYMENT_METHODS = ["Master SC", "Visa BOC", "Master BEA", "Master Citic", "МИР Сбер"];
const PAYMENT_METHODS = [
  "Cash",
  "Master SC",
  "Visa BOC",
  "Master BEA",
  "Master Citic",
  "МИР Сбер",
  "Bank Transfer",
  "Octopus",
  "PayMe",
  "FPS",
  "QR code",
  "WeChat Pay",
];
const PAYMENT_METHOD_UNCLEAR = "Unclear";

module.exports = {
  CREDIT_CARD_PAYMENT_METHODS,
  PAYMENT_METHODS,
  PAYMENT_METHOD_UNCLEAR,
};
