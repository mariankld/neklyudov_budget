require("dotenv").config();
const fxRates = require("../src/fxRates");

// CLI helper for manually-typed Excel rows: prints the frozen historical HKD rate for a
// given currency + date, the same number the bot would have computed automatically had it
// logged the transaction. Use this to fill in Курс / Сумма (HKD) by hand when adding a row
// directly in Excel instead of through Telegram.
//
// Usage:
//   node scripts/lookup-historical-rate.js USD 15/03/2024
//   node scripts/lookup-historical-rate.js RUB 15/03/2024 5000   (also prints Sum (HKD) for that amount)

function roundMoney(n) {
  return Math.round(n * 100) / 100;
}

(async () => {
  const [, , currencyArg, dateArg, amountArg] = process.argv;

  if (!currencyArg || !dateArg) {
    console.error("Usage: node scripts/lookup-historical-rate.js <CURRENCY> <dd/mm/yyyy> [amount]");
    console.error(`Tracked currencies: ${fxRates.TRACKED_CURRENCIES.join(", ")}`);
    process.exit(1);
  }

  const currency = currencyArg.trim().toUpperCase();
  const date = dateArg.trim();

  if (!fxRates.TRACKED_CURRENCIES.includes(currency)) {
    console.error(`"${currency}" is not tracked. Tracked currencies: ${fxRates.TRACKED_CURRENCIES.join(", ")}`);
    process.exit(1);
  }

  try {
    const rate = await fxRates.getRateToHkd(currency, date);
    console.log(`Курс (${currency} → HKD) on ${date}: ${rate}`);

    if (amountArg) {
      const amount = Number(amountArg);
      if (!Number.isFinite(amount)) {
        console.error(`"${amountArg}" is not a valid amount.`);
        process.exit(1);
      }
      console.log(`Сумма (HKD) for ${amount} ${currency}: ${roundMoney(amount * rate)}`);
    }
  } catch (err) {
    console.error(`Lookup failed: ${err.message}`);
    process.exit(1);
  }
})();
