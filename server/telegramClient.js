const TELEGRAM_TOKEN = process.env.TELEGRAM_TOKEN || '';

async function sendTelegramMessage(chatId, text) {
  if (!TELEGRAM_TOKEN || !chatId) return;
  try {
    await fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ chat_id: chatId, text }),
    });
  } catch (err) {
    console.warn('[telegram] sendMessage failed:', err.message);
  }
}

module.exports = { sendTelegramMessage };
