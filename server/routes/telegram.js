const express = require('express');
const { pool } = require('../db');
const asyncHandler = require('../asyncHandler');
const { sendTelegramMessage } = require('../telegramClient');

const router = express.Router();

// De-dupe Telegram's at-least-once delivery retries for a while.
const seenUpdateIds = new Map();
const SEEN_TTL_MS = 6 * 3600 * 1000;
function alreadySeen(updateId) {
  const now = Date.now();
  for (const [id, ts] of seenUpdateIds) if (now - ts > SEEN_TTL_MS) seenUpdateIds.delete(id);
  if (seenUpdateIds.has(updateId)) return true;
  seenUpdateIds.set(updateId, now);
  return false;
}

async function registerStaff(phone, chatId) {
  const pn0 = phone.replace(/^0+/, '');
  const { rows } = await pool.query(
    `SELECT * FROM staff WHERE regexp_replace(phone, '[^0-9]', '', 'g') = $1
        OR regexp_replace(phone, '[^0-9]', '', 'g') = $2
        OR right(regexp_replace(phone, '[^0-9]', '', 'g'), 9) = right($1, 9)
     LIMIT 1`,
    [phone, pn0]
  );
  const staff = rows[0];
  if (!staff) return `Phone 0${pn0} not found. Contact Admin.`;
  await pool.query('UPDATE staff SET telegram_chat_id = $1, updated_at = now() WHERE staff_code = $2', [chatId, staff.staff_code]);
  return `Register OK!\n${staff.name} (${staff.staff_code})\nPhone: 0${pn0}`;
}

async function getRegisteredInfo(chatId) {
  const { rows } = await pool.query('SELECT * FROM staff WHERE telegram_chat_id = $1 LIMIT 1', [chatId]);
  if (!rows[0]) return null;
  return `Registered: ${rows[0].name} (${rows[0].staff_code})`;
}

// POST /api/telegram/webhook
router.post('/webhook', asyncHandler(async (req, res) => {
  res.json({ ok: true }); // ack immediately, Telegram doesn't need to wait on our processing

  const body = req.body || {};
  if (body.update_id != null && alreadySeen(String(body.update_id))) return;

  const msg = body.message || body.edited_message;
  if (!msg) return;
  if (Math.floor(Date.now() / 1000) - (msg.date || 0) > 30) return; // ignore stale updates

  const chatId = String(msg.chat.id);
  const text = (msg.text || '').trim();

  if (text === '/start') {
    return sendTelegramMessage(chatId, 'LHB HR Bot!\n\nRegister:\n/register 0XXXXXXXXX');
  }
  if (text.startsWith('/register')) {
    const phone = text.replace('/register', '').trim().replace(/[^0-9]/g, '');
    if (!phone || phone.length < 8) return sendTelegramMessage(chatId, 'Format: /register 0XXXXXXXXX');
    return sendTelegramMessage(chatId, await registerStaff(phone, chatId));
  }
  if (text === '/status') {
    return sendTelegramMessage(chatId, (await getRegisteredInfo(chatId)) || 'Not registered.');
  }
}));

module.exports = router;
