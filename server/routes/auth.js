const express = require('express');
const crypto = require('crypto');
const { pool } = require('../db');
const { requireAuth, SESSION_TTL_MS } = require('../middleware/auth');
const asyncHandler = require('../asyncHandler');
const { sendTelegramMessage } = require('../telegramClient');

const router = express.Router();

function newToken() {
  return crypto.randomBytes(24).toString('hex');
}

// POST /api/auth/login
router.post('/login', asyncHandler(async (req, res) => {
  const username = String(req.body.username || '').trim();
  const password = String(req.body.password || '').trim();
  if (!username || !password) return res.status(400).json({ status: 'error', msg: 'Username and password required' });

  const { rows } = await pool.query('SELECT * FROM users WHERE username = $1', [username]);
  const user = rows[0];
  // Phase 5 swaps this for bcrypt.compare against password_hash; Phase 2 keeps the
  // plaintext comparison the Sheets-era client used, on purpose (see migration plan).
  const passwordOk = !!user && user.password_plain === password;
  if (!passwordOk) return res.status(401).json({ status: 'error', msg: 'Invalid username or password' });

  const isAdmin = (user.role || '').toLowerCase() === 'admin';
  const now = new Date();

  if (!isAdmin) {
    const active = await pool.query(
      `SELECT created_at FROM user_sessions WHERE user_id = $1 AND revoked_at IS NULL AND expires_at > now() LIMIT 1`,
      [user.id]
    );
    if (active.rows[0]) {
      return res.status(409).json({
        status: 'error', code: 'CONCURRENT_SESSION',
        msg: 'This user is already logged in on another device.',
        since: active.rows[0].created_at,
      });
    }
  }

  // Clean up any stale sessions for this user, then issue a fresh one.
  await pool.query('UPDATE user_sessions SET revoked_at = now() WHERE user_id = $1 AND revoked_at IS NULL', [user.id]);
  const token = newToken();
  const expiresAt = new Date(now.getTime() + SESSION_TTL_MS);
  await pool.query(
    'INSERT INTO user_sessions (user_id, session_token, expires_at) VALUES ($1, $2, $3)',
    [user.id, token, expiresAt]
  );

  res.json({
    status: 'ok', token,
    user: { username: user.username, name: user.name, role: user.role, email: user.email, department: user.department, position: user.position },
  });
}));

// POST /api/auth/logout
router.post('/logout', requireAuth, asyncHandler(async (req, res) => {
  await pool.query('UPDATE user_sessions SET revoked_at = now() WHERE id = $1', [req.session.id]);
  res.json({ status: 'ok' });
}));

// GET /api/auth/session — heartbeat; requireAuth already rejects expired/invalid tokens.
router.get('/session', requireAuth, asyncHandler(async (req, res) => {
  const expiresAt = new Date(Date.now() + SESSION_TTL_MS);
  await pool.query('UPDATE user_sessions SET last_seen_at = now(), expires_at = $2 WHERE id = $1', [req.session.id, expiresAt]);
  res.json({ status: 'ok', active: true, user: req.user });
}));

function normalizePhone(v) {
  return String(v || '').replace(/[^0-9]/g, '');
}

async function findStaffByPhone(phone) {
  const pn0 = phone.replace(/^0+/, '');
  const { rows } = await pool.query(
    `SELECT * FROM staff WHERE regexp_replace(phone, '[^0-9]', '', 'g') = $1
        OR regexp_replace(phone, '[^0-9]', '', 'g') = $2
        OR right(regexp_replace(phone, '[^0-9]', '', 'g'), 9) = right($1, 9)
     LIMIT 1`,
    [phone, pn0]
  );
  return rows[0] || null;
}

// POST /api/auth/otp/send
router.post('/otp/send', asyncHandler(async (req, res) => {
  const phone = normalizePhone(req.body.phone);
  if (phone.length < 8) return res.status(400).json({ status: 'error', msg: 'Phone invalid' });

  const staff = await findStaffByPhone(phone);
  if (!staff) return res.status(404).json({ status: 'error', msg: `Phone ${phone} not found` });
  if (!staff.telegram_chat_id) return res.status(400).json({ status: 'error', msg: 'Not registered. Send /start to @lhb_system_bot' });

  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const expiresAt = new Date(Date.now() + 5 * 60 * 1000);
  await pool.query('INSERT INTO staff_otp (staff_code, otp_code, expires_at) VALUES ($1, $2, $3)', [staff.staff_code, otp, expiresAt]);
  await sendTelegramMessage(staff.telegram_chat_id, `LHB HR OTP\n\n${staff.name}\nCode: ${otp}\n\nExpire 5min`);
  res.json({ status: 'ok' });
}));

// POST /api/auth/otp/verify
router.post('/otp/verify', asyncHandler(async (req, res) => {
  const phone = normalizePhone(req.body.phone);
  const code = String(req.body.code || '').trim();

  const staff = await findStaffByPhone(phone);
  if (!staff) return res.status(404).json({ status: 'error', msg: 'Phone not found' });

  const { rows } = await pool.query(
    `SELECT * FROM staff_otp WHERE staff_code = $1 AND otp_code = $2 AND used_at IS NULL AND expires_at > now()
     ORDER BY created_at DESC LIMIT 1`,
    [staff.staff_code, code]
  );
  if (!rows[0]) return res.status(400).json({ status: 'error', msg: 'OTP incorrect or expired' });

  await pool.query('UPDATE staff_otp SET used_at = now() WHERE id = $1', [rows[0].id]);
  res.json({ status: 'ok', staff });
}));

module.exports = router;
