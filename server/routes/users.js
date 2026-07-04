const express = require('express');
const { pool } = require('../db');
const asyncHandler = require('../asyncHandler');

const router = express.Router();

// GET /api/users — HR system account list for the admin Settings screen.
// No auth required, matching the old Apps Script doGet (fully unauthenticated) —
// and this is actually *safer* than that: it omits password_plain/password_hash,
// which the old Sheets-era client used to fetch in plaintext for every user just
// to do local login matching (no longer needed now that /api/auth/login does it
// server-side).
router.get('/', asyncHandler(async (req, res) => {
  const { rows } = await pool.query(
    'SELECT username, name, role, email, department, position FROM users ORDER BY username'
  );
  res.json({ status: 'ok', data: rows });
}));

module.exports = router;
