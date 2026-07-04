const express = require('express');
const { pool } = require('../db');
const { requireAuth } = require('../middleware/auth');
const asyncHandler = require('../asyncHandler');

const router = express.Router();

const COLUMNS = [
  'staff_code', 'name', 'name_latin', 'sex', 'lv', 'position', 'department', 'project_name',
  'date_of_birth', 'starting_date', 'resign_date', 'salary', 'gmail', 'bank_name', 'bank_number',
  'photo_url', 'phone', 'employment_status', 'telegram_chat_id',
];

// GET /api/staff — full directory (mirrors the old read?sheet=StaffInfo dump, which
// had no server-side auth either — staff-portal.html's Gmail+phone login matches
// client-side against this same full list before any session exists). Writes below
// require a bearer token; that's new hardening, not a behavior the old system had.
router.get('/', asyncHandler(async (req, res) => {
  const { rows } = await pool.query(`SELECT ${COLUMNS.join(', ')} FROM staff ORDER BY staff_code`);
  res.json({ status: 'ok', data: rows });
}));

router.use(requireAuth);

// POST /api/staff — create
router.post('/', asyncHandler(async (req, res) => {
  const row = req.body || {};
  if (!row.staff_code) return res.status(400).json({ status: 'error', msg: 'staff_code required' });

  const values = COLUMNS.map((c) => row[c] ?? null);
  const placeholders = COLUMNS.map((_, i) => `$${i + 1}`).join(', ');
  await pool.query(
    `INSERT INTO staff (${COLUMNS.join(', ')}) VALUES (${placeholders})`,
    values
  );
  res.json({ status: 'ok' });
}));

// PUT /api/staff/:staffCode — partial update
router.put('/:staffCode', asyncHandler(async (req, res) => {
  const { staffCode } = req.params;
  const row = req.body || {};
  const setCols = COLUMNS.filter((c) => c !== 'staff_code' && row[c] !== undefined);
  if (setCols.length === 0) return res.status(400).json({ status: 'error', msg: 'No fields to update' });

  const setClause = setCols.map((c, i) => `${c} = $${i + 2}`).join(', ');
  const values = [staffCode, ...setCols.map((c) => row[c])];
  const result = await pool.query(
    `UPDATE staff SET ${setClause}, updated_at = now() WHERE staff_code = $1`,
    values
  );
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'Staff not found: ' + staffCode });
  res.json({ status: 'ok' });
}));

// DELETE /api/staff/:staffCode
router.delete('/:staffCode', asyncHandler(async (req, res) => {
  const result = await pool.query('DELETE FROM staff WHERE staff_code = $1', [req.params.staffCode]);
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'Staff not found: ' + req.params.staffCode });
  res.json({ status: 'ok' });
}));

module.exports = router;
