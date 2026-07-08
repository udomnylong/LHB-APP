const express = require('express');
const { pool } = require('../db');
const { requireAuth } = require('../middleware/auth');
const asyncHandler = require('../asyncHandler');

const router = express.Router();

// time_from/time_to/reason exist in the table but are legacy/unused — the real UI
// (hr-system.html SHEET_HEADERS.StaffOT) only ever populates type_of_work + remark.
const COLUMNS = ['staff_code', 'date', 'hours', 'type_of_work', 'status', 'remark'];

// GET /api/ot?staffCode= — public, no auth (mirrors the old read?sheet=StaffOT, which had no
// server-side auth either).
router.get('/', asyncHandler(async (req, res) => {
  const { staffCode } = req.query;
  const clauses = [];
  const params = [];
  if (staffCode) { params.push(staffCode); clauses.push(`staff_code = $${params.length}`); }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  // Joined to staff for the display Name — the old StaffOT sheet stored Name as a
  // literal duplicate column; this reflects the staff member's *current* name instead.
  const { rows } = await pool.query(
    `SELECT o.*, s.name AS staff_name
     FROM staff_ot o LEFT JOIN staff s ON s.staff_code = o.staff_code
     ${where.replace(/staff_code/g, 'o.staff_code')}
     ORDER BY o.date DESC, o.id DESC`,
    params
  );
  res.json({ status: 'ok', data: rows });
}));

router.use(requireAuth);

// POST /api/ot — create
router.post('/', asyncHandler(async (req, res) => {
  const row = req.body || {};
  if (!row.staff_code) return res.status(400).json({ status: 'error', msg: 'staff_code required' });

  const values = COLUMNS.map((c) => row[c] ?? null);
  const placeholders = COLUMNS.map((_, i) => `$${i + 1}`).join(', ');
  const { rows } = await pool.query(
    `INSERT INTO staff_ot (${COLUMNS.join(', ')}) VALUES (${placeholders}) RETURNING id`,
    values
  );
  res.json({ status: 'ok', id: rows[0].id });
}));

// PUT /api/ot/:id — partial update by numeric id (real PK — replaces the old ID-only matching,
// which used the *staff* ID and could silently hit the wrong row for a staff with 2+ OT records)
router.put('/:id', asyncHandler(async (req, res) => {
  const { id } = req.params;
  const row = req.body || {};
  const setCols = COLUMNS.filter((c) => c !== 'staff_code' && row[c] !== undefined);
  if (setCols.length === 0) return res.status(400).json({ status: 'error', msg: 'No fields to update' });

  const setClause = setCols.map((c, i) => `${c} = $${i + 2}`).join(', ');
  const values = [id, ...setCols.map((c) => row[c])];
  const result = await pool.query(
    `UPDATE staff_ot SET ${setClause}, updated_at = now() WHERE id = $1`,
    values
  );
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'OT record not found: ' + id });
  res.json({ status: 'ok' });
}));

// DELETE /api/ot/:id
router.delete('/:id', asyncHandler(async (req, res) => {
  const result = await pool.query('DELETE FROM staff_ot WHERE id = $1', [req.params.id]);
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'OT record not found: ' + req.params.id });
  res.json({ status: 'ok' });
}));

module.exports = router;
