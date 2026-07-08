const express = require('express');
const { pool } = require('../db');
const { requireAuth } = require('../middleware/auth');
const asyncHandler = require('../asyncHandler');

const router = express.Router();

const COLUMNS = ['staff_code', 'type_of_leave', 'start_date', 'end_date', 'days', 'reason', 'status'];

// GET /api/leave?staffCode= — public, no auth (mirrors the old read?sheet=StaffLeave, which had
// no server-side auth either — staff-portal.html's leave-balance widget reads this unauthenticated
// exactly like it did against Apps Script).
router.get('/', asyncHandler(async (req, res) => {
  const { staffCode } = req.query;
  const clauses = [];
  const params = [];
  if (staffCode) { params.push(staffCode); clauses.push(`staff_code = $${params.length}`); }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  // Joined to staff for the display Name — the old StaffLeave sheet stored Name as a
  // literal duplicate column; this reflects the staff member's *current* name instead.
  const { rows } = await pool.query(
    `SELECT l.*, s.name AS staff_name
     FROM staff_leave l LEFT JOIN staff s ON s.staff_code = l.staff_code
     ${where.replace(/staff_code/g, 'l.staff_code')}
     ORDER BY l.start_date DESC, l.id DESC`,
    params
  );
  res.json({ status: 'ok', data: rows });
}));

router.use(requireAuth);

// POST /api/leave — create
router.post('/', asyncHandler(async (req, res) => {
  const row = req.body || {};
  if (!row.staff_code) return res.status(400).json({ status: 'error', msg: 'staff_code required' });

  const values = COLUMNS.map((c) => row[c] ?? null);
  const placeholders = COLUMNS.map((_, i) => `$${i + 1}`).join(', ');
  const { rows } = await pool.query(
    `INSERT INTO staff_leave (${COLUMNS.join(', ')}) VALUES (${placeholders}) RETURNING id`,
    values
  );
  res.json({ status: 'ok', id: rows[0].id });
}));

// PUT /api/leave/:id — partial update by numeric id (real PK — replaces the old ID+StartDate
// composite-key matching, which could silently hit the wrong row for a staff with 2+ leave records)
router.put('/:id', asyncHandler(async (req, res) => {
  const { id } = req.params;
  const row = req.body || {};
  const setCols = COLUMNS.filter((c) => c !== 'staff_code' && row[c] !== undefined);
  if (setCols.length === 0) return res.status(400).json({ status: 'error', msg: 'No fields to update' });

  const setClause = setCols.map((c, i) => `${c} = $${i + 2}`).join(', ');
  const values = [id, ...setCols.map((c) => row[c])];
  const result = await pool.query(
    `UPDATE staff_leave SET ${setClause}, updated_at = now() WHERE id = $1`,
    values
  );
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'Leave record not found: ' + id });
  res.json({ status: 'ok' });
}));

// DELETE /api/leave/:id
router.delete('/:id', asyncHandler(async (req, res) => {
  const result = await pool.query('DELETE FROM staff_leave WHERE id = $1', [req.params.id]);
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'Leave record not found: ' + req.params.id });
  res.json({ status: 'ok' });
}));

module.exports = router;
