const express = require('express');
const { pool } = require('../db');
const { requireAuth } = require('../middleware/auth');
const asyncHandler = require('../asyncHandler');

const router = express.Router();

const COLUMNS = [
  'request_by', 'staff_code', 'staff_name', 'date_evaluate',
  'kpi_score', 'previous_salary', 'current_salary', 'approved_by', 'remark',
];

// GET /api/evaluations?staffCode= — public, no auth (mirrors the old read?sheet=EvaluateStaff,
// which had no server-side auth either).
router.get('/', asyncHandler(async (req, res) => {
  const { staffCode } = req.query;
  const clauses = [];
  const params = [];
  if (staffCode) { params.push(staffCode); clauses.push(`staff_code = $${params.length}`); }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  const { rows } = await pool.query(
    `SELECT * FROM staff_evaluations ${where} ORDER BY date_evaluate DESC, id DESC`,
    params
  );
  res.json({ status: 'ok', data: rows });
}));

router.use(requireAuth);

// POST /api/evaluations — create. Admin-only (hr-system.html's Evaluate Staff page),
// unlike Leave/OT/Food/Comment/Workplace which have a staff-portal self-service caller.
router.post('/', asyncHandler(async (req, res) => {
  const row = req.body || {};
  const values = COLUMNS.map((c) => row[c] ?? null);
  const placeholders = COLUMNS.map((_, i) => `$${i + 1}`).join(', ');
  const { rows } = await pool.query(
    `INSERT INTO staff_evaluations (${COLUMNS.join(', ')}) VALUES (${placeholders}) RETURNING id`,
    values
  );
  res.json({ status: 'ok', id: rows[0].id });
}));

// PUT /api/evaluations/:id — partial update by numeric id (real PK — replaces the old
// client-generated EvalNo matching, which was collision-prone: 'EVAL-' + array.length)
router.put('/:id', asyncHandler(async (req, res) => {
  const { id } = req.params;
  const row = req.body || {};
  const setCols = COLUMNS.filter((c) => row[c] !== undefined);
  if (setCols.length === 0) return res.status(400).json({ status: 'error', msg: 'No fields to update' });

  const setClause = setCols.map((c, i) => `${c} = $${i + 2}`).join(', ');
  const values = [id, ...setCols.map((c) => row[c])];
  const result = await pool.query(`UPDATE staff_evaluations SET ${setClause} WHERE id = $1`, values);
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'Evaluation not found: ' + id });
  res.json({ status: 'ok' });
}));

// DELETE /api/evaluations/:id
router.delete('/:id', asyncHandler(async (req, res) => {
  const result = await pool.query('DELETE FROM staff_evaluations WHERE id = $1', [req.params.id]);
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'Evaluation not found: ' + req.params.id });
  res.json({ status: 'ok' });
}));

module.exports = router;
