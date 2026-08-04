const express = require('express');
const { pool } = require('../db');
const { requireAuth } = require('../middleware/auth');
const asyncHandler = require('../asyncHandler');

const router = express.Router();

const COLUMNS = ['date', 'time', 'staff_code', 'department', 'project_name', 'comment', 'photo_url', 'status'];

// GET /api/comments?staffCode=&date=&from=&to= — public, no auth (mirrors the old
// read?sheet=Comment, which had no server-side auth either).
router.get('/', asyncHandler(async (req, res) => {
  const { staffCode, date, from, to } = req.query;
  const clauses = [];
  const params = [];
  if (staffCode) { params.push(staffCode); clauses.push(`staff_code = $${params.length}`); }
  if (date)      { params.push(date);      clauses.push(`date = $${params.length}`); }
  if (from)      { params.push(from);      clauses.push(`date >= $${params.length}`); }
  if (to)        { params.push(to);        clauses.push(`date <= $${params.length}`); }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  const { rows } = await pool.query(
    `SELECT * FROM comments ${where} ORDER BY date DESC, id DESC`,
    params
  );
  res.json({ status: 'ok', data: rows });
}));

// POST /api/comments — create. Public, no auth: staff-portal.html's submitReport()
// (the primary caller) has no session/token concept at all, same reasoning as
// attendance.js's /checkins /checkouts.
router.post('/', asyncHandler(async (req, res) => {
  const row = req.body || {};
  const values = COLUMNS.map((c) => row[c] ?? null);
  const placeholders = COLUMNS.map((_, i) => `$${i + 1}`).join(', ');
  const { rows } = await pool.query(
    `INSERT INTO comments (${COLUMNS.join(', ')}) VALUES (${placeholders}) RETURNING id`,
    values
  );
  res.json({ status: 'ok', id: rows[0].id });
}));

router.use(requireAuth);

// PUT /api/comments/:id
router.put('/:id', asyncHandler(async (req, res) => {
  const { id } = req.params;
  const row = req.body || {};
  const setCols = COLUMNS.filter((c) => row[c] !== undefined);
  if (setCols.length === 0) return res.status(400).json({ status: 'error', msg: 'No fields to update' });

  const setClause = setCols.map((c, i) => `${c} = $${i + 2}`).join(', ');
  const values = [id, ...setCols.map((c) => row[c])];
  const result = await pool.query(`UPDATE comments SET ${setClause} WHERE id = $1`, values);
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'Comment not found: ' + id });
  res.json({ status: 'ok' });
}));

// DELETE /api/comments/:id
router.delete('/:id', asyncHandler(async (req, res) => {
  const result = await pool.query('DELETE FROM comments WHERE id = $1', [req.params.id]);
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'Comment not found: ' + req.params.id });
  res.json({ status: 'ok' });
}));

module.exports = router;
