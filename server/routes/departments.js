const express = require('express');
const { pool } = require('../db');
const { requireAuth } = require('../middleware/auth');
const asyncHandler = require('../asyncHandler');

const router = express.Router();

// Admin-only — no staff-portal caller.
router.use(requireAuth);

// GET /api/departments
router.get('/', asyncHandler(async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM departments ORDER BY name');
  res.json({ status: 'ok', data: rows });
}));

// POST /api/departments — create
router.post('/', asyncHandler(async (req, res) => {
  const { name, head_staff_code } = req.body || {};
  if (!name) return res.status(400).json({ status: 'error', msg: 'name required' });
  const { rows } = await pool.query(
    'INSERT INTO departments (name, head_staff_code) VALUES ($1, $2) RETURNING id',
    [name, head_staff_code || null]
  );
  res.json({ status: 'ok', id: rows[0].id });
}));

// PUT /api/departments/:id
router.put('/:id', asyncHandler(async (req, res) => {
  const { id } = req.params;
  const { name, head_staff_code } = req.body || {};
  const setCols = [];
  const values = [id];
  if (name !== undefined)            { values.push(name);            setCols.push(`name = $${values.length}`); }
  if (head_staff_code !== undefined) { values.push(head_staff_code); setCols.push(`head_staff_code = $${values.length}`); }
  if (!setCols.length) return res.status(400).json({ status: 'error', msg: 'No fields to update' });

  const result = await pool.query(`UPDATE departments SET ${setCols.join(', ')} WHERE id = $1`, values);
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'Department not found: ' + id });
  res.json({ status: 'ok' });
}));

// DELETE /api/departments/:id
router.delete('/:id', asyncHandler(async (req, res) => {
  const result = await pool.query('DELETE FROM departments WHERE id = $1', [req.params.id]);
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'Department not found: ' + req.params.id });
  res.json({ status: 'ok' });
}));

module.exports = router;
