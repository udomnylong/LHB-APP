const express = require('express');
const { pool } = require('../db');
const { requireAuth } = require('../middleware/auth');
const asyncHandler = require('../asyncHandler');

const router = express.Router();

// Admin-only end to end (payroll deductions), unlike Leave/OT/Food/Comment/Workplace —
// there is no staff-portal self-service caller for this resource.
router.use(requireAuth);

// GET /api/penalty?staffCode=&month=
router.get('/', asyncHandler(async (req, res) => {
  const { staffCode, month } = req.query;
  const clauses = [];
  const params = [];
  if (staffCode) { params.push(staffCode); clauses.push(`staff_code = $${params.length}`); }
  if (month)     { params.push(month);     clauses.push(`month = $${params.length}`); }
  const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
  const { rows } = await pool.query(`SELECT * FROM payroll_penalties ${where}`, params);
  res.json({ status: 'ok', data: rows });
}));

// PUT /api/penalty — upsert by (staff_code, month)
router.put('/', asyncHandler(async (req, res) => {
  const { staff_code, month, amount } = req.body || {};
  if (!staff_code || !month) return res.status(400).json({ status: 'error', msg: 'staff_code and month required' });
  await pool.query(
    `INSERT INTO payroll_penalties (staff_code, month, amount) VALUES ($1, $2, $3)
     ON CONFLICT (staff_code, month) DO UPDATE SET amount = EXCLUDED.amount, updated_at = now()`,
    [staff_code, month, amount || 0]
  );
  res.json({ status: 'ok' });
}));

module.exports = router;
