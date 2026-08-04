const express = require('express');
const { pool } = require('../db');
const { requireAuth } = require('../middleware/auth');
const asyncHandler = require('../asyncHandler');

const router = express.Router();

// Admin-only — single-row config blob (late-grace/deduction rules, per-department
// shift schedules), no staff-portal caller.
router.use(requireAuth);

// GET /api/att-settings — returns the settings JSON blob, or null if never saved
// (the frontend falls back to its own hardcoded defaults in that case).
router.get('/', asyncHandler(async (req, res) => {
  const { rows } = await pool.query('SELECT settings FROM attendance_settings ORDER BY id LIMIT 1');
  res.json({ status: 'ok', data: rows[0] ? rows[0].settings : null });
}));

// PUT /api/att-settings — upsert the single settings row
router.put('/', asyncHandler(async (req, res) => {
  const settings = req.body || {};
  const { rows } = await pool.query('SELECT id FROM attendance_settings ORDER BY id LIMIT 1');
  if (rows[0]) {
    await pool.query('UPDATE attendance_settings SET settings = $1, updated_at = now() WHERE id = $2', [settings, rows[0].id]);
  } else {
    await pool.query('INSERT INTO attendance_settings (settings) VALUES ($1)', [settings]);
  }
  res.json({ status: 'ok' });
}));

module.exports = router;
