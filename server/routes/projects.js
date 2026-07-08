const express = require('express');
const { pool } = require('../db');
const { requireAuth } = require('../middleware/auth');
const asyncHandler = require('../asyncHandler');

const router = express.Router();

const COLUMNS = ['project_id', 'project_name', 'location', 'latitude', 'longitude', 'radius', 'status'];

// GET /api/projects — public, no auth. staff-portal.html reads this unauthenticated for GPS
// geofence data during check-in/out, exactly like it did against the old read?sheet=Project.
router.get('/', asyncHandler(async (req, res) => {
  const { rows } = await pool.query(`SELECT ${COLUMNS.join(', ')} FROM projects ORDER BY project_id`);
  res.json({ status: 'ok', data: rows });
}));

router.use(requireAuth);

// POST /api/projects — create
router.post('/', asyncHandler(async (req, res) => {
  const row = req.body || {};
  if (!row.project_id) return res.status(400).json({ status: 'error', msg: 'project_id required' });

  const values = COLUMNS.map((c) => row[c] ?? null);
  const placeholders = COLUMNS.map((_, i) => `$${i + 1}`).join(', ');
  await pool.query(
    `INSERT INTO projects (${COLUMNS.join(', ')}) VALUES (${placeholders})`,
    values
  );
  res.json({ status: 'ok' });
}));

// PUT /api/projects/:projectId — partial update
router.put('/:projectId', asyncHandler(async (req, res) => {
  const { projectId } = req.params;
  const row = req.body || {};
  const setCols = COLUMNS.filter((c) => c !== 'project_id' && row[c] !== undefined);
  if (setCols.length === 0) return res.status(400).json({ status: 'error', msg: 'No fields to update' });

  const setClause = setCols.map((c, i) => `${c} = $${i + 2}`).join(', ');
  const values = [projectId, ...setCols.map((c) => row[c])];
  const result = await pool.query(
    `UPDATE projects SET ${setClause}, updated_at = now() WHERE project_id = $1`,
    values
  );
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'Project not found: ' + projectId });
  res.json({ status: 'ok' });
}));

// DELETE /api/projects/:projectId
router.delete('/:projectId', asyncHandler(async (req, res) => {
  const result = await pool.query('DELETE FROM projects WHERE project_id = $1', [req.params.projectId]);
  if (result.rowCount === 0) return res.status(404).json({ status: 'error', msg: 'Project not found: ' + req.params.projectId });
  res.json({ status: 'ok' });
}));

module.exports = router;
