const express = require('express');
const asyncHandler = require('../asyncHandler');
const { runSync } = require('../syncCheckEvents');
const { runBackup } = require('../backupToSheets');
const { requireAuth } = require('../middleware/auth');

const router = express.Router();

// Shared-secret auth (not a user session) — this endpoint is only ever called by
// Cloud Scheduler, never by the frontend.
function requireCronSecret(req, res, next) {
  const expected = process.env.SYNC_CRON_SECRET;
  if (!expected) return res.status(500).json({ status: 'error', msg: 'SYNC_CRON_SECRET not configured' });
  const provided = req.headers['x-cron-secret'];
  if (provided !== expected) return res.status(401).json({ status: 'error', msg: 'Unauthorized' });
  next();
}

// POST /api/admin/sync-checkevents — pulls any CheckIn/CheckOut rows written directly
// to the Sheet (e.g. via the admin manual-attendance-edit forms) into Cloud SQL.
router.post('/sync-checkevents', requireCronSecret, asyncHandler(async (req, res) => {
  const result = await runSync();
  console.log('[sync-checkevents]', JSON.stringify(result));
  res.json({ status: 'ok', ...result });
}));

// POST /api/admin/backup-to-sheets — HR-admin-triggered (real login token required).
// Pushes current Cloud SQL data for the 8 migrated resources back into the Sheet as
// a backup snapshot. See backupToSheets.js for why it's scoped to just those 8.
router.post('/backup-to-sheets', requireAuth, asyncHandler(async (req, res) => {
  const result = await runBackup();
  console.log('[backup-to-sheets]', JSON.stringify(result));
  res.json({ status: 'ok', ...result });
}));

module.exports = router;
