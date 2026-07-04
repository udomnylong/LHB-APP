// CLI wrapper for manual runs — see ../syncCheckEvents.js for the actual logic
// (also used by the scheduled POST /api/admin/sync-checkevents endpoint).
require('dotenv').config();
const { runSync } = require('../syncCheckEvents');
const { pool } = require('../db');

runSync()
  .then((result) => {
    console.log(JSON.stringify(result, null, 2));
    return pool.end();
  })
  .catch((e) => { console.error('sync failed:', e); process.exit(1); });
