// One-off: clears staff.salary for every staff row (ព័ត៌មានបុគ្គលិក > ប្រាក់ខែ).
// Dry-run by default — shows which staff currently have a salary value, without
// changing anything. Pass --yes to actually apply the update.
//
// Usage:
//   node scripts/clear-staff-salaries.js          # dry run (safe, read-only)
//   node scripts/clear-staff-salaries.js --yes    # clears salary for ALL staff
//
// Strongly recommended before running with --yes: take a backup snapshot first
// (POST /api/admin/backup-to-sheets — see backupToSheets.js) so today's salary
// values are recoverable in the Google Sheet if this needs to be undone.

require('dotenv').config();
const { Pool } = require('pg');

const APPLY = process.argv.includes('--yes');

async function main() {
  const pool = new Pool({ connectionString: process.env.DATABASE_URL });
  try {
    const { rows } = await pool.query(
      `SELECT staff_code, name, salary FROM staff WHERE salary IS NOT NULL ORDER BY staff_code`
    );
    console.log(`${rows.length} staff currently have a salary value set.`);
    for (const r of rows) {
      console.log(`  ${r.staff_code}\t${r.name ?? ''}\t${r.salary}`);
    }

    if (!APPLY) {
      console.log('\nDry run only — no changes made. Re-run with --yes to clear these values.');
      return;
    }

    const result = await pool.query(
      `UPDATE staff SET salary = NULL, updated_at = now() WHERE salary IS NOT NULL`
    );
    console.log(`\nCleared salary for ${result.rowCount} staff.`);
  } finally {
    await pool.end();
  }
}

main().catch((err) => {
  console.error('Failed:', err);
  process.exit(1);
});
