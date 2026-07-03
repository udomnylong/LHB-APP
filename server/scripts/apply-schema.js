// Applies server/schema.sql to the DB pointed at by DATABASE_URL.
// Safe to re-run — schema.sql uses CREATE TABLE/INDEX IF NOT EXISTS throughout.

require('dotenv').config();
const fs = require('fs');
const path = require('path');
const { Pool } = require('pg');

async function main() {
  const sql = fs.readFileSync(path.join(__dirname, '..', 'schema.sql'), 'utf8');
  const pool = new Pool({ connectionString: process.env.DATABASE_URL });
  try {
    await pool.query(sql);
    console.log('Schema applied successfully.');
  } finally {
    await pool.end();
  }
}

main().catch((err) => {
  console.error('Failed to apply schema:', err);
  process.exit(1);
});
