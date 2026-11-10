require('dotenv').config();
const { pool } = require('../db');
const { hashPassword } = require('../passwords');

async function main() {
  const { rows } = await pool.query(
    'SELECT id, password_plain, password_hash FROM users WHERE password_plain IS NOT NULL ORDER BY id'
  );
  let migrated = 0;

  for (const row of rows) {
    const passwordHash = row.password_hash || await hashPassword(row.password_plain);
    const result = await pool.query(
      `UPDATE users
       SET password_hash = $1, password_plain = NULL, updated_at = now()
       WHERE id = $2 AND password_plain = $3`,
      [passwordHash, row.id, row.password_plain]
    );
    migrated += result.rowCount;
  }

  console.log(`Migrated ${migrated} legacy user passwords to scrypt hashes.`);
}

main()
  .catch((err) => {
    console.error('Failed to migrate legacy passwords:', err);
    process.exitCode = 1;
  })
  .finally(() => pool.end());
