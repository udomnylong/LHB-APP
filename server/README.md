# LHB HR — server

Cloud Run API and Sheets → Cloud SQL migration tooling for LHB HR.

## Setup

```
cd server
npm install
cp .env.example .env   # fill in SHEET_ID, GOOGLE_APPLICATION_CREDENTIALS, DATABASE_URL
```

You need:
- A Cloud SQL for PostgreSQL instance (or any reachable Postgres for local testing), reachable at `DATABASE_URL`.
- A GCP service account with the Sheets API enabled, JSON key saved locally, and **the Google Sheet shared with that service account's email as Viewer**.

## Database setup

```
npm run db:schema     # creates tables and applies additive schema updates (safe to re-run)
npm run migrate:etl   # pulls all 12 sheets into Postgres (safe to re-run — full refresh per table)
```

The ETL is an initial/full refresh from Sheets; re-running it replaces Cloud SQL data (including account passwords) with the current sheet contents. Use it only when an intentional refresh is needed. It prints a per-table row count summary and a punch list of orphaned rows (e.g. a CheckIn row whose staff ID doesn't exist in StaffInfo).

Run `npm run db:schema` against the target database after deploying schema updates. The API stores HR and staff-portal passwords using scrypt; existing HR passwords are upgraded from the legacy plaintext column on successful login. The ETL imports HR passwords as hashes.
For an existing database, run `npm run db:hash-passwords` once after applying the schema to hash all remaining legacy HR passwords and clear their plaintext copies.

Republish `Code.gs` in Apps Script before publishing `hr-system.html`; HR login now verifies the User sheet through Apps Script and does not require Cloud Run. Publish the updated HTML/JavaScript files after the Apps Script deployment is ready. HR data operations that use Cloud SQL still require the Cloud Run API.

## Run the API

```
npm start
```

The API listens on port 8080 by default. `GET /health` is the readiness endpoint.
