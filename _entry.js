require('pg');
require('pg-hstore');

const app = require('./index.js');

// ─────────────────────────────────────────────────────────────────────────────
// COLD-START GUARD
// Vercel invokes the handler the moment the module loads, but startServer() in
// index.js calls initializeDatabase() *asynchronously*. The first request can
// arrive before the admin user is created → 401.
// We poll the DB (max 10 s) until the admin row exists, then cache the result.
// ─────────────────────────────────────────────────────────────────────────────

let readyPromise = null;

async function waitUntilReady() {
  if (!process.env.DATABASE_URL) return; // no DB configured, let it fail naturally

  const { Pool } = require('pg');
  const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: { rejectUnauthorized: false },
    max: 1,                      // single connection, avoids Neon limit issues
    connectionTimeoutMillis: 5000,
    idleTimeoutMillis: 3000,
  });

  const deadline = Date.now() + 10000; // wait at most 10 seconds

  while (Date.now() < deadline) {
    try {
      const { rows } = await pool.query(
        `SELECT 1 FROM "Users" WHERE email = 'admin@validiant.com' LIMIT 1`
      );
      if (rows.length > 0) break; // admin created — ready to serve
    } catch (_) {
      // table may not exist yet; keep polling
    }
    await new Promise(r => setTimeout(r, 300));
  }

  await pool.end().catch(() => {}); // release the polling connection
}

module.exports = async (req, res) => {
  // Cache the promise so only the first request of each cold start polls;
  // all subsequent requests resolve instantly against the cached promise.
  if (!readyPromise) readyPromise = waitUntilReady();
  await readyPromise;
  return app(req, res);
};
