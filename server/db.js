const { Pool } = require('pg');
require('dotenv').config();

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL?.includes('localhost') ? false : { rejectUnauthorized: false }
});

async function setupDatabase() {
  const client = await pool.connect();
  try {
    // Templates table — stores the xlsx files as binary blobs
    await client.query(`
      CREATE TABLE IF NOT EXISTS xlsx_templates (
        id SERIAL PRIMARY KEY,
        service VARCHAR(50) NOT NULL UNIQUE,  -- 'Assisted' or 'Negotiation'
        name VARCHAR(255) NOT NULL,
        file_data BYTEA NOT NULL,
        file_name VARCHAR(255) NOT NULL,
        uploaded_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW()
      )
    `);

    // Processing log — records every webhook call
    await client.query(`
      CREATE TABLE IF NOT EXISTS processing_log (
        id SERIAL PRIMARY KEY,
        hs_object_id BIGINT NOT NULL,
        service VARCHAR(50) NOT NULL,
        status VARCHAR(20) NOT NULL DEFAULT 'pending',  -- pending, success, error
        error_message TEXT,
        n8n_response TEXT,
        created_at TIMESTAMP DEFAULT NOW()
      )
    `);

    // Add generated file columns to processing_log (idempotent)
    await client.query(`
      ALTER TABLE processing_log
        ADD COLUMN IF NOT EXISTS generated_file BYTEA,
        ADD COLUMN IF NOT EXISTS generated_file_name VARCHAR(255)
    `);

    console.log('✅ Database tables created');
  } finally {
    client.release();
  }
}

module.exports = { pool, setupDatabase };
