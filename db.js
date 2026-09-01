const { Pool } = require('pg');

const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

// Paksa koneksi PostgreSQL menggunakan WIB (Asia/Jakarta)
pool.on('connect', (client) => {
    client.query("SET timezone = 'Asia/Jakarta'");
});

pool.on('error', (err) => {
    console.error('❌ Database Error:', err.message);
});

module.exports = pool;
