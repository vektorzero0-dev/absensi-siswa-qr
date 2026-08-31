const express = require('express');
const { Pool } = require('pg');
const path = require('path');
const fs = require('fs');
const pino = require('pino');
const makeWASocket = require('@whiskeysockets/baileys').default;
const { useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const QRCode = require('qrcode');

const app = express();
const PORT = process.env.PORT || 3000;

// 1. Buat folder wa_sessions secara aman
const sessionsDir = path.join(__dirname, 'wa_sessions');
if (!fs.existsSync(sessionsDir)) {
    fs.mkdirSync(sessionsDir, { recursive: true });
}

// 2. Middleware Parsing Body Data
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Setup View Engine
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));

// 3. Database Connection
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

pool.on('error', (err) => {
    console.error('❌ Database Pool Error:', err.message);
});

const waSessions = {};
const qrCodes = {};
const waStatus = {};

// 4. Inisialisasi Tabel Database dengan Urutan SQL yang Benar
async function initDB() {
    if (!process.env.DATABASE_URL) {
        console.error("⚠️ DATABASE_URL belum diset di Environment Variables Render!");
        return;
    }

    const client = await pool.connect();
    try {
        await client.query('BEGIN');

        // Step A: Buat Tabel Kelas Utama
        await client.query(`
            CREATE TABLE IF NOT EXISTS kelas (
                id SERIAL PRIMARY KEY,
                nama_kelas VARCHAR(50) NOT NULL
            );
        `);

        // Step B: Buat Tabel Users
        await client.query(`
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                nama VARCHAR(100) NOT NULL,
                username VARCHAR(50) UNIQUE NOT NULL,
                password VARCHAR(100) NOT NULL,
                role VARCHAR(20) NOT NULL,
                kelas_id INT REFERENCES kelas(id) ON DELETE SET NULL
            );
        `);

        // Step C: Buat Tabel Siswa
        await client.query(`
            CREATE TABLE IF NOT EXISTS siswa (
                id SERIAL PRIMARY KEY,
                nama VARCHAR(100) NOT NULL,
                kelas_id INT REFERENCES kelas(id) ON DELETE SET NULL,
                nomor_wa_ortu VARCHAR(20) NOT NULL
            );
        `);

        // Step D: Buat Tabel Absensi
        await client.query(`
            CREATE TABLE IF NOT EXISTS absensi (
                id SERIAL PRIMARY KEY,
                siswa_id INT REFERENCES siswa(id) ON DELETE CASCADE,
                waktu TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status VARCHAR(20) DEFAULT 'HADIR',
                scanned_by INT
            );
        `);

        // Step E: Masukkan Data Awal (Default Data)
        await client.query(`
            INSERT INTO kelas (id, nama_kelas) VALUES (1, 'Kelas 1A') ON CONFLICT (id) DO NOTHING;
            INSERT INTO users (id, nama, username, password, role, kelas_id) 
            VALUES (1, 'Admin Sekolah', 'admin', 'admin123', 'ADMIN', NULL),
                   (2, 'Guru Kelas 1A', 'wali1', 'wali123', 'WALI_KELAS', 1) 
            ON CONFLICT (id) DO NOTHING;
        `);

        await client.query('COMMIT');
        console.log("✅ Database Postgres & Tabel Berhasil Disiapkan.");
    } catch (err) {
        await client.query('ROLLBACK');
        console.error("❌ Gagal Inisialisasi Database:", err.message);
    } finally {
        client.release();
    }
}
initDB();

// 5. WhatsApp Engine (Baileys)
async function connectToWhatsApp(userId) {
    try {
        const authFolder = path.join(sessionsDir, `user_${userId}`);
        const { state, saveCreds } = await useMultiFileAuthState(authFolder);

        const sock = makeWASocket({
            logger: pino({ level: 'silent' }),
            auth: state,
            printQRInTerminal: false
        });

        waSessions[userId] = sock;
        sock.ev.on('creds.update', saveCreds);

        sock.ev.on('connection.update', async (update) => {
            const { connection, lastDisconnect, qr } = update;

            if (qr) {
                qrCodes[userId] = await QRCode.toDataURL(qr);
                waStatus[userId] = 'MENUNGGU_SCAN';
            }

            if (connection === 'open') {
                waStatus[userId] = 'TERHUBUNG';
                delete qrCodes[userId];
                console.log(`✅ WA User ${userId} Terhubung.`);
            }

            if (connection === 'close') {
                const shouldReconnect = (lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut);
                waStatus[userId] = 'TERPUTUS';
                delete waSessions[userId];
                if (shouldReconnect) connectToWhatsApp(userId);
            }
        });
    } catch (err) {
        console.error(`❌ Error WA User ${userId}:`, err.message);
    }
}

// 6. Web Routes
app.get('/', (req, res) => res.render('login'));

app.get('/admin', async (req, res) => {
    const userId = req.query.userId || 1;
    try {
        const usersRes = await pool.query(`SELECT u.id, u.nama, u.username, u.role, k.nama_kelas FROM users u LEFT JOIN kelas k ON u.kelas_id = k.id`);
        const siswaRes = await pool.query(`SELECT s.id, s.nama, s.nomor_wa_ortu, k.nama_kelas FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id ORDER BY s.id ASC`);
        const kelasRes = await pool.query(`SELECT * FROM kelas ORDER BY id ASC`);

        res.render('admin-dashboard', {
            users: usersRes.rows,
            siswa: siswaRes.rows,
            kelas: kelasRes.rows,
            userId: userId
        });
    } catch (err) {
        console.error("Error /admin:", err.message);
        res.status(500).send("Database Error di Halaman Admin: " + err.message);
    }
});

app.get('/wali', async (req, res) => {
    const userId = req.query.userId || 2;
    try {
        const userRes = await pool.query(`
            SELECT u.id, u.nama, u.role, u.kelas_id, COALESCE(k.nama_kelas, 'Kelas 1A') AS nama_kelas 
            FROM users u LEFT JOIN kelas k ON u.kelas_id = k.id 
            WHERE u.id = $1
        `, [userId]);

        const user = userRes.rows[0] || { id: userId, nama: 'Guru Kelas', role: 'WALI_KELAS', nama_kelas: 'Kelas 1A' };
        const targetKelasId = user.kelas_id || 1;

        const siswaKelasRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, 'Kelas 1A') AS nama_kelas 
            FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id 
            WHERE s.kelas_id = $1 ORDER BY s.nama ASC
        `, [targetKelasId]);

        const absensiHariIniRes = await pool.query(`
            SELECT a.waktu, s.nama 
            FROM absensi a 
            JOIN siswa s ON a.siswa_id = s.id 
            WHERE s.kelas_id = $1 AND DATE(a.waktu) = CURRENT_DATE 
            ORDER BY a.waktu DESC
        `, [targetKelasId]);

        res.render('walikelas-dashboard', {
            user: user,
            siswaList: siswaKelasRes.rows,
            absensiHariIni: absensiHariIniRes.rows,
            userId: userId,
            statusWA: waStatus[userId] || 'BELUM_TERHUBUNG',
            qrCodeWA: qrCodes[userId] || null
        });
    } catch (err) {
        console.error("Error /wali:", err.message);
        res.status(500).send("Database Error di Halaman Wali Kelas: " + err.message);
    }
});

app.get('/scan', (req, res) => {
    const userId = req.query.userId || 1;
    res.render('scanner', { userId: userId });
});

// 7. API Scanner & WA
app.get('/api/start-wa', (req, res) => {
    const userId = req.query.userId || 1;
    connectToWhatsApp(userId);
    res.json({ success: true, message: 'Menyiapkan koneksi WA...' });
});

app.get('/api/status-wa', (req, res) => {
    const userId = req.query.userId || 1;
    res.json({
        status: waStatus[userId] || 'BELUM_TERHUBUNG',
        qr: qrCodes[userId] || null
    });
});

app.post('/api/scan', async (req, res) => {
    const { siswa_id, scanned_by } = req.body;

    if (!siswa_id) {
        return res.status(400).json({ success: false, message: "ID Siswa kosong." });
    }

    try {
        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, 'Tanpa Kelas') AS nama_kelas 
            FROM siswa s 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            WHERE s.id = $1
        `, [parseInt(siswa_id)]);

        if (siswaRes.rows.length === 0) {
            return res.status(404).json({ 
                success: false, 
                message: `Siswa ID #${siswa_id} Tidak Ditemukan di Database!` 
            });
        }

        const siswa = siswaRes.rows[0];

        await pool.query(`
            INSERT INTO absensi (siswa_id, status, scanned_by) 
            VALUES ($1, 'HADIR', $2)
        `, [siswa.id, scanned_by || 1]);

        const waClient = waSessions[scanned_by || 1];
        if (waClient && siswa.nomor_wa_ortu) {
            let formattedPhone = siswa.nomor_wa_ortu.replace(/[^0-9]/g, '');
            if (formattedPhone.startsWith('0')) formattedPhone = '62' + formattedPhone.slice(1);
            if (!formattedPhone.endsWith('@s.whatsapp.net')) formattedPhone += '@s.whatsapp.net';

            const jamMasuk = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }) + ' WIB';
            const pesanWA = `*UPTD SD NEGERI 1 KARYA MULYA SARI*\n` +
                            `--------------------------------------------------\n` +
                            `*NOTIFIKASI PRESENSI KEHADIRAN SISWA*\n\n` +
                            `Yth. Bapak/Ibu Orang Tua/Wali dari:\n` +
                            `Nama: *${siswa.nama}*\n` +
                            `Kelas: *${siswa.nama_kelas}*\n` +
                            `Waktu Presensi: *${jamMasuk}*\n` +
                            `Status: *HADIR ✅*\n\n` +
                            `Ananda telah tiba di sekolah dengan selamat. Terima kasih.`;

            waClient.sendMessage(formattedPhone, { text: pesanWA }).catch(err => console.error("WA Error:", err));
        }

        return res.json({
            success: true,
            message: "Presensi Berhasil Dicatat!",
            siswa: { id: siswa.id, nama: siswa.nama, nama_kelas: siswa.nama_kelas }
        });

    } catch (err) {
        console.error("Error API Scan:", err);
        return res.status(500).json({ success: false, message: "Terjadi kesalahan server: " + err.message });
    }
});

app.listen(PORT, () => console.log(`🚀 Server aktif di port ${PORT}`));
