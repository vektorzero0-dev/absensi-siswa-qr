const express = require('express');
const { Pool } = require('pg');
const path = require('path');
const fs = require('fs');
const pino = require('pino');
const makeWASocket = require('@whiskeysockets/baileys').default;
const { useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const QRCode = require('qrcode');

// 1. INSIALISASI EXPRESS (Wajib sebelum app.get / app.post)
const app = express();
const PORT = process.env.PORT || 3000;

// 2. Buat folder wa_sessions jika belum ada
const sessionsDir = path.join(__dirname, 'wa_sessions');
if (!fs.existsSync(sessionsDir)) {
    fs.mkdirSync(sessionsDir, { recursive: true });
}

// 3. Middleware Body Parser & Asset Static
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));

// 4. Koneksi PostgreSQL (Neon.tech)
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

// 5. Otomatisasi Tabel Database
async function initDB() {
    if (!process.env.DATABASE_URL) return;

    const client = await pool.connect();
    try {
        await client.query('BEGIN');

        await client.query(`
            CREATE TABLE IF NOT EXISTS kelas (
                id SERIAL PRIMARY KEY,
                nama_kelas VARCHAR(50) NOT NULL
            );
        `);

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

        await client.query(`
            CREATE TABLE IF NOT EXISTS siswa (
                id SERIAL PRIMARY KEY,
                nama VARCHAR(100) NOT NULL,
                kelas_id INT REFERENCES kelas(id) ON DELETE SET NULL,
                nomor_wa_ortu VARCHAR(20) NOT NULL
            );
        `);

        await client.query(`
            CREATE TABLE IF NOT EXISTS absensi (
                id SERIAL PRIMARY KEY,
                siswa_id INT REFERENCES siswa(id) ON DELETE CASCADE,
                waktu TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status VARCHAR(20) DEFAULT 'HADIR',
                scanned_by INT
            );
        `);

        await client.query(`
            INSERT INTO kelas (id, nama_kelas) VALUES (1, 'Kelas 1A') ON CONFLICT (id) DO NOTHING;
            INSERT INTO users (id, nama, username, password, role, kelas_id) 
            VALUES (1, 'Admin Sekolah', 'admin', 'admin123', 'ADMIN', NULL),
                   (2, 'Guru Kelas 1A', 'wali1', 'wali123', 'WALI_KELAS', 1) 
            ON CONFLICT (id) DO NOTHING;
            INSERT INTO siswa (id, nama, kelas_id, nomor_wa_ortu)
            VALUES (1, 'Budi Santoso', 1, '081234567890')
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

// 6. Integrasi WhatsApp Engine
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

// 7. ROUTE HALAMAN WEB & AKSI LOGIN

// Tampilan Login
app.get('/', (req, res) => {
    res.render('login', { error: null });
});

// Proses Form Login
app.post('/login', async (req, res) => {
    const { username, password } = req.body;

    try {
        const result = await pool.query(
            'SELECT * FROM users WHERE username = $1 AND password = $2',
            [username, password]
        );

        if (result.rows.length === 0) {
            return res.render('login', { error: 'Username atau password salah!' });
        }

        const user = result.rows[0];

        if (user.role === 'ADMIN') {
            return res.redirect(`/admin?userId=${user.id}`);
        } else {
            return res.redirect(`/wali?userId=${user.id}`);
        }

    } catch (err) {
        console.error("Error pada saat login:", err.message);
        return res.render('login', { error: 'Terjadi kesalahan sistem saat login.' });
    }
});

// Admin Dashboard
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

// Wali Kelas Dashboard
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

// Scanner Page
app.get('/scan', (req, res) => {
    const userId = req.query.userId || 1;
    res.render('scanner', { userId: userId });
});

// 8. API Endpoints
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
                message: `Siswa ID #${siswa_id} Tidak Ditemukan!` 
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

// 9. Jalankan Server
app.listen(PORT, () => console.log(`🚀 Server aktif di port ${PORT}`));
