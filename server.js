const express = require('express');
const { Pool } = require('pg');
const path = require('path');
const pino = require('pino');
const makeWASocket = require('@whiskeysockets/baileys').default;
const { useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const QRCode = require('qrcode');

const app = express();
const PORT = process.env.PORT || 3000;

// ==========================================
// 1. MIDDLEWARE PARSING (PENTING AGAR API SCAN JALAN)
// ==========================================
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Setup View Engine & Folder Statis
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));

// ==========================================
// 2. KONEKSI DATABASE POSTGRESQL (NEON.TECH)
// ==========================================
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: { rejectUnauthorized: false }
});

// Penyimpanan Sesi WhatsApp Multi-User
const waSessions = {};
const qrCodes = {};
const waStatus = {};

// Inisialisasi Otomatis Tabel Database
async function initDB() {
    try {
        await pool.query(`
            CREATE TABLE IF NOT EXISTS kelas (
                id SERIAL PRIMARY KEY,
                nama_kelas VARCHAR(50) NOT NULL
            );

            CREATE TABLE IF NOT EXISTS siswa (
                id SERIAL PRIMARY KEY,
                nama VARCHAR(100) NOT NULL,
                kelas_id INT REFERENCES kelas(id) ON DELETE SET NULL,
                nomor_wa_ortu VARCHAR(20) NOT NULL
            );

            CREATE TABLE IF NOT EXISTS absensi (
                id SERIAL PRIMARY KEY,
                siswa_id INT REFERENCES siswa(id) ON DELETE CASCADE,
                waktu TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status VARCHAR(20) DEFAULT 'HADIR',
                scanned_by INT
            );
        `);
        console.log("✅ Database Postgres siap digunakan.");
    } catch (err) {
        console.error("❌ Gagal inisialisasi database:", err.message);
    }
}
initDB();

// ==========================================
// 3. INTEGRASI WHATSAPP ENGINE (BAILEYS)
// ==========================================
async function connectToWhatsApp(userId) {
    const authFolder = path.join(__dirname, 'wa_sessions', `user_${userId}`);
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
            console.log(`✅ WA User ID ${userId} berhasil terhubung.`);
        }

        if (connection === 'close') {
            const shouldReconnect = (lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut);
            waStatus[userId] = 'TERPUTUS';
            delete waSessions[userId];
            if (shouldReconnect) {
                connectToWhatsApp(userId);
            }
        }
    });
}

// ==========================================
// 4. ROUTE HALAMAN WEB
// ==========================================

// Halaman Login
app.get('/', (req, res) => {
    res.render('login');
});

// Halaman Dashboard Admin
app.get('/admin', async (req, res) => {
    const userId = req.query.userId;
    try {
        const usersRes = await pool.query(`SELECT u.id, u.nama, u.username, u.role, k.nama_kelas FROM users u LEFT JOIN kelas k ON u.kelas_id = k.id`);
        const siswaRes = await pool.query(`SELECT s.id, s.nama, s.nomor_wa_ortu, k.nama_kelas FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id ORDER BY s.id ASC`);
        const kelasRes = await pool.query(`SELECT * FROM kelas ORDER BY id ASC`);

        res.render('admin-dashboard', {
            users: usersRes.rows,
            siswa: siswaRes.rows,
            kelas: kelasRes.rows,
            userId: userId || 1
        });
    } catch (err) {
        console.error("Error /admin:", err);
        res.status(500).send("Terjadi kesalahan memuat data admin.");
    }
});

// Halaman Dashboard Wali Kelas
app.get('/wali', async (req, res) => {
    const userId = req.query.userId;
    if (!userId) return res.redirect('/');

    try {
        const userRes = await pool.query(`
            SELECT u.id, u.nama, u.role, u.kelas_id, COALESCE(k.nama_kelas, 'Kelas 1A') AS nama_kelas 
            FROM users u LEFT JOIN kelas k ON u.kelas_id = k.id 
            WHERE u.id = $1
        `, [userId]);

        if (userRes.rows.length === 0) return res.redirect('/');
        const user = userRes.rows[0];
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
        console.error("Error /wali:", err);
        res.status(500).send("Terjadi kesalahan memuat dashboard guru.");
    }
});

// Halaman Scanner QR Presensi
app.get('/scan', (req, res) => {
    const userId = req.query.userId || 1;
    res.render('scanner', { userId: userId });
});

// ==========================================
// 5. API ENDPOINTS (WA & SCANNER)
// ==========================================

// API Mulai Sesi WA
app.get('/api/start-wa', (req, res) => {
    const userId = req.query.userId;
    if (!userId) return res.json({ success: false, message: 'User ID diperlukan' });

    connectToWhatsApp(userId);
    res.json({ success: true, message: 'Socket WA sedang dinyalakan.' });
});

// API Cek Status WA
app.get('/api/status-wa', (req, res) => {
    const userId = req.query.userId;
    res.json({
        status: waStatus[userId] || 'BELUM_TERHUBUNG',
        qr: qrCodes[userId] || null
    });
});

// API PROSES SCAN PRESENSI SISWA (ROUTER UTAMA)
app.post('/api/scan', async (req, res) => {
    const { siswa_id, scanned_by } = req.body;

    if (!siswa_id) {
        return res.status(400).json({ success: false, message: "ID Siswa kosong." });
    }

    try {
        // 1. Cari siswa di database
        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, 'Tanpa Kelas') AS nama_kelas 
            FROM siswa s 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            WHERE s.id = $1
        `, [parseInt(siswa_id)]);

        if (siswaRes.rows.length === 0) {
            return res.status(404).json({ 
                success: false, 
                message: `ID Siswa #${siswa_id} Tidak Ditemukan di Database!` 
            });
        }

        const siswa = siswaRes.rows[0];

        // 2. Simpan catatan absensi
        await pool.query(`
            INSERT INTO absensi (siswa_id, status, scanned_by) 
            VALUES ($1, 'HADIR', $2)
        `, [siswa.id, scanned_by || null]);

        // 3. Kirim WhatsApp ke Orang Tua jika WA Guru Aktif
        const waClient = waSessions[scanned_by];
        if (waClient && siswa.nomor_wa_ortu) {
            let formattedPhone = siswa.nomor_wa_ortu.replace(/[^0-9]/g, '');
            if (formattedPhone.startsWith('0')) {
                formattedPhone = '62' + formattedPhone.slice(1);
            }
            if (!formattedPhone.endsWith('@s.whatsapp.net')) {
                formattedPhone = formattedPhone + '@s.whatsapp.net';
            }

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

            waClient.sendMessage(formattedPhone, { text: pesanWA }).catch(err => {
                console.error("Gagal kirim pesan WA:", err);
            });
        }

        // 4. Balasan ke Tampilan Scanner Web
        return res.json({
            success: true,
            message: "Presensi berhasil dicatat.",
            siswa: {
                id: siswa.id,
                nama: siswa.nama,
                nama_kelas: siswa.nama_kelas
            }
        });

    } catch (err) {
        console.error("Error API Scan:", err);
        return res.status(500).json({ 
            success: false, 
            message: "Terjadi kesalahan server: " + err.message 
        });
    }
});

// ==========================================
// 6. MENJALANKAN SERVER
// ==========================================
app.listen(PORT, () => {
    console.log(`🚀 Server absensi berjalan pada port ${PORT}`);
});
