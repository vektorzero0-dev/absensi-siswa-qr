const express = require('express');
const { Pool } = require('pg');
const path = require('path');
const QRCode = require('qrcode');
const session = require('express-session');
const makeWASocket = require('@whiskeysockets/baileys').default;
const { useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// ==========================================
// 1. MIDDLEWARE UTAMA
// ==========================================
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

app.use(session({
    secret: process.env.SESSION_SECRET || 'presensi-sdn1-karyamulyasari-secret',
    resave: false,
    saveUninitialized: true,
    cookie: { maxAge: 24 * 60 * 60 * 1000 } // 1 hari
}));

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// ==========================================
// 2. KONEKSI DATABASE POSTGRESQL
// ==========================================
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.DATABASE_URL && !process.env.DATABASE_URL.includes('localhost') 
        ? { rejectUnauthorized: false } 
        : false
});

// Otomatis buat Tabel jika belum ada di Database
async function initDatabase() {
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
                nomor_wa_ortu VARCHAR(20)
            );

            CREATE TABLE IF NOT EXISTS absensi (
                id SERIAL PRIMARY KEY,
                siswa_id INT REFERENCES siswa(id) ON DELETE CASCADE,
                status VARCHAR(20) DEFAULT 'HADIR',
                scanned_by VARCHAR(50),
                waktu TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        `);
        console.log("✅ Database Postgres siap dan tabel telah diverifikasi.");
    } catch (err) {
        console.error("❌ Gagal inisialisasi tabel database:", err.message);
    }
}
initDatabase();

// ==========================================
// 3. WHATSAPP GATEWAY (BAILEYS MULTI-SESSION)
// ==========================================
const waSessions = {};
const waQrCodes = {};

async function connectToWhatsApp(sessionId = 'default') {
    const authFolder = path.join(__dirname, 'auth_info_baileys', sessionId);
    if (!fs.existsSync(authFolder)) {
        fs.mkdirSync(authFolder, { recursive: true });
    }

    const { state, saveCreds } = await useMultiFileAuthState(authFolder);

    const sock = makeWASocket({
        auth: state,
        printQRInTerminal: true,
        browser: ["Presensi Sekolah", "Chrome", "1.0.0"]
    });

    sock.ev.on('creds.update', saveCreds);

    sock.ev.on('connection.update', async (update) => {
        const { connection, lastDisconnect, qr } = update;

        if (qr) {
            waQrCodes[sessionId] = await QRCode.toDataURL(qr);
        }

        if (connection === 'close') {
            const shouldReconnect = (lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut);
            console.log(`[WA:${sessionId}] Koneksi terputus. Reconnect: ${shouldReconnect}`);
            delete waSessions[sessionId];
            delete waQrCodes[sessionId];
            if (shouldReconnect) {
                setTimeout(() => connectToWhatsApp(sessionId), 5000);
            }
        } else if (connection === 'open') {
            console.log(`✅ [WA:${sessionId}] WhatsApp Terhubung!`);
            waSessions[sessionId] = sock;
            delete waQrCodes[sessionId];
        }
    });
}

// Inisialisasi WhatsApp Session bawaan saat server mulai
connectToWhatsApp('default');

// ==========================================
// 4. ROUTE VIEW (HALAMAN WEB)
// ==========================================
app.get('/', async (req, res) => {
    try {
        const totalSiswa = await pool.query('SELECT COUNT(*) FROM siswa');
        const totalHadir = await pool.query(`SELECT COUNT(DISTINCT siswa_id) FROM absensi WHERE DATE(waktu) = CURRENT_DATE`);
        
        res.render('dashboard', {
            totalSiswa: totalSiswa.rows[0].count,
            totalHadir: totalHadir.rows[0].count,
            user: req.session.user || { name: 'Admin' }
        });
    } catch (err) {
        res.render('dashboard', { totalSiswa: 0, totalHadir: 0, user: { name: 'Admin' } });
    }
});

app.get('/scanner', (req, res) => {
    res.render('scanner', {
        userId: req.session.user ? req.session.user.id : 'default'
    });
});

app.get('/wa-setting', (req, res) => {
    res.render('wa-setting');
});

// ==========================================
// 5. API ENDPOINT WHATSAPP
// ==========================================
app.get('/api/wa/qr', (req, res) => {
    const sessionId = req.query.sessionId || 'default';
    if (waSessions[sessionId]) {
        return res.json({ connected: true, message: "WhatsApp sudah terhubung." });
    }
    if (waQrCodes[sessionId]) {
        return res.json({ connected: false, qr: waQrCodes[sessionId] });
    }
    return res.json({ connected: false, message: "Membuat QR Code... Silakan refresh sebentar lagi." });
});

app.get('/api/wa/status', (req, res) => {
    const sessionId = req.query.sessionId || 'default';
    return res.json({ connected: !!waSessions[sessionId] });
});

// ==========================================
// 6. API ENDPOINT SCANNER PRESENSI (UTAMA)
// ==========================================
app.post('/api/scan', async (req, res) => {
    const { siswa_id, scanned_by } = req.body;
    const sessionKey = scanned_by || 'default';

    if (!siswa_id) {
        return res.status(400).json({ success: false, message: "ID Siswa tidak terdeteksi." });
    }

    try {
        // 1. Cari Siswa di Database
        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, 'Tanpa Kelas') AS nama_kelas 
            FROM siswa s 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            WHERE s.id = $1
        `, [siswa_id]);

        if (siswaRes.rows.length === 0) {
            return res.status(404).json({ 
                success: false, 
                message: `Siswa dengan ID #${siswa_id} Tidak Ditemukan!` 
            });
        }

        const siswa = siswaRes.rows[0];

        // 2. Simpan Rekam Absensi ke Database
        await pool.query(`
            INSERT INTO absensi (siswa_id, status, scanned_by) 
            VALUES ($1, 'HADIR', $2)
        `, [siswa.id, sessionKey]);

        // 3. Kirim Pesan WhatsApp ke Orang Tua
        const waClient = waSessions[sessionKey] || waSessions['default'];
        let waStatus = "Pesan WA Tidak Terkirim (WA Belum Konek / No Ortu Kosong)";

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

            waClient.sendMessage(formattedPhone, { text: pesanWA })
                .then(() => console.log(`✅ WA Terkirim ke Ortu ${siswa.nama}`))
                .catch(err => console.error("❌ Gagal Kirim WA:", err.message));
                
            waStatus = "Notifikasi WA Berhasil Dikirim!";
        }

        // 4. Balasan Sukses ke Web Scanner
        return res.json({
            success: true,
            message: "Presensi berhasil dicatat.",
            waStatus: waStatus,
            siswa: {
                id: siswa.id,
                nama: siswa.nama,
                nama_kelas: siswa.nama_kelas
            }
        });

    } catch (err) {
        console.error("❌ Error API /api/scan:", err);
        return res.status(500).json({ 
            success: false, 
            message: "Terjadi kesalahan internal server: " + err.message 
        });
    }
});

// ==========================================
// 7. API ENDPOINT DATA SISWA & KELAS
// ==========================================
app.get('/api/siswa', async (req, res) => {
    try {
        const result = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, k.nama_kelas 
            FROM siswa s 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            ORDER BY s.id ASC
        `);
        res.json(result.rows);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/siswa', async (req, res) => {
    const { nama, kelas_id, nomor_wa_ortu } = req.body;
    try {
        const result = await pool.query(`
            INSERT INTO siswa (nama, kelas_id, nomor_wa_ortu) 
            VALUES ($1, $2, $3) RETURNING *
        `, [nama, kelas_id || null, nomor_wa_ortu || null]);
        res.json({ success: true, data: result.rows[0] });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

app.get('/api/absensi/hari-ini', async (req, res) => {
    try {
        const result = await pool.query(`
            SELECT a.id, s.nama, COALESCE(k.nama_kelas, 'Tanpa Kelas') AS nama_kelas, a.waktu, a.status 
            FROM absensi a 
            JOIN siswa s ON a.siswa_id = s.id 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            WHERE DATE(a.waktu) = CURRENT_DATE 
            ORDER BY a.waktu DESC
        `);
        res.json(result.rows);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Fallback Route untuk 404
app.use((req, res) => {
    res.status(404).send("Halaman tidak ditemukan.");
});

// ==========================================
// 8. MENJALANKAN SERVER
// ==========================================
app.listen(PORT, () => {
    console.log(`🚀 Server berjalan di http://localhost:${PORT}`);
});
