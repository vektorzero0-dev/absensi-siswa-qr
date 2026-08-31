const express = require('express');
const path = require('path');
const fs = require('fs');
const pino = require('pino');
const makeWASocket = require('@whiskeysockets/baileys').default;
const { useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const QRCode = require('qrcode');
const pool = require('./db');

const app = express();
const PORT = process.env.PORT || 3000;

// Sesi WhatsApp folder setup
const sessionsDir = path.join(__dirname, 'wa_sessions');
if (!fs.existsSync(sessionsDir)) fs.mkdirSync(sessionsDir, { recursive: true });

// Middleware Setup
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));

const waSessions = {};
const qrCodes = {};
const waStatus = {};

// WhatsApp Connection Engine
async function connectToWhatsApp(userId) {
    try {
        const authFolder = path.join(sessionsDir, `user_${userId}`);
        const { state, saveCreds } = await useMultiFileAuthState(authFolder);
        const sock = makeWASocket({ logger: pino({ level: 'silent' }), auth: state, printQRInTerminal: false });

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
            }
            if (connection === 'close') {
                const shouldReconnect = (lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut);
                waStatus[userId] = 'TERPUTUS';
                delete waSessions[userId];
                if (shouldReconnect) connectToWhatsApp(userId);
            }
        });
    } catch (err) {
        console.error(`❌ WA Error User ${userId}:`, err.message);
    }
}

// ---------------- VIEWS ROUTES ---------------- //

// Halaman Login
app.get('/', (req, res) => res.render('login', { error: null }));

// Proses Login
app.post('/login', async (req, res) => {
    const { username, password } = req.body;
    try {
        const result = await pool.query(
            'SELECT * FROM users WHERE LOWER(username) = LOWER($1) AND password = $2',
            [username.trim(), password.trim()]
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
        return res.render('login', { error: 'Kesalahan sistem database.' });
    }
});

// Dashboard Admin
app.get('/admin', async (req, res) => {
    const userId = req.query.userId || 1;
    try {
        const usersRes = await pool.query(`
            SELECT u.id, u.nama, u.username, u.role, COALESCE(k.nama_kelas, 'Guru Mapel (PAI / PJOK)') AS nama_kelas 
            FROM users u LEFT JOIN kelas k ON u.kelas_id = k.id 
            ORDER BY u.id ASC
        `);
        const siswaRes = await pool.query(`SELECT s.id, s.nama, s.nomor_wa_ortu, k.nama_kelas FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id ORDER BY s.id ASC`);
        const kelasRes = await pool.query(`SELECT * FROM kelas ORDER BY id ASC`);

        res.render('admin-dashboard', {
            users: usersRes.rows,
            siswa: siswaRes.rows,
            kelas: kelasRes.rows,
            userId: userId
        });
    } catch (err) {
        res.status(500).send("Database Error: " + err.message);
    }
});

// Dashboard Wali Kelas / Guru Mapel
app.get('/wali', async (req, res) => {
    const userId = req.query.userId || 2;
    try {
        const userRes = await pool.query(`
            SELECT u.id, u.nama, u.role, u.kelas_id, COALESCE(k.nama_kelas, 'Semua Kelas') AS nama_kelas 
            FROM users u LEFT JOIN kelas k ON u.kelas_id = k.id WHERE u.id = $1
        `, [userId]);

        const user = userRes.rows[0] || { id: userId, nama: 'Guru', role: 'WALI_KELAS', nama_kelas: 'Kelas 1A' };
        const targetKelasId = user.kelas_id;

        let siswaQuery = `SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, '-') AS nama_kelas FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id`;
        let absensiQuery = `SELECT a.waktu, s.nama FROM absensi a JOIN siswa s ON a.siswa_id = s.id WHERE DATE(a.waktu) = CURRENT_DATE`;

        if (targetKelasId) {
            siswaQuery += ` WHERE s.kelas_id = ${targetKelasId}`;
            absensiQuery += ` AND s.kelas_id = ${targetKelasId}`;
        }
        siswaQuery += ` ORDER BY s.nama ASC`;
        absensiQuery += ` ORDER BY a.waktu DESC`;

        const siswaKelasRes = await pool.query(siswaQuery);
        const absensiRes = await pool.query(absensiQuery);

        res.render('walikelas-dashboard', {
            user: user,
            siswaList: siswaKelasRes.rows,
            absensiHariIni: absensiRes.rows,
            userId: userId,
            statusWA: waStatus[userId] || 'BELUM_TERHUBUNG',
            qrCodeWA: qrCodes[userId] || null
        });
    } catch (err) {
        res.status(500).send("Database Error: " + err.message);
    }
});

// Route Scanner QR
app.get(['/scan', '/scanner'], (req, res) => {
    const userId = req.query.userId || 1;
    res.render('scan', { userId: userId });
});

// ---------------- API ENDPOINTS ---------------- //

app.get('/api/start-wa', (req, res) => {
    const userId = req.query.userId || 1;
    connectToWhatsApp(userId);
    res.json({ success: true, message: 'Menyiapkan koneksi WhatsApp...' });
});

app.get('/api/status-wa', (req, res) => {
    const userId = req.query.userId || 1;
    res.json({ status: waStatus[userId] || 'BELUM_TERHUBUNG', qr: qrCodes[userId] || null });
});

app.post('/api/scan', async (req, res) => {
    const { siswa_id, scanned_by } = req.body;
    if (!siswa_id) return res.status(400).json({ success: false, message: "ID Siswa kosong." });

    try {
        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, 'Tanpa Kelas') AS nama_kelas 
            FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id WHERE s.id = $1
        `, [parseInt(siswa_id)]);

        if (siswaRes.rows.length === 0) {
            return res.status(404).json({ success: false, message: `Siswa ID #${siswa_id} Tidak Ditemukan!` });
        }

        const siswa = siswaRes.rows[0];
        await pool.query(`INSERT INTO absensi (siswa_id, status, scanned_by) VALUES ($1, 'HADIR', $2)`, [siswa.id, scanned_by || 1]);

        const waClient = waSessions[scanned_by || 1];
        if (waClient && siswa.nomor_wa_ortu) {
            let phone = siswa.nomor_wa_ortu.replace(/[^0-9]/g, '');
            if (phone.startsWith('0')) phone = '62' + phone.slice(1);
            if (!phone.endsWith('@s.whatsapp.net')) phone += '@s.whatsapp.net';

            const jam = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }) + ' WIB';
            const message = `*UPTD SD NEGERI 1 KARYA MULYA SARI*\n` +
                            `--------------------------------------------------\n` +
                            `*NOTIFIKASI PRESENSI KEHADIRAN SISWA*\n\n` +
                            `Halo Bapak/Ibu Wali Murid dari:\n` +
                            `• Nama: *${siswa.nama}*\n` +
                            `• Kelas: *${siswa.nama_kelas}*\n` +
                            `• Jam Masuk: *${jam}*\n` +
                            `• Status: *HADIR ✅*\n\n` +
                            `Ananda telah tiba di sekolah dengan selamat. Terima kasih.`;

            waClient.sendMessage(phone, { text: message }).catch(e => console.error("WA Send Error:", e));
        }

        return res.json({
            success: true,
            message: "Presensi Berhasil Dicatat!",
            siswa: { id: siswa.id, nama: siswa.nama, nama_kelas: siswa.nama_kelas }
        });
    } catch (err) {
        return res.status(500).json({ success: false, message: "Server error: " + err.message });
    }
});

app.listen(PORT, () => console.log(`🚀 Server aktif di port ${PORT}`));
