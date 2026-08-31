const express = require('express');
const { Pool } = require('pg');
const path = require('path');
const qrcode = require('qrcode');
const { 
    default: makeWASocket, 
    useMultiFileAuthState, 
    DisconnectReason 
} = require('@whiskeysockets/baileys');

const app = express();
const port = process.env.PORT || 10000;

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Database Neon PostgreSQL
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: { rejectUnauthorized: false }
});

// Penyimpanan Sesi WhatsApp per User
const waSockets = {};
const qrCodes = {};
const waStatus = {};

// Inisialisasi WA via Baileys (Sangat Ringan & Bebas Chromium)
async function initWA(userId) {
    if (waSockets[userId]) return;

    waStatus[userId] = 'MENYIAPKAN';
    console.log(`[Baileys] Menyiapkan sesi WA untuk User ID: ${userId}...`);

    try {
        const { state, saveCreds } = await useMultiFileAuthState(`baileys_session_user_${userId}`);

        const sock = makeWASocket({
            auth: state,
            printQRInTerminal: false,
            browser: ["Absensi Sekolah", "Chrome", "1.0.0"]
        });

        sock.ev.on('creds.update', saveCreds);

        sock.ev.on('connection.update', (update) => {
            const { connection, lastDisconnect, qr } = update;

            if (qr) {
                qrcode.toDataURL(qr, (err, url) => {
                    if (!err) {
                        qrCodes[userId] = url;
                        waStatus[userId] = 'BELUM_TERHUBUNG';
                        console.log(`[Baileys] QR Code Siap untuk User ID: ${userId}`);
                    }
                });
            }

            if (connection === 'open') {
                console.log(`[Baileys] User ID ${userId} BERHASIL TERHUBUNG!`);
                waStatus[userId] = 'TERHUBUNG';
                qrCodes[userId] = null;
            }

            if (connection === 'close') {
                const shouldReconnect = (lastDisconnect?.error)?.output?.statusCode !== DisconnectReason.loggedOut;
                console.log(`[Baileys] Koneksi terputus untuk User ID ${userId}. Reconnect: ${shouldReconnect}`);
                
                delete waSockets[userId];
                delete qrCodes[userId];
                waStatus[userId] = 'BELUM_TERHUBUNG';

                if (shouldReconnect) {
                    setTimeout(() => initWA(userId), 3000);
                }
            }
        });

        waSockets[userId] = sock;

    } catch (err) {
        console.error(`[Baileys Error] Gagal inisialisasi User ID ${userId}:`, err.message);
        waStatus[userId] = 'BELUM_TERHUBUNG';
    }
}

// ------------------- ROUTES -------------------

app.get('/', (req, res) => {
    res.render('login', { error: null });
});

app.post('/login', async (req, res) => {
    const { username, password } = req.body;
    try {
        const query = `
            SELECT u.id, u.nama, u.username, u.password, u.role, u.kelas_id,
                   COALESCE(k.nama_kelas, 'Tidak Ada Kelas') AS nama_kelas
            FROM users u
            LEFT JOIN kelas k ON u.kelas_id = k.id
            WHERE LOWER(u.username) = LOWER($1) AND u.password = $2
        `;
        const result = await pool.query(query, [username.trim(), password.trim()]);

        if (result.rows.length === 0) {
            return res.render('login', { error: 'Username atau Password salah!' });
        }

        const user = result.rows[0];
        if (user.role === 'ADMIN_OPS') {
            return res.redirect(`/admin?userId=${user.id}`);
        }

        return res.redirect(`/wali?userId=${user.id}`);
    } catch (err) {
        res.render('login', { error: 'Terjadi kesalahan sistem database.' });
    }
});

app.get('/admin', async (req, res) => {
    const userId = req.query.userId;
    try {
        const siswaRes = await pool.query(`SELECT s.id, s.nama, s.nomor_wa_ortu, k.nama_kelas FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id ORDER BY k.id ASC, s.nama ASC`);
        const usersRes = await pool.query(`SELECT u.id, u.nama, u.username, u.role, COALESCE(k.nama_kelas, 'Guru Mapel / Admin') AS nama_kelas FROM users u LEFT JOIN kelas k ON u.kelas_id = k.id ORDER BY u.id ASC`);
        const absensiRes = await pool.query(`SELECT a.id, a.siswa_id, a.waktu, a.status, s.nama, k.nama_kelas FROM absensi a JOIN siswa s ON a.siswa_id = s.id LEFT JOIN kelas k ON s.kelas_id = k.id WHERE DATE(a.waktu) = CURRENT_DATE ORDER BY a.waktu DESC`);

        res.render('admin-dashboard', {
            users: usersRes.rows,
            siswa: siswaRes.rows,
            absensi: absensiRes.rows,
            userId: userId
        });
    } catch (err) {
        res.status(500).send("Gagal memuat Dashboard Admin.");
    }
});

app.get('/wali', async (req, res) => {
    const userId = req.query.userId;
    if (!userId) return res.redirect('/');

    try {
        const query = `
            SELECT u.id, u.nama, u.role, 
                   COALESCE(k.nama_kelas, 'Guru Mapel') AS nama_kelas
            FROM users u
            LEFT JOIN kelas k ON u.kelas_id = k.id
            WHERE u.id = $1
        `;
        const result = await pool.query(query, [userId]);
        if (result.rows.length === 0) return res.redirect('/');

        res.render('walikelas-dashboard', {
            user: result.rows[0],
            userId: userId,
            statusWA: waStatus[userId] || 'BELUM_TERHUBUNG',
            qrCodeWA: qrCodes[userId] || null
        });
    } catch (err) {
        res.status(500).send("Gagal memuat Dashboard Wali Kelas.");
    }
});

// Endpoint pemicu pembuatan QR
app.get('/api/start-wa', (req, res) => {
    const userId = req.query.userId;
    if (userId) {
        initWA(userId);
    }
    res.json({ status: 'PROSES' });
});

// Endpoint pengecekan status & data QR
app.get('/api/status-wa', (req, res) => {
    const userId = req.query.userId;
    res.json({
        status: waStatus[userId] || 'BELUM_TERHUBUNG',
        qr: qrCodes[userId] || null
    });
});

app.get('/scan', (req, res) => {
    const userId = req.query.userId;
    if (!userId) return res.redirect('/');
    res.render('scan', { userId: userId });
});

// API Pindaian Absensi & Pengiriman WA instan
app.post('/api/absen', async (req, res) => {
    const { siswaId, waliKelasId } = req.body;

    try {
        const siswaRes = await pool.query(
            `SELECT s.id, s.nama, s.nomor_wa_ortu, k.nama_kelas 
             FROM siswa s 
             LEFT JOIN kelas k ON s.kelas_id = k.id 
             WHERE s.id = $1`, 
            [siswaId]
        );

        if (siswaRes.rows.length === 0) {
            return res.json({ success: false, message: 'ID Siswa tidak terdaftar!' });
        }

        const siswa = siswaRes.rows[0];

        await pool.query(
            `INSERT INTO absensi (siswa_id, scanned_by) VALUES ($1, $2)`, 
            [siswa.id, waliKelasId]
        );

        const sock = waSockets[waliKelasId];
        let waTerkirim = false;

        if (sock && waStatus[waliKelasId] === 'TERHUBUNG') {
            try {
                let noWA = siswa.nomor_wa_ortu.replace(/[^0-9]/g, '');
                if (noWA.startsWith('0')) noWA = '62' + noWA.slice(1);
                const jid = `${noWA}@s.whatsapp.net`;

                const waktuFormat = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
                const pesan = `*NOTIFIKASI ABSENSI SEKOLAH*\n\nYth. Bapak/Ibu Orang Tua/Wali,\n\nKami menginformasikan bahwa siswa:\n*Nama*: ${siswa.nama}\n*Kelas*: ${siswa.nama_kelas}\n*Waktu Masuk*: ${waktuFormat} WIB\n*Status*: HADIR ✅\n\nTerima kasih.`;

                await sock.sendMessage(jid, { text: pesan });
                waTerkirim = true;
            } catch (waErr) {
                console.error("Gagal kirim WA:", waErr.message);
            }
        }

        return res.json({
            success: true,
            message: `Absen Berhasil: ${siswa.nama} (${siswa.nama_kelas}) ${waTerkirim ? '📲 WA Terkirim' : '⚠️ WA Gagal Terkirim'}`
        });

    } catch (err) {
        return res.json({ success: false, message: 'Terjadi kesalahan sistem.' });
    }
});

app.listen(port, () => {
    console.log(`Server Absensi aktif di port ${port}`);
});
