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

const sessionsDir = path.join(__dirname, 'wa_sessions');
if (!fs.existsSync(sessionsDir)) fs.mkdirSync(sessionsDir, { recursive: true });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));

const waSessions = {};
const qrCodes = {};
const waStatus = {};

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

// ---------------- ROUTES SYSTEM ---------------- //

app.get('/', (req, res) => res.render('login', { error: null }));

app.post('/login', async (req, res) => {
    const { username, password } = req.body;
    try {
        if (!username || !password) {
            return res.render('login', { error: 'Username dan password wajib diisi!' });
        }

        const result = await pool.query(
            'SELECT * FROM users WHERE LOWER(username) = LOWER($1) AND password = $2',
            [username.trim(), password.trim()]
        );

        if (result.rows.length === 0) {
            return res.render('login', { error: 'Username atau password salah.' });
        }

        const user = result.rows[0];
        if (user.role === 'ADMIN') {
            return res.redirect(`/admin?userId=${user.id}`);
        } else {
            return res.redirect(`/wali?userId=${user.id}`);
        }
    } catch (err) {
        return res.render('login', { error: 'Gangguan Database: ' + err.message });
    }
});

app.get('/admin', async (req, res) => {
    const userId = parseInt(req.query.userId) || 1;
    try {
        const usersRes = await pool.query(`
            SELECT u.id, u.nama, u.username, u.role, COALESCE(k.nama_kelas, 'Guru Mapel') AS nama_kelas 
            FROM users u LEFT JOIN kelas k ON u.kelas_id = k.id ORDER BY u.id ASC
        `);
        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, 'Tanpa Kelas') AS nama_kelas 
            FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id ORDER BY s.id ASC
        `);
        const kelasRes = await pool.query(`SELECT * FROM kelas ORDER BY id ASC`);

        const siswaData = await Promise.all(siswaRes.rows.map(async (s) => {
            const qrImage = await QRCode.toDataURL(s.id.toString());
            return { ...s, qrImage };
        }));

        res.render('admin-dashboard', {
            users: usersRes.rows || [],
            siswa: siswaData || [],
            kelas: kelasRes.rows || [],
            userId: userId
        });
    } catch (err) {
        res.status(500).send("Database Error: " + err.message);
    }
});

app.get('/wali', async (req, res) => {
    const userId = parseInt(req.query.userId) || 1;
    try {
        const userRes = await pool.query(`
            SELECT u.id, u.nama, u.role, u.kelas_id, COALESCE(k.nama_kelas, 'Semua Kelas') AS nama_kelas 
            FROM users u LEFT JOIN kelas k ON u.kelas_id = k.id WHERE u.id = $1
        `, [userId]);

        const user = userRes.rows[0] || { id: userId, nama: 'Pendidik', role: 'WALI_KELAS', nama_kelas: 'Semua Kelas' };
        const targetKelasId = user.kelas_id;

        let siswaQuery = `SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, 'Tanpa Kelas') AS nama_kelas FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id`;
        let absensiQuery = `SELECT a.waktu, s.nama FROM absensi a JOIN siswa s ON a.siswa_id = s.id WHERE DATE(a.waktu) = CURRENT_DATE`;

        if (targetKelasId) {
            siswaQuery += ` WHERE s.kelas_id = ${parseInt(targetKelasId)}`;
            absensiQuery += ` AND s.kelas_id = ${parseInt(targetKelasId)}`;
        }
        siswaQuery += ` ORDER BY s.nama ASC`;
        absensiQuery += ` ORDER BY a.waktu DESC`;

        const siswaKelasRes = await pool.query(siswaQuery);
        const absensiRes = await pool.query(absensiQuery);

        const siswaData = await Promise.all(siswaKelasRes.rows.map(async (s) => {
            const qrImage = await QRCode.toDataURL(s.id.toString());
            return { ...s, qrImage };
        }));

        res.render('walikelas-dashboard', {
            user: user,
            siswaList: siswaData || [],
            absensiHariIni: absensiRes.rows || [],
            userId: userId,
            statusWA: waStatus[userId] || 'BELUM_TERHUBUNG',
            qrCodeWA: qrCodes[userId] || null
        });
    } catch (err) {
        res.status(500).send("Database Error: " + err.message);
    }
});

app.get(['/scan', '/scanner'], (req, res) => {
    const userId = parseInt(req.query.userId) || 1;
    res.render('scan', { userId: userId });
});

app.get('/api/start-wa', (req, res) => {
    const userId = parseInt(req.query.userId) || 1;
    connectToWhatsApp(userId);
    res.json({ success: true, message: 'Menghubungkan WhatsApp...' });
});

app.post('/api/scan', async (req, res) => {
    const { siswa_id, scanned_by } = req.body;
    if (!siswa_id) return res.status(400).json({ success: false, message: "ID Siswa tidak ditemukan." });

    try {
        const parsedSiswaId = parseInt(siswa_id);
        const parsedScannedBy = parseInt(scanned_by) || 1;

        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, 'Tanpa Kelas') AS nama_kelas 
            FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id WHERE s.id = $1
        `, [parsedSiswaId]);

        if (siswaRes.rows.length === 0) {
            return res.status(404).json({ success: false, message: `Siswa ID #${siswa_id} Tidak Ditemukan!` });
        }

        const siswa = siswaRes.rows[0];
        await pool.query(`INSERT INTO absensi (siswa_id, status, scanned_by) VALUES ($1, 'HADIR', $2)`, [siswa.id, parsedScannedBy]);

        const waClient = waSessions[parsedScannedBy];
        if (waClient && siswa.nomor_wa_ortu) {
            let phone = siswa.nomor_wa_ortu.replace(/[^0-9]/g, '');
            if (phone.startsWith('0')) phone = '62' + phone.slice(1);
            if (!phone.endsWith('@s.whatsapp.net')) phone += '@s.whatsapp.net';

            const jam = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }) + ' WIB';
            const message = `*UPTD SD NEGERI 1 KARYA MULYA SARI*\n` +
                            `--------------------------------------------------\n` +
                            `*NOTIFIKASI KEHADIRAN SISWA*\n\n` +
                            `Yth. Orang Tua / Wali Murid dari:\n` +
                            `• Nama Siswa: *${siswa.nama}*\n` +
                            `• Kelas: *${siswa.nama_kelas}*\n` +
                            `• Waktu Scan: *${jam}*\n` +
                            `• Status: *HADIR DI SEKOLAH ✅*\n\n` +
                            `Terima kasih. Pesan ini dikirim secara otomatis oleh Sistem Presensi Digital UPTD SD Negeri 1 Karya Mulya Sari.`;

            waClient.sendMessage(phone, { text: message }).catch(e => console.error("WA Send Error:", e));
        }

        return res.json({
            success: true,
            message: "Presensi Berhasil!",
            siswa: { id: siswa.id, nama: siswa.nama, nama_kelas: siswa.nama_kelas }
        });
    } catch (err) {
        return res.status(500).json({ success: false, message: "Server Error: " + err.message });
    }
});

app.listen(PORT, () => console.log(`🚀 Server Berjalan di Port ${PORT}`));
