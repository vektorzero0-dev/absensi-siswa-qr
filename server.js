const express = require('express');
const path = require('path');
const fs = require('fs');
const pino = require('pino');
const makeWASocket = require('@whiskeysockets/baileys').default;
const { useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const QRCode = require('qrcode');
const pool = require('./db');

// Set Timezone global Node.js ke WIB
process.env.TZ = 'Asia/Jakarta';

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
                console.log(`✅ WA User #${userId} TERHUBUNG`);
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

// ---------------- ROUTES ---------------- //

app.get('/', (req, res) => res.render('login', { error: null }));

app.post('/login', async (req, res) => {
    const { username, password } = req.body;
    try {
        if (!username || !password) return res.render('login', { error: 'Username dan password wajib diisi!' });

        const result = await pool.query(
            'SELECT * FROM users WHERE LOWER(username) = LOWER($1) AND password = $2',
            [username.trim(), password.trim()]
        );

        if (result.rows.length === 0) return res.render('login', { error: 'Username atau password salah.' });

        const user = result.rows[0];
        if (user.role === 'ADMIN') {
            return res.redirect(`/admin?userId=${user.id}`);
        } else {
            return res.redirect(`/wali?userId=${user.id}`);
        }
    } catch (err) {
        return res.render('login', { error: 'Database Error: ' + err.message });
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
        
        // Memastikan Tanggal Hari Ini Menggunakan Zona Waktu WIB (Asia/Jakarta)
        let absensiQuery = `
            SELECT a.id, a.waktu, s.nama, COALESCE(k.nama_kelas, 'Tanpa Kelas') AS nama_kelas 
            FROM absensi a 
            JOIN siswa s ON a.siswa_id = s.id 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            WHERE DATE(a.waktu AT TIME ZONE 'UTC' AT TIME ZONE 'Asia/Jakarta') = CURRENT_DATE
        `;

        if (targetKelasId) {
            siswaQuery += ` WHERE s.kelas_id = ${parseInt(targetKelasId)}`;
            absensiQuery += ` AND s.kelas_id = ${parseInt(targetKelasId)}`;
        }
        siswaQuery += ` ORDER BY s.nama ASC`;
        absensiQuery += ` ORDER BY a.waktu DESC`;

        const siswaKelasRes = await pool.query(siswaQuery);
        const absensiRes = await pool.query(absensiQuery);

        // Format waktu absensi ke format Jam:Menit WIB
        const absensiFormatted = absensiRes.rows.map(row => {
            const dateObj = new Date(row.waktu);
            const waktuWIB = dateObj.toLocaleTimeString('id-ID', {
                timeZone: 'Asia/Jakarta',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            }) + ' WIB';

            return {
                ...row,
                waktu_formatted: waktuWIB
            };
        });

        const siswaData = await Promise.all(siswaKelasRes.rows.map(async (s) => {
            const qrImage = await QRCode.toDataURL(s.id.toString());
            return { ...s, qrImage };
        }));

        res.render('walikelas-dashboard', {
            user: user,
            siswaList: siswaData || [],
            absensiHariIni: absensiFormatted || [],
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

// API SCAN QR CODE / INPUT MANUAL
app.post('/api/scan', async (req, res) => {
    const { siswa_id, scanned_by } = req.body;
    if (!siswa_id) return res.status(400).json({ success: false, message: "ID Siswa wajib diisi!" });

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

        // 1. Simpan Ke Database (Gunakan NOW() server)
        await pool.query(
            `INSERT INTO absensi (siswa_id, status, scanned_by, waktu) VALUES ($1, 'HADIR', $2, NOW())`, 
            [siswa.id, parsedScannedBy]
        );

        // 2. Ambil WhatsApp Session yang Aktif
        let waClient = waSessions[parsedScannedBy];
        if (!waClient) {
            const keys = Object.keys(waSessions);
            if (keys.length > 0) waClient = waSessions[keys[0]];
        }

        let waNote = "WA Tidak Terkirim (Belum Konek)";

        // 3. Kirim WhatsApp Jika Ada Koneksi & Nomor HP Ortu
        if (waClient && siswa.nomor_wa_ortu) {
            let phone = siswa.nomor_wa_ortu.toString().trim().replace(/[^0-9]/g, '');
            if (phone.startsWith('0')) phone = '62' + phone.slice(1);
            const targetJid = phone + '@s.whatsapp.net';

            const now = new Date();
            const jamStr = now.toLocaleTimeString('id-ID', { timeZone: 'Asia/Jakarta', hour: '2-digit', minute: '2-digit' }) + ' WIB';
            const tglStr = now.toLocaleDateString('id-ID', { timeZone: 'Asia/Jakarta', weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

            const msg = `*UPTD SD NEGERI 1 KARYA MULYA SARI*\n` +
                        `----------------------------------------\n` +
                        `*PRESENSI SISWA BERHASIL*\n\n` +
                        `Nama: *${siswa.nama}*\n` +
                        `Kelas: *${siswa.nama_kelas}*\n` +
                        `Hari/Tgl: *${tglStr}*\n` +
                        `Jam: *${jamStr}*\n` +
                        `Status: *HADIR ✅*\n\n` +
                        `_Pesan otomatis dari Sistem Presensi Sekolah._`;

            // Fire and forget / handle error tanpa menggagalkan respon scanner
            waClient.sendMessage(targetJid, { text: msg }).then(() => {
                console.log(`✅ Pesan WA terkirim ke ${phone}`);
            }).catch(e => {
                console.error(`❌ Gagal kirim WA:`, e.message);
            });

            waNote = "WA Terkirim ✅";
        }

        return res.json({
            success: true,
            message: `Presensi Berhasil! (${waNote})`,
            siswa: { id: siswa.id, nama: siswa.nama, nama_kelas: siswa.nama_kelas }
        });

    } catch (err) {
        console.error("Scan Error:", err);
        return res.status(500).json({ success: false, message: "Terjadi kesalahan: " + err.message });
    }
});

app.listen(PORT, () => console.log(`🚀 Server berjalan pada port ${PORT}`));
