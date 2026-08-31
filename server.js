/**
 * ============================================================================
 * SISTEM ABSENSI SEKOLAH AUTOMATION ENGINE
 * Instansi      : UPTD SD NEGERI 1 KARYA MULYA SARI
 * Developed by  : Zeeo
 * Email         : vektorzero0@gmail.com
 * WhatsApp      : 082371729760
 * ============================================================================
 */

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

const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: { rejectUnauthorized: false }
});

const waSockets = {};
const qrCodes = {};
const waStatus = {};

async function initWA(userId) {
    if (waSockets[userId]) return;

    waStatus[userId] = 'MENYIAPKAN';

    try {
        const { state, saveCreds } = await useMultiFileAuthState(`baileys_session_user_${userId}`);

        const sock = makeWASocket({
            auth: state,
            printQRInTerminal: false,
            browser: ["SDN 1 Karya Mulya Sari", "Chrome", "1.0.0"]
        });

        sock.ev.on('creds.update', saveCreds);

        sock.ev.on('connection.update', (update) => {
            const { connection, lastDisconnect, qr } = update;

            if (qr) {
                qrcode.toDataURL(qr, (err, url) => {
                    if (!err) {
                        qrCodes[userId] = url;
                        waStatus[userId] = 'BELUM_TERHUBUNG';
                    }
                });
            }

            if (connection === 'open') {
                waStatus[userId] = 'TERHUBUNG';
                qrCodes[userId] = null;
            }

            if (connection === 'close') {
                const statusCode = (lastDisconnect?.error)?.output?.statusCode;
                const shouldReconnect = statusCode !== DisconnectReason.loggedOut;
                
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
                   COALESCE(k.nama_kelas, 'Guru Mapel / Staf') AS nama_kelas
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
        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, 'Tanpa Kelas') AS nama_kelas 
            FROM siswa s 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            ORDER BY k.id ASC, s.nama ASC
        `);
        
        const usersRes = await pool.query(`
            SELECT u.id, u.nama, u.username, u.role, COALESCE(k.nama_kelas, 'Guru Mapel / Admin') AS nama_kelas 
            FROM users u 
            LEFT JOIN kelas k ON u.kelas_id = k.id 
            ORDER BY u.id ASC
        `);
        
        const kelasRes = await pool.query(`
            SELECT k.id, k.nama_kelas, COUNT(s.id) AS jumlah_siswa 
            FROM kelas k 
            LEFT JOIN siswa s ON s.kelas_id = k.id 
            GROUP BY k.id, k.nama_kelas 
            ORDER BY k.id ASC
        `);

        const absensiRes = await pool.query(`
            SELECT a.id, a.siswa_id, a.waktu, a.status, s.nama, COALESCE(k.nama_kelas, '-') AS nama_kelas 
            FROM absensi a 
            JOIN siswa s ON a.siswa_id = s.id 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            WHERE DATE(a.waktu) = CURRENT_DATE 
            ORDER BY a.waktu DESC
        `);

        res.render('admin-dashboard', {
            users: usersRes.rows,
            siswa: siswaRes.rows,
            kelas: kelasRes.rows,
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
        const userRes = await pool.query(`
            SELECT u.id, u.nama, u.role, u.kelas_id,
                   COALESCE(k.nama_kelas, 'Guru Mapel') AS nama_kelas
            FROM users u
            LEFT JOIN kelas k ON u.kelas_id = k.id
            WHERE u.id = $1
        `, [userId]);

        if (userRes.rows.length === 0) return res.redirect('/');

        const user = userRes.rows[0];

        // Ambil data siswa di kelas wali kelas tersebut
        const siswaKelasRes = await pool.query(`
            SELECT id, nama, nomor_wa_ortu FROM siswa WHERE kelas_id = $1 ORDER BY nama ASC
        `, [user.kelas_id]);

        // Ambil riwayat absen kelas hari ini
        const absensiHariIniRes = await pool.query(`
            SELECT a.waktu, s.nama 
            FROM absensi a 
            JOIN siswa s ON a.siswa_id = s.id 
            WHERE s.kelas_id = $1 AND DATE(a.waktu) = CURRENT_DATE 
            ORDER BY a.waktu DESC
        `, [user.kelas_id]);

        res.render('walikelas-dashboard', {
            user: user,
            siswaList: siswaKelasRes.rows,
            absensiHariIni: absensiHariIniRes.rows,
            userId: userId,
            statusWA: waStatus[userId] || 'BELUM_TERHUBUNG',
            qrCodeWA: qrCodes[userId] || null
        });
    } catch (err) {
        res.status(500).send("Gagal memuat Dashboard Guru.");
    }
});

app.get('/api/start-wa', (req, res) => {
    const userId = req.query.userId;
    if (userId) initWA(userId);
    res.json({ status: 'PROSES' });
});

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
                const pesan = `*UPTD SD NEGERI 1 KARYA MULYA SARI*\n\n*NOTIFIKASI KEHADIRAN SISWA*\n\nYth. Bapak/Ibu Orang Tua/Wali,\n\nInformasi absensi putra/putri Anda:\n*Nama*: ${siswa.nama}\n*Kelas*: ${siswa.nama_kelas}\n*Waktu Masuk*: ${waktuFormat} WIB\n*Status*: HADIR ✅\n\n_Pesan otomatis dari Sistem Absensi Digital Sekolah._`;

                await sock.sendMessage(jid, { text: pesan });
                waTerkirim = true;
            } catch (waErr) {
                console.error("Gagal kirim WA:", waErr.message);
            }
        }

        return res.json({
            success: true,
            message: `Absen Berhasil: ${siswa.nama} (${siswa.nama_kelas}) ${waTerkirim ? '📲 Notifikasi WA Terkirim' : '⚠️ WA Belum Terkoneksi'}`
        });

    } catch (err) {
        return res.json({ success: false, message: 'Terjadi kesalahan sistem.' });
    }
});

app.listen(port, () => {
    console.log(`Server Absensi Aktif di Port ${port}`);
});
