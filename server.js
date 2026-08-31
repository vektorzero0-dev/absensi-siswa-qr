const express = require('express');
const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode');
const { Pool } = require('pg');
const path = require('path');

const app = express();
const port = process.env.PORT || 10000;

// Setup EJS & Body Parser
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Koneksi Database PostgreSQL (Neon.tech)
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: { rejectUnauthorized: false }
});

// Penyimpanan Sesi WhatsApp & QR Per User
const waClients = {};
const qrCodes = {};
const waStatus = {};

// Fungsi Inisialisasi WhatsApp Web per User (Aman untuk RAM Terbatas di Cloud)
function initWA(userId) {
    if (waClients[userId]) return;

    console.log(`[WA] Menyiapkan koneksi WhatsApp untuk User ID: ${userId}...`);
    waStatus[userId] = 'MENYIAPKAN';
    
    try {
        const client = new Client({
            authStrategy: new LocalAuth({ clientId: `session_user_${userId}` }),
            puppeteer: {
                headless: true,
                args: [
                    '--no-sandbox',
                    '--disable-setuid-sandbox',
                    '--disable-dev-shm-usage',
                    '--disable-accelerated-2d-canvas',
                    '--no-first-run',
                    '--no-zygote',
                    '--single-process',
                    '--disable-gpu'
                ]
            }
        });

        client.on('qr', (qr) => {
            qrcode.toDataURL(qr, (err, url) => {
                if (!err) {
                    qrCodes[userId] = url;
                    waStatus[userId] = 'BELUM_TERHUBUNG';
                    console.log(`[WA] QR Code siap untuk User ID: ${userId}`);
                }
            });
        });

        client.on('ready', () => {
            console.log(`[WA] User ID ${userId} Berhasil Terhubung!`);
            waStatus[userId] = 'TERHUBUNG';
            qrCodes[userId] = null;
        });

        client.on('disconnected', (reason) => {
            console.log(`[WA] User ID ${userId} Terputus! Alasan: ${reason}`);
            waStatus[userId] = 'BELUM_TERHUBUNG';
            delete waClients[userId];
            delete qrCodes[userId];
        });

        // Mencegah Server Crash (502 Error) jika Puppeteer gagal
        client.initialize().catch(err => {
            console.error(`[WA Error] Gagal inisialisasi Puppeteer untuk User ID ${userId}:`, err.message);
            waStatus[userId] = 'BELUM_TERHUBUNG';
            delete waClients[userId];
        });

        waClients[userId] = client;
    } catch (err) {
        console.error(`[WA Error] Gagal membuat instance WA untuk User ID ${userId}:`, err.message);
        waStatus[userId] = 'BELUM_TERHUBUNG';
    }
}

// ------------------- ROUTES / HALAMAN -------------------

// 1. Halaman Login Utama
app.get('/', (req, res) => {
    res.render('login', { error: null });
});

// 2. Proses Login
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

        // Jika Admin Utama
        if (user.role === 'ADMIN_OPS') {
            return res.redirect(`/admin?userId=${user.id}`);
        }

        // Jika Wali Kelas atau Guru Mapel
        initWA(user.id);
        return res.redirect(`/wali?userId=${user.id}`);

    } catch (err) {
        console.error("Login Error:", err);
        res.render('login', { error: 'Terjadi kesalahan sistem database.' });
    }
});

// 3. Dashboard Admin
app.get('/admin', async (req, res) => {
    const userId = req.query.userId;
    try {
        // Ambil Data Siswa + Nama Kelas
        const siswaQuery = `
            SELECT s.id, s.nama, s.nomor_wa_ortu, k.nama_kelas 
            FROM siswa s
            LEFT JOIN kelas k ON s.kelas_id = k.id
            ORDER BY k.id ASC, s.nama ASC
        `;
        const siswaRes = await pool.query(siswaQuery);

        // Ambil Data Users / Guru + Nama Kelas
        const usersQuery = `
            SELECT u.id, u.nama, u.username, u.role, 
                   COALESCE(k.nama_kelas, 'Guru Mapel / Admin') AS nama_kelas
            FROM users u
            LEFT JOIN kelas k ON u.kelas_id = k.id
            ORDER BY u.id ASC
        `;
        const usersRes = await pool.query(usersQuery);

        // Ambil Log Absensi Hari Ini
        const absensiQuery = `
            SELECT a.id, a.siswa_id, a.waktu, a.status, s.nama, k.nama_kelas
            FROM absensi a
            JOIN siswa s ON a.siswa_id = s.id
            LEFT JOIN kelas k ON s.kelas_id = k.id
            WHERE DATE(a.waktu) = CURRENT_DATE
            ORDER BY a.waktu DESC
        `;
        const absensiRes = await pool.query(absensiQuery);

        res.render('admin-dashboard', {
            users: usersRes.rows,
            siswa: siswaRes.rows,
            absensi: absensiRes.rows,
            userId: userId
        });

    } catch (err) {
        console.error("Admin Dashboard Error:", err);
        res.status(500).send("Gagal memuat Dashboard Admin.");
    }
});

// 4. Dashboard Wali Kelas & Guru Mapel
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

        const user = result.rows[0];

        // Pastikan proses WA berjalan tanpa menghentikan rendering halaman
        initWA(user.id);

        res.render('walikelas-dashboard', {
            user: user,
            userId: user.id,
            statusWA: waStatus[user.id] || 'BELUM_TERHUBUNG',
            qrCodeWA: qrCodes[user.id] || null
        });

    } catch (err) {
        console.error("Wali Dashboard Error:", err);
        res.status(500).send("Gagal memuat Dashboard Wali Kelas.");
    }
});

// 5. Halaman QR Scanner
app.get('/scan', (req, res) => {
    const userId = req.query.userId;
    if (!userId) return res.redirect('/');
    res.render('scan', { userId: userId });
});

// 6. API Scan Absensi & Kirim Notifikasi WA
app.post('/api/absen', async (req, res) => {
    const { siswaId, waliKelasId } = req.body;

    try {
        // Cari Siswa
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

        // Simpan Log Absensi ke DB
        await pool.query(
            `INSERT INTO absensi (siswa_id, scanned_by) VALUES ($1, $2)`, 
            [siswa.id, waliKelasId]
        );

        // Kirim WhatsApp jika Client WA terhubung
        const clientWA = waClients[waliKelasId];
        let waTerkirim = false;

        if (clientWA && waStatus[waliKelasId] === 'TERHUBUNG') {
            try {
                let noWA = siswa.nomor_wa_ortu.replace(/[^0-9]/g, '');
                if (noWA.startsWith('0')) noWA = '62' + noWA.slice(1);
                if (!noWA.endsWith('@c.us')) noWA += '@c.us';

                const waktuFormat = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
                const pesan = `*NOTIFIKASI ABSENSI SEKOLAH*\n\nYth. Bapak/Ibu Orang Tua/Wali,\n\nKami menginformasikan bahwa siswa:\n*Nama*: ${siswa.nama}\n*Kelas*: ${siswa.nama_kelas}\n*Waktu Masuk*: ${waktuFormat} WIB\n*Status*: HADIR ✅\n\nTerima kasih.`;

                await clientWA.sendMessage(noWA, pesan);
                waTerkirim = true;
            } catch (waErr) {
                console.error("Gagal mengirim WA:", waErr.message);
            }
        }

        return res.json({
            success: true,
            message: `Absen Berhasil: ${siswa.nama} (${siswa.nama_kelas}) ${waTerkirim ? '📲 WA Terkirim' : '⚠️ WA Tidak Terkirim'}`
        });

    } catch (err) {
        console.error("API Absen Error:", err);
        return res.json({ success: false, message: 'Terjadi kesalahan sistem.' });
    }
});

app.listen(port, () => {
    console.log(`Server Absensi aktif di port ${port}`);
});
