const express = require('express');
const session = require('express-session');
const path = require('path');
const pino = require('pino');
const makeWASocket = require('@whiskeysockets/baileys').default;
const { DisconnectReason, useMultiFileAuthState, fetchLatestBaileysVersion } = require('@whiskeysockets/baileys');
const QRCode = require('qrcode');
const qrcodeTerminal = require('qrcode-terminal');
const fs = require('fs');
const pool = require('./db');
const cron = require('node-cron');

// Package Import & Upload Excel
const multer = require('multer');
const XLSX = require('xlsx');
const upload = multer({ storage: multer.memoryStorage() });

// Package Ekspor Laporan
const ExcelJS = require('exceljs');
const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, AlignmentType } = require('docx');
const PDFDocument = require('pdfkit');

process.env.TZ = 'Asia/Jakarta';

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));

app.use(session({
    secret: 'secret-key-presensi-sd',
    resave: false,
    saveUninitialized: true,
    cookie: { secure: false }
}));

// ----------------- AUTO-CREATE & MIGRATE TABEL DATABASE MULTI-TENANT ----------------- //
async function initDB() {
    try {
        // 1. Tabel Sekolah (Dengan Kolom is_active)
        await pool.query(`
            CREATE TABLE IF NOT EXISTS sekolah (
                id SERIAL PRIMARY KEY,
                nama_sekolah VARCHAR(100) NOT NULL,
                is_active BOOLEAN DEFAULT TRUE
            );
            ALTER TABLE sekolah ADD COLUMN IF NOT EXISTS is_active BOOLEAN DEFAULT TRUE;
        `);

        // 2. Tabel Kelas (Terhubung ke Sekolah)
        await pool.query(`
            CREATE TABLE IF NOT EXISTS kelas (
                id SERIAL PRIMARY KEY,
                nama_kelas VARCHAR(50) NOT NULL,
                sekolah_id INT REFERENCES sekolah(id) ON DELETE CASCADE
            );
            ALTER TABLE kelas ADD COLUMN IF NOT EXISTS sekolah_id INT REFERENCES sekolah(id) ON DELETE CASCADE;
        `);

        // 3. Tabel Users (Mendukung SUPER_ADMIN, ADMIN, WALI_KELAS)
        await pool.query(`
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                nama VARCHAR(100) NOT NULL,
                username VARCHAR(50) UNIQUE NOT NULL,
                password VARCHAR(255) NOT NULL,
                role VARCHAR(20) NOT NULL DEFAULT 'WALI_KELAS',
                kelas_id INT REFERENCES kelas(id) ON DELETE SET NULL,
                sekolah_id INT REFERENCES sekolah(id) ON DELETE CASCADE
            );
            ALTER TABLE users ADD COLUMN IF NOT EXISTS sekolah_id INT REFERENCES sekolah(id) ON DELETE CASCADE;
        `);

        // 4. Tabel Siswa
        await pool.query(`
            CREATE TABLE IF NOT EXISTS siswa (
                id SERIAL PRIMARY KEY,
                nama VARCHAR(100) NOT NULL,
                nomor_wa_ortu VARCHAR(20),
                kelas_id INT REFERENCES kelas(id) ON DELETE SET NULL
            );
            ALTER TABLE siswa ADD COLUMN IF NOT EXISTS nomor_wa_ortu VARCHAR(20);
        `);

        // 5. Tabel Absensi
        await pool.query(`
            CREATE TABLE IF NOT EXISTS absensi (
                id SERIAL PRIMARY KEY,
                siswa_id INT REFERENCES siswa(id) ON DELETE CASCADE,
                waktu TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status VARCHAR(20) DEFAULT 'HADIR',
                scanned_by INT,
                tipe VARCHAR(10) DEFAULT 'MASUK'
            );
            ALTER TABLE absensi ADD COLUMN IF NOT EXISTS tipe VARCHAR(10) DEFAULT 'MASUK';
            ALTER TABLE absensi ADD COLUMN IF NOT EXISTS scanned_by INT;
        `);

        // 6. Tabel Settings (Lama) & Migrasi ke Sekolah
        await pool.query(`
            CREATE TABLE IF NOT EXISTS settings (
                key VARCHAR(50) PRIMARY KEY,
                value VARCHAR(255)
            );
            INSERT INTO settings (key, value) 
            VALUES ('pengirim_wa', 'ADMIN'), ('nama_sekolah', 'NAMA SEKOLAH BELUM DIATUR')
            ON CONFLICT (key) DO NOTHING;
        `);

        // Inisialisasi Sekolah Pertama dari DB / ENV
        const currentSekolahId = parseInt(process.env.SEKOLAH_ID) || 1;
        const oldSetting = await pool.query("SELECT value FROM settings WHERE key = 'nama_sekolah'");
        const defaultNama = oldSetting.rows.length > 0 ? oldSetting.rows[0].value : 'SEKOLAH UTAMA';

        await pool.query(`
            INSERT INTO sekolah (id, nama_sekolah, is_active) VALUES ($1, $2, TRUE)
            ON CONFLICT (id) DO NOTHING;
            SELECT setval('sekolah_id_seq', (SELECT GREATEST(MAX(id), 1) FROM sekolah));
        `, [currentSekolahId, defaultNama]);

        // Paksa buat/update akun superadmin berdasarkan USERNAME
        await pool.query(`
            INSERT INTO users (nama, username, password, role, sekolah_id)
            VALUES ('Super Administrator', 'superadmin', 'super123', 'SUPER_ADMIN', NULL)
            ON CONFLICT (username) 
            DO UPDATE SET password = 'super123', role = 'SUPER_ADMIN';
        `);

        // Paksa buat/update akun admin sekolah berdasarkan USERNAME
        await pool.query(`
            INSERT INTO users (nama, username, password, role, sekolah_id)
            VALUES ('Admin Sekolah', 'admin', 'admin123', 'ADMIN', $1)
            ON CONFLICT (username) 
            DO UPDATE SET password = 'admin123', role = 'ADMIN', sekolah_id = $1;
        `, [currentSekolahId]);

        console.log("✅ Database Multi-Tenant Initialized: SUPER_ADMIN & Multi-Sekolah Siap!");
    } catch (err) {
        console.error("❌ Gagal inisialisasi/migrasi database:", err.message);
    }
}
initDB();

const waSessions = {};
const qrCodes = {};
const waStatus = {};
const pairingCodes = {};
const reconnectTimers = {};

function bersihkanGelar(nama) {
    if (!nama) return '';
    return nama.replace(/,?\s*\b(S\.Pd|M\.Pd|S\.Ag|S\.T|S\.Kom|M\.Si|S\.Sos|S\.SE|M\.M|A\.Ma|Sd)\b\.?/gi, '').trim();
}

async function getNamaSekolah(sekolahId = null) {
    try {
        const targetId = sekolahId || process.env.SEKOLAH_ID || 1;
        const res = await pool.query("SELECT nama_sekolah FROM sekolah WHERE id = $1", [targetId]);
        if (res.rows.length > 0) return res.rows[0].nama_sekolah;

        const oldRes = await pool.query("SELECT value FROM settings WHERE key = 'nama_sekolah'");
        return oldRes.rows.length > 0 ? oldRes.rows[0].value : 'NAMA SEKOLAH BELUM DIATUR';
    } catch (err) {
        return 'NAMA SEKOLAH BELUM DIATUR';
    }
}

async function generateQRDataURL(text) {
    try {
        if (!text) return 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORTH5CYII=';
        return await QRCode.toDataURL(text.toString());
    } catch (err) {
        return 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORTH5CYII=';
    }
}

async function getAuthState(userId) {
    const authFolder = path.join(__dirname, 'auth_sessions', `user_${userId}`);
    if (!fs.existsSync(authFolder)) {
        fs.mkdirSync(authFolder, { recursive: true });
    }
    return await useMultiFileAuthState(authFolder);
}

async function connectToWhatsApp(userId, phoneNumber = null) {
    try {
        if (reconnectTimers[userId]) {
            clearTimeout(reconnectTimers[userId]);
            delete reconnectTimers[userId];
        }

        if (waSessions[userId]) {
            try { waSessions[userId].end(undefined); } catch (e) {}
            delete waSessions[userId];
        }

        waStatus[userId] = phoneNumber ? 'MENUNGGU_PAIRING_CODE' : (pairingCodes[userId] ? 'MENUNGGU_PAIRING_CODE' : 'PROSES_INIT');
        delete qrCodes[userId];

        if (phoneNumber) delete pairingCodes[userId];

        const { state, saveCreds } = await getAuthState(userId);
        const { version } = await fetchLatestBaileysVersion();

        console.log(`⚡ [User #${userId}] Inisialisasi WA Socket (Baileys v${version.join('.')})...`);

        const sock = makeWASocket({
            logger: pino({ level: 'silent' }),
            auth: state,
            printQRInTerminal: false,
            browser: ["Ubuntu", "Chrome", "120.0.0.0"],
            connectTimeoutMs: 60000,
            defaultQueryTimeoutMs: 60000,
            keepAliveIntervalMs: 25000,
            qrTimeout: 45000,
            syncFullHistory: false
        });

        waSessions[userId] = sock;
        sock.ev.on('creds.update', saveCreds);

        if (phoneNumber && !sock.authState.creds.registered) {
            setTimeout(async () => {
                try {
                    let cleanPhone = phoneNumber.toString().replace(/[^0-9]/g, '');
                    if (cleanPhone.startsWith('0')) cleanPhone = '62' + cleanPhone.slice(1);
                    
                    const code = await sock.requestPairingCode(cleanPhone);
                    pairingCodes[userId] = code;
                    waStatus[userId] = 'MENUNGGU_PAIRING_CODE';

                    setTimeout(() => {
                        if (waStatus[userId] !== 'TERHUBUNG') delete pairingCodes[userId];
                    }, 180000);
                } catch (pErr) {
                    waStatus[userId] = 'ERROR_PAIRING';
                }
            }, 5000);
        }

        sock.ev.on('connection.update', async (update) => {
            const { connection, lastDisconnect, qr } = update;

            if (qr && !phoneNumber && !sock.authState.creds.registered) {
                try {
                    qrCodes[userId] = await generateQRDataURL(qr);
                    waStatus[userId] = 'MENUNGGU_SCAN';
                    qrcodeTerminal.generate(qr, { small: true });
                } catch (qrErr) {}
            }

            if (connection === 'open') {
                waStatus[userId] = 'TERHUBUNG';
                delete qrCodes[userId];
                delete pairingCodes[userId];
                if (reconnectTimers[userId]) {
                    clearTimeout(reconnectTimers[userId]);
                    delete reconnectTimers[userId];
                }
                console.log(`✅ [User #${userId}] WhatsApp Berhasil Terhubung!`);
            }

            if (connection === 'close') {
                const statusCode = lastDisconnect?.error?.output?.statusCode;
                const isLoggedOut = (statusCode === DisconnectReason.loggedOut || statusCode === 401);

                waStatus[userId] = 'TERPUTUS';
                delete waSessions[userId];

                if (!isLoggedOut) {
                    if (!reconnectTimers[userId]) {
                        reconnectTimers[userId] = setTimeout(() => {
                            delete reconnectTimers[userId];
                            connectToWhatsApp(userId);
                        }, 8000);
                    }
                } else {
                    delete qrCodes[userId];
                    delete pairingCodes[userId];
                    const authFolder = path.join(__dirname, 'auth_sessions', `user_${userId}`);
                    if (fs.existsSync(authFolder)) fs.rmSync(authFolder, { recursive: true, force: true });
                }
            }
        });
    } catch (err) {
        waStatus[userId] = 'ERROR';
    }
}

// ---------------- ROUTES HALAMAN ---------------- //

app.get(['/', '/login'], async (req, res) => {
    try {
        const namaSekolah = await getNamaSekolah();
        res.render('login', { error: null, namaSekolah });
    } catch (err) {
        res.render('login', { error: null, namaSekolah: 'NAMA SEKOLAH BELUM DIATUR' });
    }
});

app.post('/login', async (req, res) => {
    const { username, password } = req.body;
    try {
        const namaSekolah = await getNamaSekolah();
        if (!username || !password) return res.render('login', { error: 'Username dan kata sandi wajib diisi.', namaSekolah });

        // Tarik data user beserta status keaktifan sekolah
        const result = await pool.query(`
            SELECT u.*, COALESCE(s.is_active, TRUE) AS is_active 
            FROM users u 
            LEFT JOIN sekolah s ON u.sekolah_id = s.id 
            WHERE LOWER(u.username) = LOWER($1) AND u.password = $2
        `, [username.trim(), password.trim()]);

        if (result.rows.length === 0) return res.render('login', { error: 'Username atau kata sandi tidak valid.', namaSekolah });

        const user = result.rows[0];

        // Ditolak login jika sekolah NONAKTIF (kecuali SUPER_ADMIN)
        if (user.role !== 'SUPER_ADMIN' && user.is_active === false) {
            return res.render('login', { 
                error: 'Akses sekolah Anda telah dinonaktifkan oleh Super Admin.', 
                namaSekolah 
            });
        }

        req.session.userId = user.id;

        if (user.role === 'SUPER_ADMIN') {
            return res.redirect(`/superadmin?userId=${user.id}`);
        } else if (user.role === 'ADMIN') {
            return res.redirect(`/admin?userId=${user.id}`);
        } else {
            return res.redirect(`/wali?userId=${user.id}`);
        }
    } catch (err) {
        return res.render('login', { error: 'Kesalahan Sistem Database: ' + err.message, namaSekolah: 'NAMA SEKOLAH BELUM DIATUR' });
    }
});

// ----------------- DASBOR SUPER ADMIN ----------------- //
app.get('/superadmin', async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId;
    if (!userId) return res.redirect('/');

    try {
        const userRes = await pool.query('SELECT * FROM users WHERE id = $1 AND role = $2', [userId, 'SUPER_ADMIN']);
        if (userRes.rows.length === 0) return res.redirect('/');

        // Ganti query sekolahRes di dalam app.get('/superadmin', ...) di server.js
        // Ganti query sekolahRes di dalam app.get('/superadmin', ...) pada server.js
        const sekolahRes = await pool.query(`
            SELECT s.id, s.nama_sekolah, COALESCE(s.is_active, TRUE) AS is_active,
                   COUNT(DISTINCT k.id) AS total_kelas,
                   COUNT(DISTINCT sis.id) AS total_siswa,
                   COUNT(DISTINCT CASE WHEN u.role != 'SUPER_ADMIN' THEN u.id END) AS total_pengguna
             FROM sekolah s
             LEFT JOIN kelas k ON k.sekolah_id = s.id
             LEFT JOIN siswa sis ON sis.kelas_id = k.id
             LEFT JOIN users u ON u.sekolah_id = s.id
             GROUP BY s.id, s.nama_sekolah, s.is_active
             ORDER BY s.id ASC
        `);
        
        const adminRes = await pool.query(`
            SELECT u.id, u.nama, u.username, u.sekolah_id, COALESCE(s.nama_sekolah, 'Sistem Global') AS nama_sekolah
            FROM users u
            LEFT JOIN sekolah s ON u.sekolah_id = s.id
            WHERE u.role = 'ADMIN'
            ORDER BY u.id ASC
        `);

        req.session.userId = userId;

        res.render('superadmin-dashboard', {
            user: userRes.rows[0],
            sekolahList: sekolahRes.rows || [],
            adminList: adminRes.rows || [],
            userId
        });
    } catch (err) {
        res.status(500).send("Gagal Memuat Dasbor Super Admin: " + err.message);
    }
});

app.post('/api/sekolah/tambah', async (req, res) => {
    const { nama_sekolah, admin_nama, admin_username, admin_password } = req.body;
    const client = await pool.connect();
    try {
        if (!nama_sekolah || !admin_username || !admin_password) {
            return res.status(400).send("Nama Sekolah, Username Admin, dan Password wajib diisi.");
        }

        await client.query('BEGIN');
        const schRes = await client.query('INSERT INTO sekolah (nama_sekolah, is_active) VALUES ($1, TRUE) RETURNING id', [nama_sekolah.trim()]);
        const newSekolahId = schRes.rows[0].id;

        await client.query(`
            INSERT INTO users (nama, username, password, role, sekolah_id)
            VALUES ($1, $2, $3, 'ADMIN', $4)
        `, [admin_nama ? admin_nama.trim() : `Admin ${nama_sekolah}`, admin_username.trim(), admin_password.trim(), newSekolahId]);

        await client.query('COMMIT');
        return res.redirect(`/superadmin?userId=${req.session.userId || 1}`);
    } catch (err) {
        await client.query('ROLLBACK');
        return res.status(500).send("Gagal menambah sekolah baru: " + err.message);
    } finally {
        client.release();
    }
});

// TOGGLE AKTIF / NONAKTIFKAN SEKOLAH
app.post('/api/sekolah/toggle-status/:id', async (req, res) => {
    const sekolahId = parseInt(req.params.id);
    try {
        await pool.query(`
            UPDATE sekolah 
            SET is_active = NOT COALESCE(is_active, TRUE) 
            WHERE id = $1
        `, [sekolahId]);
        return res.redirect(`/superadmin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal mengubah status keaktifan sekolah: " + err.message);
    }
});

// EDIT NAMA SEKOLAH
app.post('/api/sekolah/edit/:id', async (req, res) => {
    const sekolahId = parseInt(req.params.id);
    const { nama_sekolah } = req.body;
    try {
        if (!nama_sekolah) return res.status(400).send("Nama sekolah wajib diisi.");
        await pool.query('UPDATE sekolah SET nama_sekolah = $1 WHERE id = $2', [nama_sekolah.trim(), sekolahId]);
        return res.redirect(`/superadmin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal mengedit sekolah: " + err.message);
    }
});

// HAPUS SEKOLAH
app.post('/api/sekolah/hapus/:id', async (req, res) => {
    const sekolahId = parseInt(req.params.id);
    try {
        await pool.query('DELETE FROM sekolah WHERE id = $1', [sekolahId]);
        return res.redirect(`/superadmin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal menghapus sekolah: " + err.message);
    }
});

// EDIT AKUN ADMIN SEKOLAH
app.post('/api/admin/edit/:id', async (req, res) => {
    const adminId = parseInt(req.params.id);
    const { nama, username, password } = req.body;
    try {
        if (!nama || !username) return res.status(400).send("Nama dan Username wajib diisi.");

        if (password && password.trim() !== '') {
            await pool.query(
                `UPDATE users SET nama = $1, username = $2, password = $3 WHERE id = $4 AND role = 'ADMIN'`,
                [nama.trim(), username.trim(), password.trim(), adminId]
            );
        } else {
            await pool.query(
                `UPDATE users SET nama = $1, username = $2 WHERE id = $3 AND role = 'ADMIN'`,
                [nama.trim(), username.trim(), adminId]
            );
        }
        return res.redirect(`/superadmin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal mengedit akun admin: " + err.message);
    }
});

// HAPUS AKUN ADMIN SEKOLAH
app.post('/api/admin/hapus/:id', async (req, res) => {
    const adminId = parseInt(req.params.id);
    try {
        await pool.query("DELETE FROM users WHERE id = $1 AND role = 'ADMIN'", [adminId]);
        return res.redirect(`/superadmin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal menghapus admin: " + err.message);
    }
});

// ENDPOINT AMBIL DETAIL SEKOLAH (UNTUK MODAL LIHAT/INTIP SUPER ADMIN)
app.get('/api/sekolah/detail/:id', async (req, res) => {
    const sekolahId = parseInt(req.params.id);
    try {
        const sekolah = await pool.query('SELECT nama_sekolah FROM sekolah WHERE id = $1', [sekolahId]);
        if (sekolah.rows.length === 0) return res.status(404).json({ success: false, message: 'Sekolah tidak ditemukan' });

        const pengguna = await pool.query(`
            SELECT u.nama, u.username, u.role, COALESCE(k.nama_kelas, 'Tanpa Penugasan / Guru Mapel') AS nama_kelas 
            FROM users u 
            LEFT JOIN kelas k ON u.kelas_id = k.id 
            WHERE u.sekolah_id = $1 AND u.role != 'SUPER_ADMIN'
            ORDER BY u.id ASC
        `, [sekolahId]);

        const kelas = await pool.query('SELECT nama_kelas FROM kelas WHERE sekolah_id = $1 ORDER BY id ASC', [sekolahId]);
        
        const siswa = await pool.query(`
            SELECT s.nama, COALESCE(k.nama_kelas, 'Tanpa Kelas') AS nama_kelas 
            FROM siswa s 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            WHERE k.sekolah_id = $1 
            ORDER BY s.nama ASC
        `, [sekolahId]);

        return res.json({
            success: true,
            namaSekolah: sekolah.rows[0].nama_sekolah,
            penggunaList: pengguna.rows,
            kelasList: kelas.rows,
            siswaList: siswa.rows
        });
    } catch (err) {
        return res.status(500).json({ success: false, message: err.message });
    }
});

// TOGGLE / UBAH MODE PENGIRIM WA SEKOLAH (SUPER ADMIN)
app.post('/api/sekolah/wa-mode/:id', async (req, res) => {
    const sekolahId = parseInt(req.params.id);
    const { wa_mode } = req.body;
    try {
        const validMode = ['WALI_KELAS', 'PETUGAS'].includes(wa_mode) ? wa_mode : 'WALI_KELAS';
        await pool.query('UPDATE sekolah SET wa_mode = $1 WHERE id = $2', [validMode, sekolahId]);
        return res.redirect(`/superadmin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal mengupdate mode pengirim WA: " + err.message);
    }
});

// SUPER ADMIN: TAMBAH AKUN PETUGAS ABSEN UNTUK SEKOLAH CERTAIN
app.post('/api/superadmin/petugas/tambah', async (req, res) => {
    const { sekolah_id, nama, username, password } = req.body;
    try {
        if (!sekolah_id || !nama || !username || !password) {
            return res.status(400).send("Semua kolom wajib diisi.");
        }

        // Buat akun khusus dengan role PETUGAS
        await pool.query(
            `INSERT INTO users (nama, username, password, role, sekolah_id) 
             VALUES ($1, $2, $3, 'PETUGAS', $4)`,
            [nama.trim(), username.trim(), password.trim(), parseInt(sekolah_id)]
        );

        return res.redirect(`/superadmin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal membuat akun Petugas: " + err.message);
    }
});

// ----------------- DASBOR ADMIN SEKOLAH (TERISOLASI PER SEKOLAH) ----------------- //
app.get('/admin', async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    
    try {
        const userRes = await pool.query('SELECT * FROM users WHERE id = $1', [userId]);
        if (userRes.rows.length === 0) return res.redirect('/');
        
        const currentUser = userRes.rows[0];
        const userSekolahId = currentUser.sekolah_id || parseInt(process.env.SEKOLAH_ID) || 1;

        // DIUBAH: Sembunyikan SUPER_ADMIN dari daftar pengguna sekolah
        const usersRes = await pool.query(`
            SELECT u.id, u.nama, u.username, u.role, u.kelas_id, COALESCE(k.nama_kelas, 'Tanpa Penugasan') AS nama_kelas 
            FROM users u 
            LEFT JOIN kelas k ON u.kelas_id = k.id 
            WHERE u.sekolah_id = $1 AND u.role != 'SUPER_ADMIN'
            ORDER BY u.id ASC
        `, [userSekolahId]);

        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, s.kelas_id, 
                   COALESCE(k.nama_kelas, '-') AS nama_kelas,
                   COALESCE(k.sekolah_id, $1) AS sekolah_id
            FROM siswa s 
            JOIN kelas k ON s.kelas_id = k.id 
            WHERE k.sekolah_id = $1 
            ORDER BY s.id ASC
        `, [userSekolahId]);

        const kelasRes = await pool.query(`SELECT * FROM kelas WHERE sekolah_id = $1 ORDER BY id ASC`, [userSekolahId]);

        const absensiRes = await pool.query(`
            SELECT a.id, a.waktu, s.nama AS nama_siswa, COALESCE(k.nama_kelas, '-') AS nama_kelas 
            FROM absensi a 
            JOIN siswa s ON a.siswa_id = s.id 
            JOIN kelas k ON s.kelas_id = k.id 
            WHERE k.sekolah_id = $1 AND DATE(a.waktu AT TIME ZONE 'Asia/Jakarta') = CURRENT_DATE
            ORDER BY a.waktu DESC
        `, [userSekolahId]);

        const settingsAll = await pool.query("SELECT key, value FROM settings WHERE key = 'pengirim_wa'");
        let pengirimWA = settingsAll.rows.length > 0 ? settingsAll.rows[0].value : 'ADMIN';
        let namaSekolah = await getNamaSekolah(userSekolahId);

        const absensiFormatted = absensiRes.rows.map(row => {
            const dateObj = new Date(row.waktu);
            const waktuWIB = dateObj.toLocaleTimeString('id-ID', { timeZone: 'Asia/Jakarta', hour: '2-digit', minute: '2-digit', second: '2-digit' }).replace(/\./g, ':') + ' WIB';
            return { ...row, waktu_formatted: waktuWIB };
        });

        const usersCleaned = usersRes.rows.map(u => ({ ...u, nama: bersihkanGelar(u.nama) }));
        
        // FORMAT QR UNIK MULTI-TENANT: SCH[sekolah_id]-S[siswa_id]
        const siswaData = await Promise.all(siswaRes.rows.map(async (s) => {
            const qrImage = await generateQRDataURL(`SCH${s.sekolah_id || userSekolahId}-S${s.id}`);
            return { ...s, qrImage };
        }));

        req.session.userId = userId;

        res.render('admin-dashboard', {
            users: usersCleaned,
            siswa: siswaData,
            kelas: kelasRes.rows || [],
            absensiHariIni: absensiFormatted,
            userId: userId,
            statusWA: waStatus[userId] || 'BELUM_TERHUBUNG',
            qrCodeWA: qrCodes[userId] || null,
            pengirimWA: pengirimWA,
            namaSekolah: namaSekolah
        });
    } catch (err) {
        res.status(500).send("Kesalahan Database: " + err.message);
    }
});

app.post('/api/settings/pengirim-wa', async (req, res) => {
    const { pengirim_wa } = req.body;
    try {
        await pool.query(
            `INSERT INTO settings (key, value) VALUES ('pengirim_wa', $1)
             ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value`,
            [pengirim_wa]
        );
        return res.redirect(`/admin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal menyimpan pengaturan: " + err.message);
    }
});

app.post('/api/settings', async (req, res) => {
    const { nama_sekolah } = req.body;
    const userId = req.session.userId || 1;

    try {
        const userRes = await pool.query('SELECT sekolah_id FROM users WHERE id = $1', [userId]);
        const userSekolahId = userRes.rows.length > 0 && userRes.rows[0].sekolah_id ? userRes.rows[0].sekolah_id : (parseInt(process.env.SEKOLAH_ID) || 1);

        await pool.query(`
            INSERT INTO sekolah (id, nama_sekolah) 
            VALUES ($1, $2)
            ON CONFLICT (id) DO UPDATE SET nama_sekolah = EXCLUDED.nama_sekolah;
        `, [userSekolahId, nama_sekolah]);

        res.json({ success: true, message: 'Nama sekolah berhasil diperbarui!' });
    } catch (err) {
        res.status(500).json({ success: false, message: 'Gagal memperbarui pengaturan: ' + err.message });
    }
});

app.get(['/admin/cetak-kartu', '/cetak-kartu'], async (req, res) => {
    const userId = req.session.userId || 1;
    try {
        const userRes = await pool.query('SELECT sekolah_id FROM users WHERE id = $1', [userId]);
        const userSekolahId = userRes.rows.length > 0 && userRes.rows[0].sekolah_id ? userRes.rows[0].sekolah_id : (parseInt(process.env.SEKOLAH_ID) || 1);

        const namaSekolah = await getNamaSekolah(userSekolahId);
        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, s.kelas_id, 
                   COALESCE(k.nama_kelas, '-') AS nama_kelas,
                   COALESCE(k.sekolah_id, $1) AS sekolah_id
            FROM siswa s 
            JOIN kelas k ON s.kelas_id = k.id 
            WHERE k.sekolah_id = $1 
            ORDER BY s.nama ASC
        `, [userSekolahId]);

        const siswaData = await Promise.all(siswaRes.rows.map(async (s) => {
            const qrImage = await generateQRDataURL(`SCH${s.sekolah_id || userSekolahId}-S${s.id}`);
            return { ...s, qrImage };
        }));

        res.render('cetak-kartu', { siswa: siswaData, namaSekolah });
    } catch (err) {
        res.status(500).send("Gagal memuat kartu: " + err.message);
    }
});

app.get(['/wali', '/walikelas-dashboard'], async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId;
    if (!userId) return res.redirect('/');

    try {
        const userRes = await pool.query(`
            SELECT u.id, u.nama, u.role, u.kelas_id, u.sekolah_id,
                   COALESCE(k.nama_kelas, 'Guru Mata Pelajaran (Semua Kelas)') AS nama_kelas
            FROM users u LEFT JOIN kelas k ON u.kelas_id = k.id WHERE u.id = $1
        `, [userId]);

        if (userRes.rows.length === 0) return res.redirect('/');

        const userRaw = userRes.rows[0];
        const userSekolahId = userRaw.sekolah_id || parseInt(process.env.SEKOLAH_ID) || 1;
        userRaw.nama = bersihkanGelar(userRaw.nama);
        const namaSekolah = await getNamaSekolah(userSekolahId);

        let siswaQuery = `
            SELECT s.id, s.nama, s.nomor_wa_ortu, s.kelas_id, 
                   COALESCE(k.nama_kelas, '-') AS nama_kelas,
                   COALESCE(k.sekolah_id, $1) AS sekolah_id 
            FROM siswa s 
            JOIN kelas k ON s.kelas_id = k.id
            WHERE k.sekolah_id = $1
        `;
        const queryParamsSiswa = [userSekolahId];

        if (userRaw.kelas_id) {
            siswaQuery += ` AND s.kelas_id = $2`;
            queryParamsSiswa.push(parseInt(userRaw.kelas_id));
        }

        siswaQuery += ` ORDER BY s.nama ASC`;
        const siswaRes = await pool.query(siswaQuery, queryParamsSiswa);

        let absensiQuery = `
            SELECT a.id, a.waktu, s.nama AS nama_siswa, COALESCE(k.nama_kelas, '-') AS nama_kelas 
            FROM absensi a 
            JOIN siswa s ON a.siswa_id = s.id 
            JOIN kelas k ON s.kelas_id = k.id 
            WHERE k.sekolah_id = $1 AND DATE(a.waktu AT TIME ZONE 'Asia/Jakarta') = CURRENT_DATE
        `;
        const queryParamsAbsensi = [userSekolahId];

        if (userRaw.kelas_id) {
            absensiQuery += ` AND s.kelas_id = $2`;
            queryParamsAbsensi.push(parseInt(userRaw.kelas_id));
        }

        absensiQuery += ` ORDER BY a.waktu DESC`;
        const absensiRes = await pool.query(absensiQuery, queryParamsAbsensi);

        const absensiFormatted = absensiRes.rows.map(row => {
            const dateObj = new Date(row.waktu);
            const waktuWIB = dateObj.toLocaleTimeString('id-ID', { timeZone: 'Asia/Jakarta', hour: '2-digit', minute: '2-digit', second: '2-digit' }).replace(/\./g, ':') + ' WIB';
            return { ...row, waktu_formatted: waktuWIB };
        });

        const siswaData = await Promise.all(siswaRes.rows.map(async (s) => {
            const qrImage = await generateQRDataURL(`SCH${s.sekolah_id || userSekolahId}-S${s.id}`);
            return { ...s, qrImage };
        }));

        req.session.userId = userId;

        res.render('walikelas-dashboard', {
            user: userRaw,
            siswaList: siswaData,
            absensiHariIni: absensiFormatted,
            userId: userId,
            statusWA: waStatus[userId] || 'BELUM_TERHUBUNG',
            qrCodeWA: qrCodes[userId] || null,
            namaSekolah
        });
    } catch (err) {
        res.status(500).send("Kesalahan Database: " + err.message);
    }
});

app.get(['/scan', '/scanner'], async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    try {
        const userRes = await pool.query('SELECT sekolah_id FROM users WHERE id = $1', [userId]);
        const userSekolahId = userRes.rows.length > 0 && userRes.rows[0].sekolah_id ? userRes.rows[0].sekolah_id : (parseInt(process.env.SEKOLAH_ID) || 1);

        const namaSekolah = await getNamaSekolah(userSekolahId);
        res.render('scan', { userId, namaSekolah });
    } catch (err) {
        res.render('scan', { userId, namaSekolah: 'NAMA SEKOLAH BELUM DIATUR' });
    }
});

app.post('/api/kelas/tambah', async (req, res) => {
    const { nama_kelas } = req.body;
    const userId = req.session.userId || 1;

    try {
        if (!nama_kelas) return res.status(400).send("Nama kelas wajib diisi.");

        const userRes = await pool.query('SELECT sekolah_id FROM users WHERE id = $1', [userId]);
        const userSekolahId = userRes.rows.length > 0 && userRes.rows[0].sekolah_id ? userRes.rows[0].sekolah_id : (parseInt(process.env.SEKOLAH_ID) || 1);

        await pool.query('INSERT INTO kelas (nama_kelas, sekolah_id) VALUES ($1, $2)', [nama_kelas.trim(), userSekolahId]);
        return res.redirect(`/admin?userId=${userId}`);
    } catch (err) {
        return res.status(500).send("Gagal menambah kelas: " + err.message);
    }
});

app.post('/api/kelas/hapus/:id', async (req, res) => {
    const kelasId = parseInt(req.params.id);
    try {
        await pool.query('UPDATE users SET kelas_id = NULL WHERE kelas_id = $1', [kelasId]);
        await pool.query('UPDATE siswa SET kelas_id = NULL WHERE kelas_id = $1', [kelasId]);
        await pool.query('DELETE FROM kelas WHERE id = $1', [kelasId]);
        return res.redirect(`/admin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal menghapus rombel: " + err.message);
    }
});

app.post('/api/kelas/hapus-semua', async (req, res) => {
    const userId = req.session.userId || 1;
    try {
        const userRes = await pool.query('SELECT sekolah_id FROM users WHERE id = $1', [userId]);
        const userSekolahId = userRes.rows.length > 0 && userRes.rows[0].sekolah_id ? userRes.rows[0].sekolah_id : (parseInt(process.env.SEKOLAH_ID) || 1);

        await pool.query('UPDATE users SET kelas_id = NULL WHERE sekolah_id = $1', [userSekolahId]);
        await pool.query('DELETE FROM siswa WHERE kelas_id IN (SELECT id FROM kelas WHERE sekolah_id = $1)', [userSekolahId]);
        await pool.query('DELETE FROM kelas WHERE sekolah_id = $1', [userSekolahId]);

        return res.redirect(`/admin?userId=${userId}`);
    } catch (err) {
        return res.status(500).send("Gagal menghapus semua rombel: " + err.message);
    }
});

app.post('/api/guru/tambah', async (req, res) => {
    const { nama, username, password, kelas_id } = req.body;
    const userId = req.session.userId || 1;

    try {
        if (!nama || !username || !password) return res.status(400).send("Semua kolom wajib diisi.");

        const userRes = await pool.query('SELECT sekolah_id FROM users WHERE id = $1', [userId]);
        const userSekolahId = userRes.rows.length > 0 && userRes.rows[0].sekolah_id ? userRes.rows[0].sekolah_id : (parseInt(process.env.SEKOLAH_ID) || 1);

        const parsedKelasId = kelas_id ? parseInt(kelas_id) : null;
        await pool.query(
            'INSERT INTO users (nama, username, password, role, kelas_id, sekolah_id) VALUES ($1, $2, $3, $4, $5, $6)',
            [nama.trim(), username.trim(), password.trim(), 'WALI_KELAS', parsedKelasId, userSekolahId]
        );
        return res.redirect(`/admin?userId=${userId}`);
    } catch (err) {
        return res.status(500).send("Gagal menambah guru: " + err.message);
    }
});

app.post('/api/guru/edit/:id', async (req, res) => {
    const guruId = parseInt(req.params.id);
    const { nama, username, password, kelas_id } = req.body;

    try {
        if (!nama || !username) return res.status(400).send("Nama dan Username wajib diisi.");
        const parsedKelasId = kelas_id ? parseInt(kelas_id) : null;

        if (password && password.trim() !== '') {
            await pool.query(
                `UPDATE users SET nama = $1, username = $2, password = $3, kelas_id = $4 WHERE id = $5`,
                [nama.trim(), username.trim(), password.trim(), parsedKelasId, guruId]
            );
        } else {
            await pool.query(
                `UPDATE users SET nama = $1, username = $2, kelas_id = $3 WHERE id = $4`,
                [nama.trim(), username.trim(), parsedKelasId, guruId]
            );
        }
        return res.redirect(`/admin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal memperbarui data guru: " + err.message);
    }
});

app.post('/api/guru/hapus/:id', async (req, res) => {
    try {
        await pool.query('DELETE FROM users WHERE id = $1 AND role NOT IN (\'ADMIN\', \'SUPER_ADMIN\')', [parseInt(req.params.id)]);
        return res.redirect(`/admin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal menghapus guru: " + err.message);
    }
});

app.post('/api/siswa/tambah', async (req, res) => {
    const { nama, nomor_wa_ortu, kelas_id } = req.body;
    try {
        if (!nama) return res.status(400).send("Nama siswa wajib diisi.");
        const parsedKelasId = kelas_id ? parseInt(kelas_id) : null;
        await pool.query(
            'INSERT INTO siswa (nama, nomor_wa_ortu, kelas_id) VALUES ($1, $2, $3)',
            [nama.trim(), nomor_wa_ortu ? nomor_wa_ortu.trim() : '', parsedKelasId]
        );
        return res.redirect(`/admin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal menambah siswa: " + err.message);
    }
});

app.post('/api/siswa/edit/:id', async (req, res) => {
    const siswaId = parseInt(req.params.id);
    const { nama, nomor_wa_ortu, kelas_id } = req.body;

    try {
        if (!nama) return res.status(400).send("Nama siswa wajib diisi.");
        const parsedKelasId = kelas_id ? parseInt(kelas_id) : null;

        await pool.query(
            `UPDATE siswa SET nama = $1, nomor_wa_ortu = $2, kelas_id = $3 WHERE id = $4`,
            [nama.trim(), nomor_wa_ortu ? nomor_wa_ortu.trim() : '', parsedKelasId, siswaId]
        );

        return res.redirect(`/admin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal memperbarui data siswa: " + err.message);
    }
});

app.post('/api/siswa/hapus/:id', async (req, res) => {
    try {
        await pool.query('DELETE FROM siswa WHERE id = $1', [parseInt(req.params.id)]);
        return res.redirect(`/admin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal menghapus siswa: " + err.message);
    }
});

app.post('/api/siswa/hapus-semua', async (req, res) => {
    const userId = req.session.userId || 1;
    try {
        const userRes = await pool.query('SELECT sekolah_id FROM users WHERE id = $1', [userId]);
        const userSekolahId = userRes.rows.length > 0 && userRes.rows[0].sekolah_id ? userRes.rows[0].sekolah_id : (parseInt(process.env.SEKOLAH_ID) || 1);

        await pool.query('DELETE FROM absensi WHERE siswa_id IN (SELECT s.id FROM siswa s JOIN kelas k ON s.kelas_id = k.id WHERE k.sekolah_id = $1)', [userSekolahId]);
        await pool.query('DELETE FROM siswa WHERE kelas_id IN (SELECT id FROM kelas WHERE sekolah_id = $1)', [userSekolahId]);

        return res.redirect(`/admin?userId=${userId}`);
    } catch (err) {
        return res.status(500).send("Gagal melakukan reset data siswa: " + err.message);
    }
});

app.post('/api/siswa/import-excel', upload.single('file_excel'), async (req, res) => {
    const client = await pool.connect();
    const userId = req.session.userId || 1;

    try {
        if (!req.file) return res.status(400).json({ success: false, message: "Berkas Excel/CSV wajib diunggah!" });

        const userRes = await client.query('SELECT sekolah_id FROM users WHERE id = $1', [userId]);
        const userSekolahId = userRes.rows.length > 0 && userRes.rows[0].sekolah_id ? userRes.rows[0].sekolah_id : (parseInt(process.env.SEKOLAH_ID) || 1);

        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });

        if (sheetData.length === 0) return res.status(400).json({ success: false, message: "Berkas Excel kosong atau tidak terbaca." });

        const rowsToInsert = [];

        for (const row of sheetData) {
            let namaSiswa = "";
            let nomorWa = "";
            let namaKelas = null;

            Object.keys(row).forEach(key => {
                const cleanKey = key.toString().toLowerCase().trim();
                const val = row[key] ? row[key].toString().trim() : "";

                if (cleanKey.includes('nama') || cleanKey.includes('peserta didik') || cleanKey.includes('siswa')) if (!namaSiswa && val) namaSiswa = val;
                if (cleanKey.includes('hp') || cleanKey.includes('wa') || cleanKey.includes('telepon') || cleanKey.includes('seluler') || cleanKey.includes('kontak')) if (!nomorWa && val) nomorWa = val;
                if (cleanKey.includes('rombel') || cleanKey.includes('kelas') || cleanKey.includes('rombongan') || cleanKey.includes('tingkat')) if (!namaKelas && val) namaKelas = val;
            });

            if (namaSiswa && namaSiswa.length > 1) {
                let kelasId = null;

                if (namaKelas) {
                    const cleanNamaKelas = namaKelas.toString().replace(/\s+/g, ' ').trim();
                    let kRes = await client.query('SELECT id FROM kelas WHERE LOWER(TRIM(nama_kelas)) = LOWER($1) AND sekolah_id = $2', [cleanNamaKelas, userSekolahId]);
                    
                    if (kRes.rows.length > 0) {
                        kelasId = kRes.rows[0].id;
                    } else {
                        let newKRes = await client.query('INSERT INTO kelas (nama_kelas, sekolah_id) VALUES ($1, $2) RETURNING id', [cleanNamaKelas, userSekolahId]);
                        kelasId = newKRes.rows[0].id;
                    }
                }

                rowsToInsert.push({ nama: namaSiswa, wa: nomorWa, kelasId: kelasId });
            }
        }

        if (rowsToInsert.length === 0) return res.json({ success: false, message: "Gagal membaca data Excel." });

        await client.query('BEGIN');
        const valueStrings = [];
        const valueParams = [];
        let paramIndex = 1;

        rowsToInsert.forEach(item => {
            valueStrings.push(`($${paramIndex}, $${paramIndex + 1}, $${paramIndex + 2})`);
            valueParams.push(item.nama, item.wa, item.kelasId);
            paramIndex += 3;
        });

        await client.query(`INSERT INTO siswa (nama, nomor_wa_ortu, kelas_id) VALUES ${valueStrings.join(', ')}`, valueParams);
        await client.query('COMMIT');

        return res.json({ success: true, message: `Berhasil mengimpor ${rowsToInsert.length} data siswa!` });
    } catch (err) {
        await client.query('ROLLBACK');
        return res.status(500).json({ success: false, message: "Gagal memproses berkas: " + err.message });
    } finally {
        client.release();
    }
});

// ----------------- ENDPOINT WA ----------------- //

app.get('/api/start-wa', async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session?.userId || 1;
    delete pairingCodes[userId];
    connectToWhatsApp(userId);
    res.json({ success: true, message: 'Inisialisasi WhatsApp dimulai...' });
});

app.get('/api/request-pairing', async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session?.userId || 1;
    const phone = req.query.phone;
    if (!phone) return res.status(400).json({ success: false, message: 'Nomor WhatsApp wajib diisi!' });

    delete qrCodes[userId];
    delete pairingCodes[userId];
    waStatus[userId] = 'MENUNGGU_PAIRING_CODE';

    connectToWhatsApp(userId, phone);
    res.json({ success: true, message: 'Mempersiapkan kode tautan...' });
});

app.get('/api/wa-status', (req, res) => {
    const userId = parseInt(req.query.userId) || req.session?.userId || 1;
    res.json({
        success: true,
        statusWA: waStatus[userId] || 'BELUM_TERHUBUNG',
        qrCodeWA: pairingCodes[userId] ? null : (qrCodes[userId] || null),
        pairingCode: pairingCodes[userId] || null
    });
});

app.get('/api/reset-wa', async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session?.userId || 1;
    if (reconnectTimers[userId]) { clearTimeout(reconnectTimers[userId]); delete reconnectTimers[userId]; }
    if (waSessions[userId]) { try { waSessions[userId].end(undefined); } catch (e) {} delete waSessions[userId]; }
    
    delete qrCodes[userId];
    delete pairingCodes[userId];
    waStatus[userId] = 'BELUM_TERHUBUNG';

    const authFolder = path.join(__dirname, 'auth_sessions', `user_${userId}`);
    if (fs.existsSync(authFolder)) fs.rmSync(authFolder, { recursive: true, force: true });
    
    res.json({ success: true, message: 'Sesi WA Berhasil Direset!' });
});

// ----------------- PROSES SCAN MULTI-SEKOLAH PINTAR ----------------- //
app.post('/api/scan', async (req, res) => {
    const { siswa_id, scanned_by, tipe = 'MASUK' } = req.body;
    if (!siswa_id) return res.status(400).json({ success: false, message: "Kode QR tidak terdeteksi." });

    try {
        let parsedSiswaId;
        const rawStr = siswa_id.toString().trim();
        
        if (rawStr.includes('S')) {
            parsedSiswaId = parseInt(rawStr.split('S').pop());
        } else {
            parsedSiswaId = parseInt(rawStr.replace(/[^0-9]/g, ''));
        }

        if (isNaN(parsedSiswaId)) return res.status(400).json({ success: false, message: "Format Kode QR Siswa Tidak Valid." });

        const parsedScannedBy = parseInt(scanned_by) || 1;
        const tipeAbsen = tipe.toUpperCase() === 'PULANG' ? 'PULANG' : 'MASUK';

        // TARIK NAMA SEKOLAH SECARA DINAMIS SESUAI SEKOLAH SISWA TERDAPAT
        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, s.kelas_id, 
                   COALESCE(k.nama_kelas, '-') AS nama_kelas,
                   COALESCE(sch.nama_sekolah, set_sch.value, 'SEKOLAH') AS nama_sekolah_siswa,
                   u.id AS wali_kelas_user_id
            FROM siswa s 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            LEFT JOIN sekolah sch ON k.sekolah_id = sch.id
            LEFT JOIN settings set_sch ON set_sch.key = 'nama_sekolah'
            LEFT JOIN users u ON u.kelas_id = s.kelas_id AND u.role = 'WALI_KELAS'
            WHERE s.id = $1
        `, [parsedSiswaId]);

        if (siswaRes.rows.length === 0) return res.status(404).json({ success: false, message: `ID Siswa #${parsedSiswaId} Tidak Terdaftar!` });

        const siswa = siswaRes.rows[0];
        const namaSekolahResmi = siswa.nama_sekolah_siswa;

        const settingWaRes = await pool.query("SELECT value FROM settings WHERE key = 'pengirim_wa'");
        const modePengirim = settingWaRes.rows.length > 0 ? settingWaRes.rows[0].value : 'ADMIN';

        const now = new Date();
        const jamWib = now.toLocaleTimeString('id-ID', { timeZone: 'Asia/Jakarta', hour: '2-digit', minute: '2-digit', second: '2-digit' }).replace(/\./g, ':') + ' WIB';
        const tglWib = now.toLocaleDateString('id-ID', { timeZone: 'Asia/Jakarta', weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

        await pool.query(
            `INSERT INTO absensi (siswa_id, status, scanned_by, tipe, waktu) 
             VALUES ($1, 'HADIR', $2, $3, NOW() AT TIME ZONE 'Asia/Jakarta')`,
            [siswa.id, parsedScannedBy, tipeAbsen]
        );

        let waClient = null;
        if (modePengirim === 'WALI_KELAS' && siswa.wali_kelas_user_id) waClient = waSessions[siswa.wali_kelas_user_id];
        if (!waClient) waClient = waSessions[1] || waSessions[parsedScannedBy];
        if (!waClient) {
            const availableKeys = Object.keys(waSessions);
            if (availableKeys.length > 0) waClient = waSessions[availableKeys[0]];
        }

        let statusWA = "Notifikasi WhatsApp Tidak Terkirim (Layanan WA Belum Terkoneksi)";

        if (waClient && siswa.nomor_wa_ortu) {
            let phone = siswa.nomor_wa_ortu.toString().trim().replace(/[^0-9]/g, '');
            if (phone.startsWith('0')) phone = '62' + phone.slice(1);
            const formattedJid = phone + '@s.whatsapp.net';

            let pesan = '';
            if (tipeAbsen === 'MASUK') {
                pesan = `*${namaSekolahResmi.toUpperCase()}*\n` +
                        `*PEMBERITAHUAN PRESENSI KEHADIRAN SISWA*\n` +
                        `_________________________________________\n\n` +
                        `Yth. Bapak/Ibu Orang Tua / Wali Murid,\n\n` +
                        `Diberitahukan bahwa putra/putri Anda telah tiba di sekolah dan melakukan presensi masuk:\n\n` +
                        `• Nama Siswa : *${siswa.nama}*\n` +
                        `• Kelas / Rombel : *${siswa.nama_kelas}*\n` +
                        `• Waktu Masuk : *${jamWib}*\n` +
                        `• Tanggal : *${tglWib}*\n` +
                        `• Status : *HADIR (Scan Masuk) ✅*\n\n` +
                        `Terima kasih atas perhatian dan kerja samanya.\n\n` +
                        `_Pesan otomatis dikirim via Sistem Presensi SD._`;
            } else {
                pesan = `*${namaSekolahResmi.toUpperCase()}*\n` +
                        `*PEMBERITAHUAN PRESENSI KEPULANGAN SISWA*\n` +
                        `_________________________________________\n\n` +
                        `Yth. Bapak/Ibu Orang Tua / Wali Murid,\n\n` +
                        `Diberitahukan bahwa putra/putri Anda telah selesai mengikuti KBM dan melakukan presensi pulang:\n\n` +
                        `• Nama Siswa : *${siswa.nama}*\n` +
                        `• Kelas / Rombel : *${siswa.nama_kelas}*\n` +
                        `• Waktu Pulang : *${jamWib}*\n` +
                        `• Tanggal : *${tglWib}*\n` +
                        `• Status : *PULANG (Scan Kepulangan) 🏠*\n\n` +
                        `Hati-hati di jalan dan terima kasih.\n\n` +
                        `_Pesan otomatis dikirim via Sistem Presensi SD._`;
            }

            waClient.sendMessage(formattedJid, { text: pesan }).catch(e => console.error("Gagal Mengirim WA:", e.message));
            statusWA = `Notifikasi WA (${tipeAbsen}) Berhasil Dikirimkan ke Wali Murid ✅`;
        }

        return res.json({
            success: true,
            message: statusWA,
            siswa: { id: siswa.id, nama: siswa.nama, nama_kelas: siswa.nama_kelas, waktu: jamWib, tipe: tipeAbsen }
        });

    } catch (err) {
        return res.status(500).json({ success: false, message: "Kendala Sistem: " + err.message });
    }
});

app.post('/api/absensi/reset-riwayat', async (req, res) => {
    try {
        await pool.query('DELETE FROM absensi');
        await pool.query('ALTER SEQUENCE absensi_id_seq RESTART WITH 1');
        return res.redirect(req.headers.referer || `/admin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal membersihkan riwayat absensi: " + err.message);
    }
});

app.get('/api/absensi/preview', async (req, res) => {
    const { bulan, tahun, kelas_id } = req.query;
    if (!bulan || !tahun) return res.status(400).json({ success: false, message: "Bulan dan Tahun wajib diisi." });

    try {
        let query = `
            SELECT s.nama AS nama_siswa, COALESCE(k.nama_kelas, '-') AS nama_kelas, COUNT(a.id) AS total_hadir
            FROM siswa s
            LEFT JOIN kelas k ON s.kelas_id = k.id
            LEFT JOIN absensi a ON s.id = a.siswa_id 
                AND EXTRACT(MONTH FROM a.waktu AT TIME ZONE 'Asia/Jakarta') = $1
                AND EXTRACT(YEAR FROM a.waktu AT TIME ZONE 'Asia/Jakarta') = $2
        `;
        const queryParams = [parseInt(bulan), parseInt(tahun)];

        if (kelas_id && kelas_id !== 'all' && kelas_id !== 'null' && kelas_id !== '') {
            query += ` WHERE s.kelas_id = $3`;
            queryParams.push(parseInt(kelas_id));
        }

        query += ` GROUP BY s.id, s.nama, k.nama_kelas ORDER BY s.nama ASC`;

        const result = await pool.query(query, queryParams);
        return res.json({ success: true, data: result.rows });
    } catch (err) {
        return res.status(500).json({ success: false, message: err.message });
    }
});

app.get('/api/absensi/export', async (req, res) => {
    const { bulan, tahun, kelas_id, format } = req.query;
    const userId = req.session.userId || 1;

    if (!bulan || !tahun) return res.status(400).send("Bulan dan Tahun wajib diisi.");

    try {
        const userRes = await pool.query('SELECT sekolah_id FROM users WHERE id = $1', [userId]);
        const userSekolahId = userRes.rows.length > 0 && userRes.rows[0].sekolah_id ? userRes.rows[0].sekolah_id : (parseInt(process.env.SEKOLAH_ID) || 1);

        let query = `
            SELECT s.nama AS nama_siswa, COALESCE(k.nama_kelas, '-') AS nama_kelas, COUNT(a.id) AS total_hadir
            FROM siswa s
            JOIN kelas k ON s.kelas_id = k.id
            LEFT JOIN absensi a ON s.id = a.siswa_id 
                AND EXTRACT(MONTH FROM a.waktu AT TIME ZONE 'Asia/Jakarta') = $1
                AND EXTRACT(YEAR FROM a.waktu AT TIME ZONE 'Asia/Jakarta') = $2
            WHERE k.sekolah_id = $3
        `;
        const queryParams = [parseInt(bulan), parseInt(tahun), userSekolahId];

        if (kelas_id && kelas_id !== 'all' && kelas_id !== 'null' && kelas_id !== '') {
            query += ` AND s.kelas_id = $4`;
            queryParams.push(parseInt(kelas_id));
        }

        query += ` GROUP BY s.id, s.nama, k.nama_kelas ORDER BY s.nama ASC`;

        const result = await pool.query(query, queryParams);
        const dataRekap = result.rows;
        const namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"][parseInt(bulan) - 1];
        const judul = `REKAP PRESENSI SISWA - ${namaBulan.toUpperCase()} ${tahun}`;
        const namaSekolahHeader = await getNamaSekolah(userSekolahId);

        if (format === 'excel') {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Rekap Absensi');

            worksheet.mergeCells('A1:D1');
            worksheet.getCell('A1').value = namaSekolahHeader;
            worksheet.getCell('A1').font = { bold: true, size: 14 };
            worksheet.getCell('A1').alignment = { horizontal: 'center' };

            worksheet.mergeCells('A2:D2');
            worksheet.getCell('A2').value = judul;
            worksheet.getCell('A2').font = { bold: true, size: 12 };
            worksheet.getCell('A2').alignment = { horizontal: 'center' };

            worksheet.addRow([]);
            const headerRow = worksheet.addRow(['No', 'Nama Siswa', 'Kelas / Rombel', 'Total Kehadiran']);
            headerRow.font = { bold: true };
            headerRow.eachCell((cell) => {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9EAD3' } };
                cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            });

            dataRekap.forEach((row, idx) => {
                const r = worksheet.addRow([idx + 1, row.nama_siswa, row.nama_kelas, `${row.total_hadir} Hari`]);
                r.eachCell(c => c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } });
            });

            worksheet.columns = [{ width: 6 }, { width: 30 }, { width: 20 }, { width: 18 }];
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', `attachment; filename=Rekap_Presensi_${bulan}_${tahun}.xlsx`);
            return workbook.xlsx.write(res).then(() => res.end());
        }

        if (format === 'word') {
            const tableRows = [
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ text: "No", bold: true })], width: { size: 10, type: WidthType.PERCENTAGE } }),
                        new TableCell({ children: [new Paragraph({ text: "Nama Siswa", bold: true })], width: { size: 45, type: WidthType.PERCENTAGE } }),
                        new TableCell({ children: [new Paragraph({ text: "Kelas", bold: true })], width: { size: 25, type: WidthType.PERCENTAGE } }),
                        new TableCell({ children: [new Paragraph({ text: "Total Hadir", bold: true })], width: { size: 20, type: WidthType.PERCENTAGE } }),
                    ]
                }),
                ...dataRekap.map((row, idx) => new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph((idx + 1).toString())] }),
                        new TableCell({ children: [new Paragraph(row.nama_siswa)] }),
                        new TableCell({ children: [new Paragraph(row.nama_kelas)] }),
                        new TableCell({ children: [new Paragraph(`${row.total_hadir} Hari`)] }),
                    ]
                }))
            ];

            const doc = new Document({
                sections: [{
                    children: [
                        new Paragraph({ text: namaSekolahHeader, heading: "Heading1", alignment: AlignmentType.CENTER }),
                        new Paragraph({ text: judul, heading: "Heading2", alignment: AlignmentType.CENTER }),
                        new Paragraph({ text: "" }),
                        new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } })
                    ]
                }]
            });

            const buffer = await Packer.toBuffer(doc);
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.setHeader('Content-Disposition', `attachment; filename=Rekap_Presensi_${bulan}_${tahun}.docx`);
            return res.send(buffer);
        }

        if (format === 'pdf') {
            const doc = new PDFDocument({ margin: 40, size: 'A4' });
            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', `attachment; filename=Rekap_Presensi_${bulan}_${tahun}.pdf`);

            doc.pipe(res);
            doc.fontSize(14).font('Helvetica-Bold').text(namaSekolahHeader, { align: 'center' });
            doc.fontSize(11).font('Helvetica').text(judul, { align: 'center' });
            doc.moveDown(1.5);

            let y = doc.y;
            const startX = 40;
            const colWidths = [40, 230, 130, 100];

            doc.font('Helvetica-Bold').fontSize(10);
            doc.text('No', startX, y);
            doc.text('Nama Siswa', startX + colWidths[0], y);
            doc.text('Kelas', startX + colWidths[0] + colWidths[1], y);
            doc.text('Total Hadir', startX + colWidths[0] + colWidths[1] + colWidths[2], y);

            doc.moveTo(startX, y + 15).lineTo(startX + 500, y + 15).stroke();
            y += 22;

            doc.font('Helvetica').fontSize(9);
            dataRekap.forEach((row, i) => {
                if (y > 750) {
                    doc.addPage();
                    y = 40;
                }
                doc.text((i + 1).toString(), startX, y);
                doc.text(row.nama_siswa, startX + colWidths[0], y);
                doc.text(row.nama_kelas, startX + colWidths[0] + colWidths[1], y);
                doc.text(`${row.total_hadir} Hari`, startX + colWidths[0] + colWidths[1] + colWidths[2], y);
                y += 18;
            });

            doc.end();
            return;
        }

        return res.status(400).send("Format ekspor tidak valid.");

    } catch (err) {
        res.status(500).send("Gagal mengekspor data: " + err.message);
    }
});

// ----------------- CRON JOB NOTIFIKASI ALPA MULTI-SEKOLAH ----------------- //
cron.schedule('0 9 * * 1-6', async () => {
    console.log('⏰ [CRON JOB] Memulai pengecekan siswa yang belum presensi masuk jam 09:00 WIB...');

    try {
        const querySiswaAbsen = `
            SELECT s.id, s.nama, s.nomor_wa_ortu, 
                   COALESCE(k.nama_kelas, '-') AS nama_kelas, 
                   COALESCE(sch.nama_sekolah, set_sch.value, 'SEKOLAH') AS nama_sekolah_siswa,
                   u.id AS wali_kelas_user_id
            FROM siswa s
            LEFT JOIN kelas k ON s.kelas_id = k.id
            LEFT JOIN sekolah sch ON k.sekolah_id = sch.id
            LEFT JOIN settings set_sch ON set_sch.key = 'nama_sekolah'
            LEFT JOIN users u ON u.kelas_id = s.kelas_id AND u.role = 'WALI_KELAS'
            WHERE s.id NOT IN (
                SELECT DISTINCT siswa_id 
                FROM absensi 
                WHERE DATE(waktu AT TIME ZONE 'Asia/Jakarta') = CURRENT_DATE 
                  AND tipe = 'MASUK'
            )
        `;

        const result = await pool.query(querySiswaAbsen);
        const siswaBelumPresensi = result.rows;

        if (siswaBelumPresensi.length === 0) return;

        const settingWaRes = await pool.query("SELECT value FROM settings WHERE key = 'pengirim_wa'");
        const modePengirim = settingWaRes.rows.length > 0 ? settingWaRes.rows[0].value : 'ADMIN';

        const now = new Date();
        const tglWib = now.toLocaleDateString('id-ID', { timeZone: 'Asia/Jakarta', weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

        for (const siswa of siswaBelumPresensi) {
            if (!siswa.nomor_wa_ortu) continue;

            let waClient = null;
            if (modePengirim === 'WALI_KELAS' && siswa.wali_kelas_user_id) waClient = waSessions[siswa.wali_kelas_user_id];
            if (!waClient) waClient = waSessions[1];
            if (!waClient) {
                const availableKeys = Object.keys(waSessions);
                if (availableKeys.length > 0) waClient = waSessions[availableKeys[0]];
            }

            if (waClient) {
                let phone = siswa.nomor_wa_ortu.toString().trim().replace(/[^0-9]/g, '');
                if (phone.startsWith('0')) phone = '62' + phone.slice(1);
                const formattedJid = phone + '@s.whatsapp.net';

                const pesan = `*${siswa.nama_sekolah_siswa.toUpperCase()}*\n` +
                              `*PEMBERITAHUAN KETIDAKHADIRAN SISWA*\n` +
                              `_________________________________________\n\n` +
                              `Yth. Bapak/Ibu Orang Tua / Wali Murid,\n\n` +
                              `Diberitahukan bahwa hingga pukul *09:00 WIB*, putra/putri Anda belum melakukan presensi kehadiran di sekolah:\n\n` +
                              `• Nama Siswa : *${siswa.nama}*\n` +
                              `• Kelas / Rombel : *${siswa.nama_kelas}*\n` +
                              `• Tanggal : *${tglWib}*\n` +
                              `• Status : *BELUM PRESENSI / ALPA* ⚠️\n\n` +
                              `Apabila putra/putri Anda berhalangan hadir karena sakit atau izin, mohon konfirmasinya kepada Wali Kelas.\n\n` +
                              `Terima kasih atas perhatiannya.\n\n` +
                              `_Pesan otomatis dikirim via Sistem Presensi SD._`;

                try {
                    await waClient.sendMessage(formattedJid, { text: pesan });
                } catch (sendErr) {}

                await new Promise(resolve => setTimeout(resolve, 3000));
            }
        }
    } catch (err) {}
});

app.get('/ping', (req, res) => res.send('OK'));

app.listen(PORT, '0.0.0.0', () => console.log(`🚀 Server Presensi Multi-Tenant Aktif di Port ${PORT}`));
