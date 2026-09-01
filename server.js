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

const multer = require('multer');
const XLSX = require('xlsx');
const upload = multer({ storage: multer.memoryStorage() });

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

const waSessions = {};
const qrCodes = {};
const waStatus = {};
const pairingCodes = {};

function bersihkanGelar(nama) {
    if (!nama) return '';
    return nama.replace(/,?\s*(S\.Pd|M\.Pd|S\.Ag|S\.T|S\.Kom|M\.Si|S\.Sos|S\.SE|M\.M|A\.Ma|Sd)\.?/gi, '').trim();
}

async function generateQRDataURL(text) {
    try {
        if (!text) return '';
        return await QRCode.toDataURL(text.toString());
    } catch (err) {
        return '';
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
        if (waSessions[userId]) {
            try { waSessions[userId].end(undefined); } catch (e) {}
            delete waSessions[userId];
        }

        waStatus[userId] = 'PROSES_INIT';
        delete qrCodes[userId];
        delete pairingCodes[userId];

        const { state, saveCreds } = await getAuthState(userId);
        const { version } = await fetchLatestBaileysVersion();

        const sock = makeWASocket({
            logger: pino({ level: 'silent' }),
            auth: state,
            printQRInTerminal: false,
            browser: ["Ubuntu", "Chrome", "120.0.0.0"],
            connectTimeoutMs: 60000,
            defaultQueryTimeoutMs: 60000,
            keepAliveIntervalMs: 10000,
            qrTimeout: 45000,
            syncFullHistory: false
        });

        waSessions[userId] = sock;
        sock.ev.on('creds.update', saveCreds);

        if (phoneNumber && !sock.authState.creds.registered) {
            setTimeout(async () => {
                try {
                    let cleanPhone = phoneNumber.replace(/[^0-9]/g, '');
                    if (cleanPhone.startsWith('0')) cleanPhone = '62' + cleanPhone.slice(1);
                    
                    console.log(`📱 Meringkas request pairing untuk nomor: ${cleanPhone}`);
                    const code = await sock.requestPairingCode(cleanPhone);
                    pairingCodes[userId] = code;
                    waStatus[userId] = 'MENUNGGU_PAIRING_CODE';
                    console.log(`🔑 [User #${userId}] Kode Pairing WA Berhasil: ${code}`);
                } catch (pErr) {
                    console.error("❌ Gagal Request Pairing Code:", pErr.message);
                    waStatus[userId] = 'ERROR_PAIRING';
                }
            }, 3000);
        }

        sock.ev.on('connection.update', async (update) => {
            const { connection, lastDisconnect, qr } = update;

            if (qr && !phoneNumber) {
                try {
                    qrCodes[userId] = await generateQRDataURL(qr);
                    waStatus[userId] = 'MENUNGGU_SCAN';
                } catch (qrErr) {
                    console.error("Gagal generate QR Code WA:", qrErr);
                }
            }

            if (connection === 'open') {
                waStatus[userId] = 'TERHUBUNG';
                delete qrCodes[userId];
                delete pairingCodes[userId];
                console.log(`✅ [User #${userId}] WhatsApp Berhasil Terhubung!`);
            }

            if (connection === 'close') {
                const statusCode = lastDisconnect?.error?.output?.statusCode;
                const shouldReconnect = (statusCode !== DisconnectReason.loggedOut);

                waStatus[userId] = 'TERPUTUS';
                delete waSessions[userId];

                if (shouldReconnect) {
                    setTimeout(() => connectToWhatsApp(userId), 3000);
                } else {
                    delete qrCodes[userId];
                    delete pairingCodes[userId];
                    const authFolder = path.join(__dirname, 'auth_sessions', `user_${userId}`);
                    if (fs.existsSync(authFolder)) {
                        fs.rmSync(authFolder, { recursive: true, force: true });
                    }
                }
            }
        });
    } catch (err) {
        console.error(`❌ WA Connect Error User #${userId}:`, err.message);
        waStatus[userId] = 'ERROR';
    }
}

// ---------------- ROUTES ---------------- //

app.get('/', (req, res) => res.render('login', { error: null }));

app.post('/login', async (req, res) => {
    const { username, password } = req.body;
    try {
        if (!username || !password) return res.render('login', { error: 'Username dan kata sandi wajib diisi.' });

        const result = await pool.query(
            'SELECT * FROM users WHERE LOWER(username) = LOWER($1) AND password = $2',
            [username.trim(), password.trim()]
        );

        if (result.rows.length === 0) return res.render('login', { error: 'Username atau kata sandi tidak valid.' });

        const user = result.rows[0];
        req.session.userId = user.id;

        if (user.role === 'ADMIN') {
            return res.redirect(`/admin?userId=${user.id}`);
        } else {
            return res.redirect(`/wali?userId=${user.id}`);
        }
    } catch (err) {
        return res.render('login', { error: 'Kesalahan Sistem Database: ' + err.message });
    }
});

app.get('/admin', async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    try {
        const usersRes = await pool.query(`
            SELECT u.id, u.nama, u.username, u.role, u.kelas_id, COALESCE(k.nama_kelas, 'Tanpa Penugasan') AS nama_kelas 
            FROM users u LEFT JOIN kelas k ON u.kelas_id = k.id ORDER BY u.id ASC
        `);
        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, s.kelas_id, COALESCE(k.nama_kelas, '-') AS nama_kelas 
            FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id ORDER BY s.id ASC
        `);
        const kelasRes = await pool.query(`SELECT * FROM kelas ORDER BY id ASC`);

        const usersCleaned = usersRes.rows.map(u => ({ ...u, nama: bersihkanGelar(u.nama) }));
        const siswaData = await Promise.all(siswaRes.rows.map(async (s) => {
            const qrImage = await generateQRDataURL(s.id.toString());
            return { ...s, qrImage };
        }));

        res.render('admin-dashboard', {
            users: usersCleaned,
            siswa: siswaData,
            kelas: kelasRes.rows || [],
            userId: userId
        });
    } catch (err) {
        res.status(500).send("Kesalahan Database: " + err.message);
    }
});

app.get(['/wali', '/walikelas-dashboard'], async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    try {
        const userRes = await pool.query(`SELECT id, nama, role, kelas_id FROM users WHERE id = $1`, [userId]);
        const userRaw = userRes.rows[0] || { id: userId, nama: 'Tenaga Pendidik', role: 'WALI_KELAS', kelas_id: null };
        userRaw.nama = bersihkanGelar(userRaw.nama);

        let namaKelas = 'Seluruh Rombongan Belajar';
        if (userRaw.kelas_id) {
            const kRes = await pool.query(`SELECT nama_kelas FROM kelas WHERE id = $1`, [userRaw.kelas_id]);
            if (kRes.rows.length > 0) namaKelas = kRes.rows[0].nama_kelas;
        }
        userRaw.nama_kelas = namaKelas;

        let siswaQuery = `SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, '-') AS nama_kelas FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id`;
        const queryParamsSiswa = [];
        if (userRaw.kelas_id) {
            siswaQuery += ` WHERE s.kelas_id = $1`;
            queryParamsSiswa.push(parseInt(userRaw.kelas_id));
        }
        siswaQuery += ` ORDER BY s.nama ASC`;
        const siswaRes = await pool.query(siswaQuery, queryParamsSiswa);

        const siswaData = await Promise.all(siswaRes.rows.map(async (s) => {
            const qrImage = await generateQRDataURL(s.id.toString());
            return { ...s, qrImage };
        }));

        res.render('walikelas-dashboard', {
            user: userRaw,
            siswaList: siswaData,
            userId: userId,
            statusWA: waStatus[userId] || 'BELUM_TERHUBUNG',
            qrCodeWA: qrCodes[userId] || null
        });
    } catch (err) {
        res.status(500).send("Kesalahan Database: " + err.message);
    }
});

app.get(['/scan', '/scanner'], (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    res.render('scan', { userId: userId });
});

// ---------------- API WA GATEWAY ---------------- //

app.get('/api/start-wa', async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    connectToWhatsApp(userId);
    res.json({ success: true, message: 'Inisialisasi WhatsApp dimulai...' });
});

app.get('/api/request-pairing', async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    const phone = req.query.phone;

    if (!phone) {
        return res.status(400).json({ success: false, message: 'Nomor WhatsApp wajib diisi!' });
    }

    connectToWhatsApp(userId, phone);
    res.json({ success: true, message: 'Mempersiapkan kode tautan...' });
});

app.get('/api/wa-status', (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    res.json({
        success: true,
        statusWA: waStatus[userId] || 'BELUM_TERHUBUNG',
        qrCodeWA: qrCodes[userId] || null,
        pairingCode: pairingCodes[userId] || null
    });
});

app.get('/api/reset-wa', async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    if (waSessions[userId]) {
        try { waSessions[userId].end(undefined); } catch (e) {}
        delete waSessions[userId];
    }
    delete qrCodes[userId];
    delete pairingCodes[userId];
    waStatus[userId] = 'BELUM_TERHUBUNG';

    const authFolder = path.join(__dirname, 'auth_sessions', `user_${userId}`);
    if (fs.existsSync(authFolder)) {
        fs.rmSync(authFolder, { recursive: true, force: true });
    }
    res.json({ success: true, message: 'Sesi WA Berhasil Direset!' });
});

app.listen(PORT, () => console.log(`🚀 Server Presensi Aktif di Port ${PORT}`));
