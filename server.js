const usePostgresAuthState = require('./usePostgresAuthState');
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

const waSessions = {};
const qrCodes = {};
const waStatus = {};
const pairingCodes = {};
const reconnectTimers = {}; // Mencegah looping restart berulang

function bersihkanGelar(nama) {
    if (!nama) return '';
    return nama.replace(/,?\s*\b(S\.Pd|M\.Pd|S\.Ag|S\.T|S\.Kom|M\.Si|S\.Sos|S\.SE|M\.M|A\.Ma|Sd)\b\.?/gi, '').trim();
}

// System QR Safe Generator
async function generateQRDataURL(text) {
    try {
        if (!text) return 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORTH5CYII=';
        return await QRCode.toDataURL(text.toString());
    } catch (err) {
        console.error("Gagal Generate QR:", err.message);
        return 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORTH5CYII=';
    }
}

// ----------------- HYBRID AUTH STATE (LOKAL AUTH FOLDER) ----------------- //

async function getAuthState(userId) {
    return await usePostgresAuthState(userId);
}
async function connectToWhatsApp(userId, phoneNumber = null) {
    try {
        // Bersihkan timer pending jika ada request ulang
        if (reconnectTimers[userId]) {
            clearTimeout(reconnectTimers[userId]);
            delete reconnectTimers[userId];
        }

        if (waSessions[userId]) {
            try { waSessions[userId].end(undefined); } catch (e) {}
            delete waSessions[userId];
        }

        waStatus[userId] = phoneNumber ? 'MENUNGGU_PAIRING_CODE' : 'PROSES_INIT';
        delete qrCodes[userId];
        delete pairingCodes[userId];

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

        // MINTA KODE PAIRING (JIKA DIPANGGUL DENGAN NOMOR HP)
        if (phoneNumber && !sock.authState.creds.registered) {
            setTimeout(async () => {
                try {
                    let cleanPhone = phoneNumber.toString().replace(/[^0-9]/g, '');
                    if (cleanPhone.startsWith('0')) cleanPhone = '62' + cleanPhone.slice(1);
                    
                    console.log(`📱 Meminta Pairing Code WA untuk nomor: ${cleanPhone}`);
                    const code = await sock.requestPairingCode(cleanPhone);
                    pairingCodes[userId] = code;
                    waStatus[userId] = 'MENUNGGU_PAIRING_CODE';
                    console.log(`🔑 [User #${userId}] Pairing Code WA Terbit: ${code}`);
                } catch (pErr) {
                    console.error("Gagal Request Pairing Code:", pErr.message);
                    waStatus[userId] = 'ERROR_PAIRING';
                }
            }, 5000);
        }

        sock.ev.on('connection.update', async (update) => {
            const { connection, lastDisconnect, qr } = update;

            // QR CODE GENERATION (DENGAN DEBOUNCE TERKONTROL)
            if (qr && !phoneNumber && !sock.authState.creds.registered) {
                try {
                    qrCodes[userId] = await generateQRDataURL(qr);
                    waStatus[userId] = 'MENUNGGU_SCAN';

                    console.log(`\n========================================`);
                    console.log(`📲 [User #${userId}] SCAN QR CODE TERBIT`);
                    console.log(`========================================\n`);
                    qrcodeTerminal.generate(qr, { small: true });

                } catch (qrErr) {
                    console.error("Gagal generate QR Code WA:", qrErr);
                }
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
                // Hanya hapus sesi jika di-logout resmi dari HP (401 / DisconnectReason.loggedOut)
                const isLoggedOut = (statusCode === DisconnectReason.loggedOut || statusCode === 401);

                console.log(`⚠️ [User #${userId}] WhatsApp Terputus. Status Code: ${statusCode}`);
                waStatus[userId] = 'TERPUTUS';
                delete waSessions[userId];

                if (!isLoggedOut) {
                    // Beri jeda 8 detik agar tidak looping kedap-kedip
                    console.log(`🔄 Sambung ulang User #${userId} dalam 8 detik...`);
                    if (!reconnectTimers[userId]) {
                        reconnectTimers[userId] = setTimeout(() => {
                            delete reconnectTimers[userId];
                            connectToWhatsApp(userId);
                        }, 8000);
                    }
                } else {
                    console.log(`🚪 User #${userId} Logged Out. Menghapus sesi lokal & database...`);
                    delete qrCodes[userId];
                    delete pairingCodes[userId];

                    // Panggil clearCreds dari auth handler hybrid untuk hapus lokal + DB Neon
                    getAuthState(userId).then(({ clearCreds }) => clearCreds()).catch(e => console.error("Gagal clearCreds:", e.message));
                }
            }
        });
    } catch (err) {
        console.error(`❌ WA Connect Error User #${userId}:`, err.message);
        waStatus[userId] = 'ERROR';
    }
}

// ---------------- ROUTES HALAMAN ---------------- //

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

        const absensiRes = await pool.query(`
            SELECT a.id, a.waktu, s.nama AS nama_siswa, COALESCE(k.nama_kelas, '-') AS nama_kelas 
            FROM absensi a 
            JOIN siswa s ON a.siswa_id = s.id 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            WHERE DATE(a.waktu AT TIME ZONE 'Asia/Jakarta') = CURRENT_DATE
            ORDER BY a.waktu DESC
        `);

        const absensiFormatted = absensiRes.rows.map(row => {
            const dateObj = new Date(row.waktu);
            const waktuWIB = dateObj.toLocaleTimeString('id-ID', {
                timeZone: 'Asia/Jakarta',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            }).replace(/\./g, ':') + ' WIB';

            return { ...row, waktu_formatted: waktuWIB };
        });

        const usersCleaned = usersRes.rows.map(u => ({ ...u, nama: bersihkanGelar(u.nama) }));
        const siswaData = await Promise.all(siswaRes.rows.map(async (s) => {
            const qrImage = await generateQRDataURL(s.id.toString());
            return { ...s, qrImage };
        }));

        res.render('admin-dashboard', {
            users: usersCleaned,
            siswa: siswaData,
            kelas: kelasRes.rows || [],
            absensiHariIni: absensiFormatted,
            userId: userId
        });
    } catch (err) {
        res.status(500).send("Kesalahan Database: " + err.message);
    }
});

app.get(['/admin/cetak-kartu', '/cetak-kartu'], async (req, res) => {
    try {
        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, '-') AS nama_kelas 
            FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id ORDER BY s.nama ASC
        `);

        const siswaData = await Promise.all(siswaRes.rows.map(async (s) => {
            const qrImage = await generateQRDataURL(s.id.toString());
            return { ...s, qrImage };
        }));

        res.render('cetak-kartu', { siswa: siswaData });
    } catch (err) {
        res.status(500).send("Gagal memuat kartu: " + err.message);
    }
});

app.get(['/wali', '/walikelas-dashboard'], async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId;
    if (!userId) return res.redirect('/');

    try {
        const userRes = await pool.query(`
            SELECT u.id, u.nama, u.role, u.kelas_id, COALESCE(k.nama_kelas, 'Guru Mata Pelajaran (Semua Kelas)') AS nama_kelas
            FROM users u
            LEFT JOIN kelas k ON u.kelas_id = k.id
            WHERE u.id = $1
        `, [userId]);

        if (userRes.rows.length === 0) return res.redirect('/');
        
        const userRaw = userRes.rows[0];
        userRaw.nama = bersihkanGelar(userRaw.nama);

        let siswaQuery = `
            SELECT s.id, s.nama, s.nomor_wa_ortu, s.kelas_id, COALESCE(k.nama_kelas, '-') AS nama_kelas 
            FROM siswa s 
            LEFT JOIN kelas k ON s.kelas_id = k.id
        `;
        const queryParamsSiswa = [];

        if (userRaw.kelas_id) {
            siswaQuery += ` WHERE s.kelas_id = $1`;
            queryParamsSiswa.push(parseInt(userRaw.kelas_id));
        }

        siswaQuery += ` ORDER BY s.nama ASC`;
        const siswaRes = await pool.query(siswaQuery, queryParamsSiswa);

        let absensiQuery = `
            SELECT a.id, a.waktu, s.nama AS nama_siswa, COALESCE(k.nama_kelas, '-') AS nama_kelas 
            FROM absensi a 
            JOIN siswa s ON a.siswa_id = s.id 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            WHERE DATE(a.waktu AT TIME ZONE 'Asia/Jakarta') = CURRENT_DATE
        `;
        const queryParamsAbsensi = [];

        if (userRaw.kelas_id) {
            absensiQuery += ` AND s.kelas_id = $1`;
            queryParamsAbsensi.push(parseInt(userRaw.kelas_id));
        }

        absensiQuery += ` ORDER BY a.waktu DESC`;
        const absensiRes = await pool.query(absensiQuery, queryParamsAbsensi);

        const absensiFormatted = absensiRes.rows.map(row => {
            const dateObj = new Date(row.waktu);
            const waktuWIB = dateObj.toLocaleTimeString('id-ID', {
                timeZone: 'Asia/Jakarta',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            }).replace(/\./g, ':') + ' WIB';

            return { ...row, waktu_formatted: waktuWIB };
        });

        const siswaData = await Promise.all(siswaRes.rows.map(async (s) => {
            const qrImage = await generateQRDataURL(s.id.toString());
            return { ...s, qrImage };
        }));

        req.session.userId = userId;

        res.render('walikelas-dashboard', {
            user: userRaw,
            siswaList: siswaData,
            absensiHariIni: absensiFormatted,
            userId: userId,
            statusWA: waStatus[userId] || 'BELUM_TERHUBUNG',
            qrCodeWA: qrCodes[userId] || null
        });
    } catch (err) {
        console.error("Dashboard Error User #" + userId + ":", err);
        res.status(500).send("Kesalahan Database: " + err.message);
    }
});

app.get(['/scan', '/scanner'], (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    res.render('scan', { userId: userId });
});

// ---------------- API KELOLA KELAS & ROMBEL ---------------- //

app.post('/api/kelas/tambah', async (req, res) => {
    const { nama_kelas } = req.body;
    try {
        if (!nama_kelas) return res.status(400).send("Nama kelas wajib diisi.");
        await pool.query('INSERT INTO kelas (nama_kelas) VALUES ($1)', [nama_kelas.trim()]);
        return res.redirect(`/admin?userId=${req.session.userId || 1}`);
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

app.post('/api/siswa/hapus-semua', async (req, res) => {
    try {
        // FITUR HAPUS DI-DISABLE DEMI KEAMANAN DATA
        console.log("⚠️ Fitur reset total diblokir demi keamanan data.");
        return res.redirect(`/admin?userId=${req.session.userId || 1}`);
    } catch (err) {
        return res.status(500).send("Gagal: " + err.message);
    }
});

// ---------------- API GURU & SISWA ---------------- //

app.post('/api/guru/tambah', async (req, res) => {
    const { nama, username, password, kelas_id } = req.body;
    try {
        if (!nama || !username || !password) return res.status(400).send("Semua kolom wajib diisi.");
        const parsedKelasId = kelas_id ? parseInt(kelas_id) : null;
        await pool.query(
            'INSERT INTO users (nama, username, password, role, kelas_id) VALUES ($1, $2, $3, $4, $5)',
            [nama.trim(), username.trim(), password.trim(), 'WALI_KELAS', parsedKelasId]
        );
        return res.redirect(`/admin?userId=${req.session.userId || 1}`);
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
        await pool.query('DELETE FROM users WHERE id = $1 AND role != $2', [parseInt(req.params.id), 'ADMIN']);
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

// ---------------- API IMPOR EXCEL SISWA DAPODIK ---------------- //

app.post('/api/siswa/import-excel', upload.single('file_excel'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ success: false, message: "Berkas Excel/CSV wajib diunggah!" });
        }

        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });

        if (sheetData.length === 0) {
            return res.status(400).json({ success: false, message: "Berkas Excel kosong atau tidak terbaca." });
        }

        let totalBerhasil = 0;

        for (const row of sheetData) {
            let namaSiswa = "";
            let nomorWa = "";
            let namaKelas = null;

            Object.keys(row).forEach(key => {
                const cleanKey = key.toString().toLowerCase().trim();
                const val = row[key] ? row[key].toString().trim() : "";

                if (cleanKey.includes('nama') || cleanKey.includes('peserta didik') || cleanKey.includes('siswa')) {
                    if (!namaSiswa && val) namaSiswa = val;
                }
                if (cleanKey.includes('hp') || cleanKey.includes('wa') || cleanKey.includes('telepon') || cleanKey.includes('seluler') || cleanKey.includes('kontak')) {
                    if (!nomorWa && val) nomorWa = val;
                }
                if (cleanKey.includes('rombel') || cleanKey.includes('kelas') || cleanKey.includes('rombongan') || cleanKey.includes('tingkat')) {
                    if (!namaKelas && val) namaKelas = val;
                }
            });

            if (namaSiswa && namaSiswa.length > 1) {
                let kelasId = null;

                if (namaKelas) {
                    const cleanNamaKelas = namaKelas.toString().replace(/\s+/g, ' ').trim();
                    let kRes = await pool.query('SELECT id FROM kelas WHERE LOWER(TRIM(nama_kelas)) = LOWER($1)', [cleanNamaKelas]);
                    
                    if (kRes.rows.length > 0) {
                        kelasId = kRes.rows[0].id;
                    } else {
                        let newKRes = await pool.query('INSERT INTO kelas (nama_kelas) VALUES ($1) RETURNING id', [cleanNamaKelas]);
                        kelasId = newKRes.rows[0].id;
                    }
                }

                await pool.query(
                    `INSERT INTO siswa (nama, nomor_wa_ortu, kelas_id) VALUES ($1, $2, $3)`,
                    [namaSiswa, nomorWa, kelasId]
                );
                totalBerhasil++;
            }
        }

        if (totalBerhasil === 0) {
            return res.json({
                success: false,
                message: "Gagal membaca data. Pastikan ada kolom 'Nama' atau 'Nama Siswa' di baris pertama Excel."
            });
        }

        return res.json({
            success: true,
            message: `Berhasil mengimpor ${totalBerhasil} data siswa dari Dapodik!`
        });

    } catch (err) {
        console.error("Import Excel Error:", err);
        return res.status(500).json({ success: false, message: "Gagal memproses berkas: " + err.message });
    }
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
    if (reconnectTimers[userId]) {
        clearTimeout(reconnectTimers[userId]);
        delete reconnectTimers[userId];
    }
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

app.post('/api/scan', async (req, res) => {
    const { siswa_id, scanned_by } = req.body;
    if (!siswa_id) return res.status(400).json({ success: false, message: "Kode QR tidak terdeteksi." });

    try {
        const parsedSiswaId = parseInt(siswa_id);
        const parsedScannedBy = parseInt(scanned_by) || 1;

        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, '-') AS nama_kelas 
            FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id WHERE s.id = $1
        `, [parsedSiswaId]);

        if (siswaRes.rows.length === 0) {
            return res.status(404).json({ success: false, message: `ID Siswa #${siswa_id} Tidak Terdaftar!` });
        }

        const siswa = siswaRes.rows[0];

        const now = new Date();
        const jamWib = now.toLocaleTimeString('id-ID', { timeZone: 'Asia/Jakarta', hour: '2-digit', minute: '2-digit', second: '2-digit' }).replace(/\./g, ':') + ' WIB';
        const tglWib = now.toLocaleDateString('id-ID', { timeZone: 'Asia/Jakarta', weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

        await pool.query(
            `INSERT INTO absensi (siswa_id, status, scanned_by, waktu) VALUES ($1, 'HADIR', $2, NOW() AT TIME ZONE 'Asia/Jakarta')`,
            [siswa.id, parsedScannedBy]
        );

        let waClient = waSessions[parsedScannedBy];
        if (!waClient) {
            const keys = Object.keys(waSessions);
            if (keys.length > 0) waClient = waSessions[keys[0]];
        }

        let statusWA = "Notifikasi WhatsApp Tidak Terkirim (Layanan WA Belum Terkoneksi)";

        if (waClient && siswa.nomor_wa_ortu) {
            let phone = siswa.nomor_wa_ortu.toString().trim().replace(/[^0-9]/g, '');
            if (phone.startsWith('0')) phone = '62' + phone.slice(1);
            const formattedJid = phone + '@s.whatsapp.net';

            const pesan = `*UPTD SD NEGERI 1 KARYA MULYA SARI*\n` +
                          `*PEMBERITAHUAN PRESENSI KEHADIRAN SISWA*\n` +
                          `_________________________________________\n\n` +
                          `Yth. Bapak/Ibu Orang Tua / Wali Murid,\n\n` +
                          `Diberitahukan bahwa putra/putri Anda telah tiba di sekolah dan melakukan presensi kehadiran:\n\n` +
                          `• Nama Siswa : *${siswa.nama}*\n` +
                          `• Kelas / Rombel : *${siswa.nama_kelas}*\n` +
                          `• Waktu Scan : *${jamWib}*\n` +
                          `• Tanggal : *${tglWib}*\n` +
                          `• Status Kehadiran : *HADIR (Scan Kartu) ✅*\n\n` +
                          `Terima kasih atas perhatian dan kerja samanya.\n\n` +
                          `_Pesan otomatis ini dikirim oleh Sistem Presensi Terpadu UPTD SD Negeri 1 Karya Mulya Sari._`;

            waClient.sendMessage(formattedJid, { text: pesan }).catch(e => console.error("Gagal Mengirim WA:", e.message));
            statusWA = "Notifikasi WhatsApp Berhasil Dikirimkan ke Wali Murid ✅";
        }

        return res.json({
            success: true,
            message: statusWA,
            siswa: {
                id: siswa.id,
                nama: siswa.nama,
                nama_kelas: siswa.nama_kelas,
                waktu: jamWib
            }
        });

    } catch (err) {
        console.error("Kesalahan Scan:", err);
        return res.status(500).json({ success: false, message: "Kendala Sistem: " + err.message });
    }
});

// ---------------- API HAPUS RIWAYAT ABSENSI UJI COBA ---------------- //
app.post('/api/absensi/reset-riwayat', async (req, res) => {
    try {
        await pool.query('DELETE FROM absensi');
        await pool.query('ALTER SEQUENCE absensi_id_seq RESTART WITH 1');

        console.log("🧹 Riwayat absensi uji coba berhasil dibersihkan.");

        const backUrl = req.headers.referer || `/admin?userId=${req.session.userId || 1}`;
        return res.redirect(backUrl);
    } catch (err) {
        console.error("Gagal menghapus riwayat absensi:", err);
        return res.status(500).send("Gagal membersihkan riwayat absensi: " + err.message);
    }
});

// ---------------- FITUR REKAP BULANAN ---------------- //

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
    if (!bulan || !tahun) return res.status(400).send("Bulan dan Tahun wajib diisi.");

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
        const dataRekap = result.rows;
        const namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"][parseInt(bulan) - 1];
        const judul = `REKAP PRESENSI SISWA - ${namaBulan.toUpperCase()} ${tahun}`;

        if (format === 'excel') {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Rekap Absensi');

            worksheet.mergeCells('A1:D1');
            worksheet.getCell('A1').value = `UPTD SD NEGERI 1 KARYA MULYA SARI`;
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
                        new Paragraph({ text: "UPTD SD NEGERI 1 KARYA MULYA SARI", heading: "Heading1", alignment: AlignmentType.CENTER }),
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
            doc.fontSize(14).font('Helvetica-Bold').text('UPTD SD NEGERI 1 KARYA MULYA SARI', { align: 'center' });
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
        console.error("Export Error:", err);
        res.status(500).send("Gagal mengekspor data: " + err.message);
    }
});
app.get('/ping', (req, res) => res.send('OK'));
app.listen(PORT, () => console.log(`🚀 Server Presensi Aktif di Port ${PORT}`));
