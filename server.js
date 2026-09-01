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

function bersihkanGelar(nama) {
    if (!nama) return '';
    return nama.replace(/,?\s*(S\.Pd|M\.Pd|S\.Ag|S\.T|S\.Kom|M\.Si|S\.Sos|S\.SE|M\.M|A\.Ma|Sd)\.?/gi, '').trim();
}

// ----------------- HYBRID AUTH STATE (FS + NEON SYNC) ----------------- //

async function getAuthState(userId) {
    const authFolder = path.join(__dirname, 'auth_sessions', `user_${userId}`);
    if (!fs.existsSync(authFolder)) {
        fs.mkdirSync(authFolder, { recursive: true });
    }
    return await useMultiFileAuthState(authFolder);
}

async function connectToWhatsApp(userId) {
    try {
        if (waSessions[userId]) {
            try { waSessions[userId].end(undefined); } catch (e) {}
            delete waSessions[userId];
        }

        waStatus[userId] = 'PROSES_INIT';
        delete qrCodes[userId];

        const { state, saveCreds } = await getAuthState(userId);
        const { version } = await fetchLatestBaileysVersion();

        console.log(`⚡ [User #${userId}] Menginisialisasi WA Socket (Baileys v${version.join('.')})...`);

        // Ganti opsi browser dan tambahkan opsi syncFullHistory: false
const sock = makeWASocket({
    logger: pino({ level: 'silent' }),
    auth: state,
    printQRInTerminal: false,
    // Gunakan signature Ubuntu / Chrome yang lebih stabil untuk Baileys
    browser: ["Ubuntu", "Chrome", "120.0.6099.109"],
    connectTimeoutMs: 60000,
    defaultQueryTimeoutMs: 60000,
    keepAliveIntervalMs: 30000,
    qrTimeout: 40000,
    // Tambahkan 2 baris ini untuk mencegah timeout penautan di Render
    syncFullHistory: false,
    markOnlineOnConnect: false
});

        waSessions[userId] = sock;
        sock.ev.on('creds.update', saveCreds);

        sock.ev.on('connection.update', async (update) => {
            const { connection, lastDisconnect, qr } = update;

            if (qr) {
                try {
                    // 1. Simpan ke DataURL untuk Web Browser Dashboard
                    qrCodes[userId] = await QRCode.toDataURL(qr);
                    waStatus[userId] = 'MENUNGGU_SCAN';

                    console.log(`\n========================================`);
                    console.log(`📲 [User #${userId}] SCAN QR CODE DI BAWAH INI:`);
                    console.log(`========================================\n`);
                    // 2. Cetak ke Log Render Terminal sebagai Fallback
                    qrcodeTerminal.generate(qr, { small: true });

                } catch (qrErr) {
                    console.error("Gagal generate QR Code:", qrErr);
                }
            }

            if (connection === 'open') {
                waStatus[userId] = 'TERHUBUNG';
                delete qrCodes[userId];
                console.log(`✅ [User #${userId}] WhatsApp Berhasil Terhubung!`);
            }

            if (connection === 'close') {
                const statusCode = lastDisconnect?.error?.output?.statusCode;
                const shouldReconnect = (statusCode !== DisconnectReason.loggedOut);

                console.log(`⚠️ [User #${userId}] WhatsApp Terputus. Reason Status: ${statusCode}`);
                waStatus[userId] = 'TERPUTUS';
                delete waSessions[userId];

                if (shouldReconnect) {
                    console.log(`🔄 Menguji sambungan ulang untuk User #${userId}...`);
                    setTimeout(() => connectToWhatsApp(userId), 3000);
                } else {
                    console.log(`🚪 User #${userId} Logged Out. Membersihkan Sesi...`);
                    delete qrCodes[userId];
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
            SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, '-') AS nama_kelas 
            FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id ORDER BY s.id ASC
        `);
        const kelasRes = await pool.query(`SELECT * FROM kelas ORDER BY id ASC`);

        const usersCleaned = usersRes.rows.map(u => ({ ...u, nama: bersihkanGelar(u.nama) }));
        const siswaData = await Promise.all(siswaRes.rows.map(async (s) => {
            const qrImage = await QRCode.toDataURL(s.id.toString());
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
            const qrImage = await QRCode.toDataURL(s.id.toString());
            return { ...s, qrImage };
        }));

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

// ---------------- API WA GATEWAY REAL-TIME ---------------- //

app.get('/api/start-wa', async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    connectToWhatsApp(userId);
    res.json({ success: true, message: 'Inisialisasi WhatsApp dimulai...' });
});

app.get('/api/wa-status', (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    res.json({
        success: true,
        statusWA: waStatus[userId] || 'BELUM_TERHUBUNG',
        qrCodeWA: qrCodes[userId] || null
    });
});

app.get('/api/reset-wa', async (req, res) => {
    const userId = parseInt(req.query.userId) || req.session.userId || 1;
    if (waSessions[userId]) {
        try { waSessions[userId].end(undefined); } catch (e) {}
        delete waSessions[userId];
    }
    delete qrCodes[userId];
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

// ---------------- FITUR REKAP BULANAN (PREVIEW & EXPORT) ---------------- //

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

app.listen(PORT, () => console.log(`🚀 Server Presensi Aktif di Port ${PORT}`));
