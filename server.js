const express = require('express');
const session = require('express-session');
const { Pool } = require('pg');
const QRCode = require('qrcode');
const path = require('path');
const pino = require('pino');
const { makeWASocket, useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');

// Package Tambahan untuk Fitur Ekspor Rekap Bulanan
const ExcelJS = require('exceljs');
const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, AlignmentType } = require('docx');
const PDFDocument = require('pdfkit');

const app = express();
const PORT = process.env.PORT || 3000;

// Database Connection (Neon PostgreSQL)
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: { rejectUnauthorized: false }
});

// Middleware setup
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.use(session({
    secret: 'secret-key-presensi-sd',
    resave: false,
    saveUninitialized: true,
    cookie: { secure: false }
}));

// Variabel Penyimpanan Sesi WhatsApp Gateway
const waSessions = {};

// Fungsi Inisialisasi Layanan WhatsApp Gateway
async function initWA(userId) {
    if (waSessions[userId] && waSessions[userId].sock) {
        return waSessions[userId];
    }

    const authFolder = path.join(__dirname, 'auth_info_baileys', `user_${userId}`);
    const { state, saveCreds } = await useMultiFileAuthState(authFolder);

    const sock = makeWASocket({
        auth: state,
        logger: pino({ level: 'silent' })
    });

    waSessions[userId] = {
        sock,
        qrCode: null,
        status: 'MENUNGGU_SCAN'
    };

    sock.ev.on('creds.update', saveCreds);

    sock.ev.on('connection.update', async (update) => {
        const { connection, lastDisconnect, qr } = update;

        if (qr) {
            waSessions[userId].qrCode = await QRCode.toDataURL(qr);
            waSessions[userId].status = 'MENUNGGU_SCAN';
        }

        if (connection === 'open') {
            waSessions[userId].qrCode = null;
            waSessions[userId].status = 'TERHUBUNG';
            console.log(`WhatsApp Gateway User #${userId} Terhubung!`);
        } else if (connection === 'close') {
            const shouldReconnect = (lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut);
            waSessions[userId].status = 'TERPUTUS';
            if (shouldReconnect) {
                initWA(userId);
            }
        }
    });

    return waSessions[userId];
}

// ================= ROUTES HALAMAN =================

// Halaman Utama / Login
app.get('/', (req, res) => {
    res.render('login');
});

// Portal Wali Kelas Dashboard
app.get('/walikelas-dashboard', async (req, res) => {
    const userId = req.session.userId || 1; // Fallback ID

    try {
        const userRes = await pool.query(`
            SELECT u.*, k.nama_kelas 
            FROM users u 
            LEFT JOIN kelas k ON u.kelas_id = k.id 
            WHERE u.id = $1
        `, [userId]);

        const user = userRes.rows[0] || { nama: 'Guru Wali Kelas', nama_kelas: 'Kelas Belum Set', kelas_id: null };

        let siswaList = [];
        if (user.kelas_id) {
            const siswaRes = await pool.query(`SELECT * FROM siswa WHERE kelas_id = $1 ORDER BY nama ASC`, [user.kelas_id]);
            siswaList = await Promise.all(siswaRes.rows.map(async (s) => {
                const qrImage = await QRCode.toDataURL(s.id.toString());
                return { ...s, qrImage };
            }));
        }

        let absensiHariIni = [];
        if (user.kelas_id) {
            const absensiRes = await pool.query(`
                SELECT a.*, s.nama AS nama_siswa,
                       TO_CHAR(a.waktu AT TIME ZONE 'Asia/Jakarta', 'HH24:MI:SS DD-MM-YYYY') AS waktu_formatted
                FROM absensi a
                JOIN siswa s ON a.siswa_id = s.id
                WHERE s.kelas_id = $1 AND DATE(a.waktu AT TIME ZONE 'Asia/Jakarta') = CURRENT_DATE
                ORDER BY a.waktu DESC
            `, [user.kelas_id]);
            absensiHariIni = absensiRes.rows;
        }

        const waData = waSessions[userId] || { status: 'BELUM_TERHUBUNG', qrCode: null };

        res.render('walikelas-dashboard', {
            user,
            userId,
            siswaList,
            absensiHariIni,
            statusWA: waData.status,
            qrCodeWA: waData.qrCode
        });

    } catch (err) {
        console.error("Dashboard Error:", err);
        res.status(500).send("Terjadi Kesalahan Server: " + err.message);
    }
});

// Endpoint Memicu Koneksi WhatsApp Gateway
app.get('/api/start-wa', async (req, res) => {
    const userId = req.query.userId || req.session.userId || 1;
    await initWA(userId);
    res.json({ success: true, message: "Proses inisialisasi WA dimulai." });
});

// Endpoint Registrasi Siswa Baru
app.post('/api/siswa/tambah', async (req, res) => {
    const { nama, nomor_wa_ortu, kelas_id } = req.body;
    try {
        await pool.query(
            `INSERT INTO siswa (nama, nomor_wa_ortu, kelas_id) VALUES ($1, $2, $3)`,
            [nama, nomor_wa_ortu, kelas_id]
        );
        res.redirect('/walikelas-dashboard');
    } catch (err) {
        console.error("Tambah Siswa Error:", err);
        res.status(500).send("Gagal menambah siswa.");
    }
});

// Endpoint Hapus Siswa
app.post('/api/siswa/hapus/:id', async (req, res) => {
    const { id } = req.params;
    try {
        await pool.query(`DELETE FROM siswa WHERE id = $1`, [id]);
        res.redirect('/walikelas-dashboard');
    } catch (err) {
        console.error("Hapus Siswa Error:", err);
        res.status(500).send("Gagal menghapus siswa.");
    }
});

// ================= FITUR REKAP PRESENSI BULANAN =================

// 1. Endpoint JSON Preview Rekapitulasi Presensi
app.get('/api/absensi/preview', async (req, res) => {
    const { bulan, tahun, kelas_id } = req.query;

    if (!bulan || !tahun) {
        return res.status(400).json({ success: false, message: "Bulan dan Tahun wajib diisi." });
    }

    try {
        let query = `
            SELECT 
                s.nama AS nama_siswa,
                COALESCE(k.nama_kelas, '-') AS nama_kelas,
                COUNT(a.id) AS total_hadir
            FROM siswa s
            LEFT JOIN kelas k ON s.kelas_id = k.id
            LEFT JOIN absensi a ON s.id = a.siswa_id 
                AND EXTRACT(MONTH FROM a.waktu AT TIME ZONE 'Asia/Jakarta') = $1
                AND EXTRACT(YEAR FROM a.waktu AT TIME ZONE 'Asia/Jakarta') = $2
        `;

        const queryParams = [parseInt(bulan), parseInt(tahun)];

        if (kelas_id && kelas_id !== 'all') {
            query += ` WHERE s.kelas_id = $3`;
            queryParams.push(parseInt(kelas_id));
        }

        query += ` GROUP BY s.id, s.nama, k.nama_kelas ORDER BY s.nama ASC`;

        const result = await pool.query(query, queryParams);
        return res.json({ success: true, data: result.rows });

    } catch (err) {
        console.error("Preview API Error:", err);
        return res.status(500).json({ success: false, message: err.message });
    }
});

// 2. Endpoint Export Rekapitulasi Presensi (Excel, Word, PDF)
app.get('/api/absensi/export', async (req, res) => {
    const { bulan, tahun, kelas_id, format } = req.query;

    if (!bulan || !tahun) {
        return res.status(400).send("Bulan dan Tahun wajib diisi.");
    }

    try {
        let query = `
            SELECT 
                s.nama AS nama_siswa,
                COALESCE(k.nama_kelas, '-') AS nama_kelas,
                COUNT(a.id) AS total_hadir
            FROM siswa s
            LEFT JOIN kelas k ON s.kelas_id = k.id
            LEFT JOIN absensi a ON s.id = a.siswa_id 
                AND EXTRACT(MONTH FROM a.waktu AT TIME ZONE 'Asia/Jakarta') = $1
                AND EXTRACT(YEAR FROM a.waktu AT TIME ZONE 'Asia/Jakarta') = $2
        `;

        const queryParams = [parseInt(bulan), parseInt(tahun)];

        if (kelas_id && kelas_id !== 'all') {
            query += ` WHERE s.kelas_id = $3`;
            queryParams.push(parseInt(kelas_id));
        }

        query += ` GROUP BY s.id, s.nama, k.nama_kelas ORDER BY s.nama ASC`;

        const result = await pool.query(query, queryParams);
        const dataRekap = result.rows;

        const namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"][parseInt(bulan) - 1];
        const judul = `REKAP PRESENSI SISWA - ${namaBulan.toUpperCase()} ${tahun}`;

        // Format Excel (.xlsx)
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

        // Format Word (.docx)
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

        // Format PDF (.pdf) - Menggunakan PDFKit standar tanpa modul bermasalah
        if (format === 'pdf') {
            const doc = new PDFDocument({ margin: 40, size: 'A4' });
            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', `attachment; filename=Rekap_Presensi_${bulan}_${tahun}.pdf`);

            doc.pipe(res);

            // Header Dokumen
            doc.fontSize(14).font('Helvetica-Bold').text('UPTD SD NEGERI 1 KARYA MULYA SARI', { align: 'center' });
            doc.fontSize(11).font('Helvetica').text(judul, { align: 'center' });
            doc.moveDown(1.5);

            // Pengaturan Tabel PDF
            let y = doc.y;
            const startX = 40;
            const colWidths = [40, 230, 130, 100]; // Lebar kolom No, Nama, Kelas, Total Hadir

            // Header Tabel
            doc.font('Helvetica-Bold').fontSize(10);
            doc.text('No', startX, y);
            doc.text('Nama Siswa', startX + colWidths[0], y);
            doc.text('Kelas', startX + colWidths[0] + colWidths[1], y);
            doc.text('Total Hadir', startX + colWidths[0] + colWidths[1] + colWidths[2], y);

            doc.moveTo(startX, y + 15).lineTo(startX + 500, y + 15).stroke();
            y += 22;

            // Isi Baris Tabel
            doc.font('Helvetica').fontSize(9);
            dataRekap.forEach((row, i) => {
                if (y > 750) { // Auto Tambah Halaman Baru jika Halaman Penuh
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

        return res.status(400).send("Format ekspor tidak valid (Gunakan format=excel, word, atau pdf).");

    } catch (err) {
        console.error("Export Error:", err);
        res.status(500).send("Gagal mengekspor data: " + err.message);
    }
});

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
