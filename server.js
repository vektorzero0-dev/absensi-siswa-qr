const express = require('express');
const { default: makeWASocket, useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const QRCode = require('qrcode');
const path = require('path');
const db = require('./db');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.set('view engine', 'ejs');
app.use(express.static(path.join(__dirname, 'public')));

const waSessions = {};
const qrCodesMemory = {};

async function initWaSession(userId) {
    if (waSessions[userId]) return waSessions[userId];
    const { state, saveCreds } = await useMultiFileAuthState(`./sessions/user_${userId}`);
    const sock = makeWASocket({ auth: state, printQRInTerminal: false });

    sock.ev.on('creds.update', saveCreds);
    sock.ev.on('connection.update', async (update) => {
        const { connection, lastDisconnect, qr } = update;
        if (qr) qrCodesMemory[userId] = await QRCode.toDataURL(qr);
        if (connection === 'open') {
            delete qrCodesMemory[userId];
            waSessions[userId] = sock;
        }
        if (connection === 'close') {
            const shouldReconnect = (lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut);
            delete waSessions[userId];
            if (shouldReconnect) initWaSession(userId);
        }
    });
    return sock;
}

app.get('/', (req, res) => res.render('login'));

app.post('/login', async (req, res) => {
    const { username, password } = req.body;
    try {
        const user = await db.query('SELECT * FROM users WHERE username = $1 AND password = $2', [username, password]);
        if (user.rows.length === 0) return res.send('Username/Password Salah!');
        const userData = user.rows[0];
        if (userData.role === 'ADMIN_OPS') {
            res.redirect(`/admin?userId=${userData.id}`);
        } else {
            res.redirect(`/wali?userId=${userData.id}`);
        }
    } catch (err) {
        res.status(500).send('Database Error');
    }
});

app.get('/admin', async (req, res) => {
    try {
        const siswa = await db.query('SELECT siswa.*, kelas.nama_kelas FROM siswa JOIN kelas ON siswa.kelas_id = kelas.id');
        const absensi = await db.query('SELECT absensi.*, siswa.nama, kelas.nama_kelas FROM absensi JOIN siswa ON absensi.siswa_id = siswa.id JOIN kelas ON siswa.kelas_id = kelas.id ORDER BY waktu DESC LIMIT 20');
        res.render('admin-dashboard', { siswa: siswa.rows, absensi: absensi.rows });
    } catch (err) {
        res.status(500).send('Error');
    }
});

app.get('/wali', async (req, res) => {
    const userId = req.query.userId;
    try {
        const user = await db.query('SELECT users.*, kelas.nama_kelas FROM users LEFT JOIN kelas ON users.kelas_id = kelas.id WHERE users.id = $1', [userId]);
        initWaSession(userId);
        const statusWA = waSessions[userId] ? 'TERHUBUNG' : 'BELUM_TERHUBUNG';
        const qrCodeWA = qrCodesMemory[userId] || null;
        res.render('walikelas-dashboard', { user: user.rows[0], statusWA, qrCodeWA, userId });
    } catch (err) {
        res.status(500).send('Error');
    }
});

app.get('/scan', (req, res) => {
    res.render('scan', { userId: req.query.userId });
});

app.post('/api/absen', async (req, res) => {
    const { siswaId, waliKelasId } = req.body;
    const sock = waSessions[waliKelasId];
    if (!sock) return res.status(400).json({ success: false, message: 'WhatsApp Anda belum terhubung!' });

    try {
        const siswaResult = await db.query('SELECT * FROM siswa WHERE id = $1', [siswaId]);
        if (siswaResult.rows.length === 0) return res.status(404).json({ success: false, message: 'Kartu Siswa Tidak Valid!' });
        const siswa = siswaResult.rows[0];

        await db.query('INSERT INTO absensi (siswa_id, scanned_by) VALUES ($1, $2)', [siswaId, waliKelasId]);

        const jam = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
        const pesanWA = `Assalamu'alaikum Bapak/Ibu,\n\nMemberitahukan bahwa siswa atas nama *${siswa.nama}* telah hadir di sekolah pada pukul *${jam} WIB*.\n\n_Pesan otomatis dari Wali Kelas._`;

        await sock.sendMessage(`${siswa.nomor_wa_ortu}@s.whatsapp.net`, { text: pesanWA });
        res.json({ success: true, message: `Berhasil Absen: ${siswa.nama}`, namaSiswa: siswa.nama });
    } catch (error) {
        res.status(500).json({ success: false, message: 'Gagal Memproses Absensi/Mengirim WA' });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
