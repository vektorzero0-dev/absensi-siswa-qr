const express = require('express');
const path = require('path');
const pino = require('pino');
const makeWASocket = require('@whiskeysockets/baileys').default;
const { DisconnectReason, initAuthCreds, proto } = require('@whiskeysockets/baileys');
const QRCode = require('qrcode');
const pool = require('./db');

process.env.TZ = 'Asia/Jakarta';

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));

const waSessions = {};
const qrCodes = {};
const waStatus = {};

// Fungsi Pembersih Gelar Guru
function bersihkanGelar(nama) {
    if (!nama) return '';
    return nama.replace(/,?\s*(S\.Pd|M\.Pd|S\.Ag|S\.T|S\.Kom|M\.Si|S\.Sos|S\.SE|M\.M|A\.Ma|Sd)\.?/gi, '').trim();
}

// ----------------- POSTGRES AUTH STATE (SIMPAN WA DI DB) ----------------- //

async function usePostgresAuthState(pool, userId) {
    const keyPrefix = `user_${userId}:`;

    const writeData = async (data, id, key) => {
        const fullKey = keyPrefix + key;
        const jsonStr = JSON.stringify(data, (k, v) => (typeof v === 'bigint' ? v.toString() + 'n' : v));
        await pool.query(
            `INSERT INTO wa_sessions (key, id, data) VALUES ($1, $2, $3)
             ON CONFLICT (key, id) DO UPDATE SET data = EXCLUDED.data`,
            [fullKey, id, jsonStr]
        );
    };

    const readData = async (id, key) => {
        const fullKey = keyPrefix + key;
        const res = await pool.query(`SELECT data FROM wa_sessions WHERE key = $1 AND id = $2`, [fullKey, id]);
        if (res.rows.length > 0) {
            return JSON.parse(res.rows[0].data, (k, v) => {
                if (typeof v === 'string' && /^\d+n$/.test(v)) return BigInt(v.slice(0, -1));
                return v;
            });
        }
        return null;
    };

    const removeData = async (id, key) => {
        const fullKey = keyPrefix + key;
        await pool.query(`DELETE FROM wa_sessions WHERE key = $1 AND id = $2`, [fullKey, id]);
    };

    let creds = await readData('creds', 'main');
    if (!creds) {
        creds = initAuthCreds();
        await writeData(creds, 'creds', 'main');
    }

    return {
        state: {
            creds,
            keys: {
                get: async (type, ids) => {
                    const data = {};
                    await Promise.all(
                        ids.map(async (id) => {
                            let value = await readData(id, type);
                            if (type === 'app-state-sync-key' && value) {
                                value = proto.Message.AppStateSyncKeyData.fromObject(value);
                            }
                            data[id] = value;
                        })
                    );
                    return data;
                },
                set: async (data) => {
                    const tasks = [];
                    for (const category in data) {
                        for (const id in data[category]) {
                            const value = data[category][id];
                            if (value) {
                                tasks.push(writeData(value, id, category));
                            } else {
                                tasks.push(removeData(id, category));
                            }
                        }
                    }
                    await Promise.all(tasks);
                }
            }
        },
        saveCreds: async () => {
            await writeData(creds, 'creds', 'main');
        }
    };
}

async function connectToWhatsApp(userId) {
    try {
        const { state, saveCreds } = await usePostgresAuthState(pool, userId);

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
                console.log(`✅ WhatsApp Gateway User #${userId} Terhubung & Tersimpan di DB`);
            }
            if (connection === 'close') {
                const statusCode = lastDisconnect?.error?.output?.statusCode;
                const shouldReconnect = (statusCode !== DisconnectReason.loggedOut);
                
                waStatus[userId] = 'TERPUTUS';
                delete waSessions[userId];
                
                if (shouldReconnect) {
                    console.log(`🔄 Menyambungkan ulang WA User #${userId}...`);
                    connectToWhatsApp(userId);
                } else {
                    await pool.query(`DELETE FROM wa_sessions WHERE key LIKE $1`, [`user_${userId}:%`]);
                    console.log(`🗑️ Sesi User #${userId} dihapus dari Database (Logged Out).`);
                }
            }
        });
    } catch (err) {
        console.error(`❌ WA Connect Error User #${userId}:`, err.message);
    }
}

// Otomatis Muat Sesi WA dari Database Saat Server Booting
async function autoLoadSavedSessions() {
    try {
        const res = await pool.query(`SELECT DISTINCT key FROM wa_sessions`);
        const userIds = new Set();
        
        res.rows.forEach(row => {
            const match = row.key.match(/^user_(\d+):/);
            if (match) userIds.add(match[1]);
        });

        userIds.forEach(id => {
            console.log(`⚙️ Mengembalikan koneksi WA terdaftar untuk User #${id}...`);
            connectToWhatsApp(id);
        });
    } catch (err) {
        console.error("Gagal auto-load sesi WA:", err.message);
    }
}

// Inisialisasi auto-load sesi setelah server siap
setTimeout(() => {
    autoLoadSavedSessions();
}, 2000);

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
    const userId = parseInt(req.query.userId) || 1;
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

app.get('/wali', async (req, res) => {
    const userId = parseInt(req.query.userId) || 1;
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
        if (userRaw.kelas_id) siswaQuery += ` WHERE s.kelas_id = ${parseInt(userRaw.kelas_id)}`;
        siswaQuery += ` ORDER BY s.nama ASC`;
        const siswaRes = await pool.query(siswaQuery);

        let absensiQuery = `
            SELECT a.id, a.waktu, s.nama AS nama_siswa, COALESCE(k.nama_kelas, '-') AS nama_kelas 
            FROM absensi a 
            JOIN siswa s ON a.siswa_id = s.id 
            LEFT JOIN kelas k ON s.kelas_id = k.id 
            WHERE DATE(a.waktu AT TIME ZONE 'UTC' AT TIME ZONE 'Asia/Jakarta') = CURRENT_DATE
        `;
        if (userRaw.kelas_id) absensiQuery += ` AND s.kelas_id = ${parseInt(userRaw.kelas_id)}`;
        absensiQuery += ` ORDER BY a.waktu DESC`;
        const absensiRes = await pool.query(absensiQuery);

        const absensiFormatted = absensiRes.rows.map(row => {
            const dateObj = new Date(row.waktu);
            const waktuWIB = dateObj.toLocaleTimeString('id-ID', {
                timeZone: 'Asia/Jakarta',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            }) + ' WIB';

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
        res.status(500).send("Kesalahan Database: " + err.message);
    }
});

app.get(['/scan', '/scanner'], (req, res) => {
    const userId = parseInt(req.query.userId) || 1;
    res.render('scan', { userId: userId });
});

// ---------------- API CRUD SISWA, GURU & KELAS ---------------- //

app.post('/api/kelas/tambah', async (req, res) => {
    const { nama_kelas } = req.body;
    try {
        await pool.query('INSERT INTO kelas (nama_kelas) VALUES ($1)', [nama_kelas]);
        res.redirect('back');
    } catch (err) {
        res.status(500).send("Gagal Menambahkan Rombel Kelas: " + err.message);
    }
});

app.post('/api/guru/tambah', async (req, res) => {
    const { nama, username, password, role, kelas_id } = req.body;
    try {
        const namaBersih = bersihkanGelar(nama);
        const kId = kelas_id ? parseInt(kelas_id) : null;
        await pool.query(
            'INSERT INTO users (nama, username, password, role, kelas_id) VALUES ($1, $2, $3, $4, $5)',
            [namaBersih, username, password, role || 'WALI_KELAS', kId]
        );
        res.redirect('back');
    } catch (err) {
        res.status(500).send("Gagal Menambahkan Pengguna Tenaga Pendidik: " + err.message);
    }
});

app.post('/api/guru/hapus/:id', async (req, res) => {
    try {
        await pool.query('DELETE FROM users WHERE id = $1', [req.params.id]);
        res.redirect('back');
    } catch (err) {
        res.status(500).send("Gagal Menghapus Akun Pengguna: " + err.message);
    }
});

app.post('/api/siswa/tambah', async (req, res) => {
    const { nama, nomor_wa_ortu, kelas_id } = req.body;
    try {
        if (!nama || !kelas_id) {
            return res.status(400).send("Nama Siswa dan Rombongan Belajar wajib diisi!");
        }

        let noWa = nomor_wa_ortu ? nomor_wa_ortu.trim().replace(/[^0-9]/g, '') : '';
        if (noWa.startsWith('0')) noWa = '62' + noWa.slice(1);

        await pool.query(
            'INSERT INTO siswa (nama, nomor_wa_ortu, kelas_id) VALUES ($1, $2, $3)',
            [nama.trim(), noWa, parseInt(kelas_id)]
        );

        res.redirect('back');
    } catch (err) {
        res.status(500).send("Gagal Menambahkan Data Siswa: " + err.message);
    }
});

app.post('/api/siswa/hapus/:id', async (req, res) => {
    try {
        await pool.query('DELETE FROM siswa WHERE id = $1', [req.params.id]);
        res.redirect('back');
    } catch (err) {
        res.status(500).send("Gagal Menghapus Data Siswa: " + err.message);
    }
});

// ---------------- API SCANNER & WA GATEWAY ---------------- //

app.get('/api/start-wa', (req, res) => {
    const userId = parseInt(req.query.userId) || 1;
    connectToWhatsApp(userId);
    res.json({ success: true, message: 'Menginisialisasi WhatsApp Gateway...' });
});

app.post('/api/scan', async (req, res) => {
    const { siswa_id, scanned_by } = req.body;
    if (!siswa_id) return res.status(400).json({ success: false, message: "Kode QR / Identitas Siswa tidak terdeteksi." });

    try {
        const parsedSiswaId = parseInt(siswa_id);
        const parsedScannedBy = parseInt(scanned_by) || 1;

        const siswaRes = await pool.query(`
            SELECT s.id, s.nama, s.nomor_wa_ortu, COALESCE(k.nama_kelas, '-') AS nama_kelas 
            FROM siswa s LEFT JOIN kelas k ON s.kelas_id = k.id WHERE s.id = $1
        `, [parsedSiswaId]);

        if (siswaRes.rows.length === 0) {
            return res.status(404).json({ success: false, message: `Kartu Identitas Siswa (ID: ${siswa_id}) Tidak Terdaftar!` });
        }

        const siswa = siswaRes.rows[0];

        await pool.query(
            `INSERT INTO absensi (siswa_id, status, scanned_by, waktu) VALUES ($1, 'HADIR', $2, NOW())`,
            [siswa.id, parsedScannedBy]
        );

        // Cari koneksi WA milik Guru yang scan, jika kosong fallback ke sesi WA aktif lain (misal Admin)
        let waClient = waSessions[parsedScannedBy];
        if (!waClient) {
            const activeUserIds = Object.keys(waSessions);
            if (activeUserIds.length > 0) {
                waClient = waSessions[activeUserIds[0]];
            }
        }

        let statusWA = "Notifikasi WhatsApp Tidak Terkirim (Layanan WA Belum Terkoneksi)";

        if (waClient && siswa.nomor_wa_ortu) {
            let phone = siswa.nomor_wa_ortu.toString().trim().replace(/[^0-9]/g, '');
            if (phone.startsWith('0')) phone = '62' + phone.slice(1);
            const formattedJid = phone + '@s.whatsapp.net';

            const now = new Date();
            const jamWib = now.toLocaleTimeString('id-ID', { timeZone: 'Asia/Jakarta', hour: '2-digit', minute: '2-digit' }) + ' WIB';
            const tglWib = now.toLocaleDateString('id-ID', { timeZone: 'Asia/Jakarta', weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

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

            waClient.sendMessage(formattedJid, { text: pesan }).catch(e => console.error("Gagal Mengirim Notifikasi WA:", e.message));
            statusWA = "Notifikasi WhatsApp Berhasil Dikirimkan ke Wali Murid ✅";
        }

        return res.json({
            success: true,
            message: statusWA,
            siswa: {
                id: siswa.id,
                nama: siswa.nama,
                nama_kelas: siswa.nama_kelas,
                waktu: new Date().toLocaleTimeString('id-ID', { timeZone: 'Asia/Jakarta', hour: '2-digit', minute: '2-digit' }) + ' WIB'
            }
        });

    } catch (err) {
        console.error("Kesalahan Scan:", err);
        return res.status(500).json({ success: false, message: "Terjadi Kendala Sistem: " + err.message });
    }
});

app.listen(PORT, () => console.log(`🚀 Server Sistem Presensi Terpadu Aktif di Port ${PORT}`));
