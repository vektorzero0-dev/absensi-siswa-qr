const { proto, initAuthCreds, BufferJSON, useMultiFileAuthState } = require('@whiskeysockets/baileys');
const path = require('path');
const fs = require('fs');
const pool = require('./db');

async function usePostgresAuthState(userId) {
    const authFolder = path.join(__dirname, 'auth_sessions', `user_${userId}`);
    
    if (!fs.existsSync(authFolder)) {
        fs.mkdirSync(authFolder, { recursive: true });
    }

    // 1. Restore sesi dari Neon DB ke folder lokal jika folder lokal kosong (setelah redeploy Render)
    try {
        const res = await pool.query("SELECT id, data FROM wa_sessions WHERE id LIKE $1", [`user_${userId}_%`]);
        if (res.rows.length > 0) {
            for (const row of res.rows) {
                const fileName = row.id.replace(`user_${userId}_`, '');
                const filePath = path.join(authFolder, fileName);
                if (!fs.existsSync(filePath)) {
                    fs.writeFileSync(filePath, row.data);
                }
            }
        }
    } catch (err) {
        console.error("⚠️ Gagal sync sesi dari DB:", err.message);
    }

    // 2. Gunakan auth folder lokal (Cepat & QR anti-gagal/kedap-kedip)
    const { state, saveCreds: saveLocalCreds } = await useMultiFileAuthState(authFolder);

    // 3. Backup otomatis file lokal ke Neon DB di latar belakang
    const backupToPostgres = async () => {
        try {
            const files = fs.readdirSync(authFolder);
            for (const file of files) {
                const filePath = path.join(authFolder, file);
                if (fs.lstatSync(filePath).isFile()) {
                    const content = fs.readFileSync(filePath, 'utf-8');
                    const key = `user_${userId}_${file}`;
                    await pool.query(
                        `INSERT INTO wa_sessions (id, data) VALUES ($1, $2)
                         ON CONFLICT (id) DO UPDATE SET data = EXCLUDED.data`,
                        [key, content]
                    );
                }
            }
        } catch (err) {
            console.error("⚠️ Gagal backup sesi ke DB:", err.message);
        }
    };

    return {
        state,
        saveCreds: async () => {
            await saveLocalCreds();
            backupToPostgres(); // Backup async tanpa mengganggu QR Code
        },
        clearCreds: async () => {
            if (fs.existsSync(authFolder)) {
                fs.rmSync(authFolder, { recursive: true, force: true });
            }
            try {
                await pool.query("DELETE FROM wa_sessions WHERE id LIKE $1", [`user_${userId}_%`]);
            } catch (err) {}
        }
    };
}

module.exports = usePostgresAuthState;
