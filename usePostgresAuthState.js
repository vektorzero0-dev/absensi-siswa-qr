const { initAuthCreds, BufferJSON, useMultiFileAuthState } = require('@whiskeysockets/baileys');
const path = require('path');
const fs = require('fs');
const pool = require('./db');

async function usePostgresAuthState(userId) {
    const authFolder = path.join(__dirname, 'auth_sessions', `user_${userId}`);
    
    if (!fs.existsSync(authFolder)) {
        fs.mkdirSync(authFolder, { recursive: true });
    }

    // 1. Trik Utama: Saat Render Restart/Redeploy, restore SEMUA file sesi dari Neon DB ke folder lokal
    try {
        const res = await pool.query("SELECT id, data FROM wa_sessions WHERE id LIKE $1", [`user_${userId}_%`]);
        if (res.rows.length > 0) {
            console.log(`📦 [User #${userId}] Mengunduh ${res.rows.length} file sesi WA dari Neon DB...`);
            for (const row of res.rows) {
                const fileName = row.id.replace(`user_${userId}_`, '');
                const filePath = path.join(authFolder, fileName);
                // Tulis/Timpa file di lokal agar sesi selalu up-to-date
                fs.writeFileSync(filePath, row.data, 'utf-8');
            }
            console.log(`✅ [User #${userId}] Sesi WA berhasil dipulihkan dari Neon DB!`);
        }
    } catch (err) {
        console.error("⚠️ Gagal sync sesi dari DB Neon:", err.message);
    }

    // 2. Pakai multi-file auth state lokal (Cepat & Anti-kedap-kedip)
    const { state, saveCreds: saveLocalCreds } = await useMultiFileAuthState(authFolder);

    // 3. Simpan ke Lokal + LANGSUNG Simpan ke Neon DB secara Wajib (Await)
    const saveCredsDirect = async () => {
        await saveLocalCreds(); // Simpan lokal
        try {
            const files = fs.readdirSync(authFolder);
            for (const file of files) {
                const filePath = path.join(authFolder, file);
                if (fs.lstatSync(filePath).isFile()) {
                    const content = fs.readFileSync(filePath, 'utf-8');
                    const key = `user_${userId}_${file}`;
                    // Pakai await langsung agar tidak ada data tertinggal saat Render restart
                    await pool.query(
                        `INSERT INTO wa_sessions (id, data) VALUES ($1, $2)
                         ON CONFLICT (id) DO UPDATE SET data = EXCLUDED.data`,
                        [key, content]
                    );
                }
            }
        } catch (err) {
            console.error("⚠️ Gagal backup otomatis ke Neon DB:", err.message);
        }
    };

    return {
        state,
        saveCreds: saveCredsDirect,
        clearCreds: async () => {
            console.log(`🧹 Membersihkan seluruh sesi User #${userId}...`);
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
