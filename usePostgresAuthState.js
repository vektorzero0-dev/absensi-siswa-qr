const { BufferJSON, initAuthCreds, proto } = require('@whiskeysockets/baileys');
const pool = require('./db');

async function usePostgresAuthState(userId = 'default') {
    const readData = async (type, id) => {
        try {
            const key = `${userId}:${type}:${id}`;
            const res = await pool.query('SELECT session_data FROM wa_sessions WHERE key_id = $1', [key]);
            if (res.rows.length > 0) {
                return JSON.parse(res.rows[0].session_data, BufferJSON.reviver);
            }
            return null;
        } catch (error) {
            return null;
        }
    };

    const writeData = async (data, type, id) => {
        try {
            const key = `${userId}:${type}:${id}`;
            const value = JSON.stringify(data, BufferJSON.replacer);
            await pool.query(
                `INSERT INTO wa_sessions (key_id, session_data, updated_at) 
                 VALUES ($1, $2, NOW()) 
                 ON CONFLICT (key_id) DO UPDATE SET session_data = $2, updated_at = NOW()`,
                [key, value]
            );
        } catch (error) {
            console.error(`❌ Gagal simpan ${type}:${id} ke Neon:`, error.message);
        }
    };

    const removeData = async (type, id) => {
        try {
            const key = `${userId}:${type}:${id}`;
            await pool.query('DELETE FROM wa_sessions WHERE key_id = $1', [key]);
        } catch (error) {
            console.error(`❌ Gagal hapus ${type}:${id} dari Neon:`, error.message);
        }
    };

    const clearCreds = async () => {
        try {
            await pool.query('DELETE FROM wa_sessions WHERE key_id LIKE $1', [`${userId}:%`]);
            console.log(`🧹 Seluruh sesi DB User #${userId} berhasil dibersihkan.`);
        } catch (error) {
            console.error('❌ Gagal clearCreds:', error.message);
        }
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
                        Object.keys(ids).map(async (id) => {
                            let value = await readData(type, id);
                            if (type === 'app-state-sync-key' && value) {
                                value = proto.Message.AppStateSyncKeyData.fromObject(value);
                            }
                            data[id] = value;
                        })
                    );
                    return data;
                },
                // BAGIAN PENTING: Menyimpan rotasi kunci enkripsi harian ke Neon DB
                set: async (data) => {
                    const tasks = [];
                    for (const category in data) {
                        for (const id in data[category]) {
                            const value = data[category][id];
                            if (value) {
                                tasks.push(writeData(value, category, id));
                            } else {
                                tasks.push(removeData(category, id));
                            }
                        }
                    }
                    await Promise.all(tasks);
                }
            }
        },
        saveCreds: () => writeData(creds, 'creds', 'main'),
        clearCreds
    };
}

module.exports = usePostgresAuthState;
