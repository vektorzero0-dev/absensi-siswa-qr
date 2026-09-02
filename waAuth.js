const { BufferJSON, initAuthCreds, proto } = require('@whiskeysockets/baileys');

async function useNeonAuthState(pool, userId = 'default') {
    // Fungsi membaca data sesi dari tabel wa_sessions di Neon Postgres
    const readData = async (type, id) => {
        try {
            const key = `${userId}:${type}:${id}`;
            const res = await pool.query('SELECT session_data FROM wa_sessions WHERE key_id = $1', [key]);
            if (res.rows.length > 0) {
                return JSON.parse(res.rows[0].session_data, BufferJSON.reviver);
            }
            return null;
        } catch (error) {
            console.error('Error membaca sesi dari Neon:', error);
            return null;
        }
    };

    // Fungsi menulis/memperbarui data sesi di tabel wa_sessions
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
            console.error('Error menyimpan sesi ke Neon DB:', error);
        }
    };

    // Fungsi menghapus sesi saat terputus/logout
    const removeData = async (type, id) => {
        try {
            const key = `${userId}:${type}:${id}`;
            await pool.query('DELETE FROM wa_sessions WHERE key_id = $1', [key]);
        } catch (error) {
            console.error('Error menghapus sesi dari Neon DB:', error);
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
        saveCreds: () => writeData(creds, 'creds', 'main')
    };
}

module.exports = useNeonAuthState;
