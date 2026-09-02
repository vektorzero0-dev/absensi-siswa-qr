const { proto } = require('@whiskeysockets/baileys');
const { initAuthCreds, BufferJSON } = require('@whiskeysockets/baileys');
const pool = require('./db');

async function usePostgresAuthState(userId) {
    const writeData = async (data, id) => {
        const key = `user_${userId}_${id}`;
        const value = JSON.stringify(data, BufferJSON.replacer);
        await pool.query(
            `INSERT INTO wa_sessions (id, data) VALUES ($1, $2)
             ON CONFLICT (id) DO UPDATE SET data = EXCLUDED.data`,
            [key, value]
        );
    };

    const readData = async (id) => {
        try {
            const key = `user_${userId}_${id}`;
            const res = await pool.query('SELECT data FROM wa_sessions WHERE id = $1', [key]);
            if (res.rows.length > 0) {
                return JSON.parse(res.rows[0].data, BufferJSON.reviver);
            }
            return null;
        } catch (error) {
            return null;
        }
    };

    const removeData = async (id) => {
        try {
            const key = `user_${userId}_${id}`;
            await pool.query('DELETE FROM wa_sessions WHERE id = $1', [key]);
        } catch (error) {}
    };

    const creds = (await readData('creds')) || initAuthCreds();

    return {
        state: {
            creds,
            keys: {
                get: async (type, ids) => {
                    const data = {};
                    await Promise.all(
                        ids.map(async (id) => {
                            let value = await readData(`${type}-${id}`);
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
                            const key = `${category}-${id}`;
                            tasks.push(value ? writeData(value, key) : removeData(key));
                        }
                    }
                    await Promise.all(tasks);
                }
            }
        },
        saveCreds: () => writeData(creds, 'creds'),
        clearCreds: async () => {
            await pool.query("DELETE FROM wa_sessions WHERE id LIKE $1", [`user_${userId}_%`]);
        }
    };
}

module.exports = usePostgresAuthState;
