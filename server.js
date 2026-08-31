// ==========================================
// ROUTE PROCESS LOGIN (POST)
// ==========================================
app.post('/login', async (req, res) => {
    const { username, password } = req.body;

    try {
        const result = await pool.query(
            'SELECT * FROM users WHERE username = $1 AND password = $2',
            [username, password]
        );

        if (result.rows.length === 0) {
            // Jika username atau password salah
            return res.render('login', { error: 'Username atau password salah!' });
        }

        const user = result.rows[0];

        // Arahkan sesuai Role (ADMIN atau WALI_KELAS)
        if (user.role === 'ADMIN') {
            return res.redirect(`/admin?userId=${user.id}`);
        } else {
            return res.redirect(`/wali?userId=${user.id}`);
        }

    } catch (err) {
        console.error("Error pada saat login:", err.message);
        return res.render('login', { error: 'Terjadi kesalahan sistem saat login.' });
    }
});
