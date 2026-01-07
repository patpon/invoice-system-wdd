/**
 * Authentication Module
 * จัดการการ Login/Logout และ Session
 */

const Auth = {
    // Storage keys
    keys: {
        SESSION: 'bill_invoice_session',
        USER: 'bill_invoice_user'
    },

    /**
     * ตรวจสอบว่า user login แล้วหรือยัง
     * @returns {boolean}
     */
    isLoggedIn() {
        const session = this.loadSession();
        if (!session) return false;

        // ตรวจสอบว่า session หมดอายุหรือยัง (24 ชั่วโมง)
        const now = new Date().getTime();
        const sessionAge = now - session.loginTime;
        const maxAge = 24 * 60 * 60 * 1000; // 24 hours in milliseconds

        if (sessionAge > maxAge) {
            this.logout();
            return false;
        }

        return true;
    },

    /**
     * Login ด้วย username และ password
     * ใช้ Hardcoded credentials
     * @param {string} username 
     * @param {string} password 
     * @returns {Promise<{success: boolean, user?: object, error?: string}>}
     */
    async login(username, password) {
        // Hardcoded credentials
        const VALID_USERS = [
            { username: 'admin', password: 'admin123', name: 'ผู้ดูแลระบบ', role: 'admin' }
        ];

        try {
            const user = VALID_USERS.find(u =>
                u.username.toLowerCase() === username.toLowerCase() &&
                u.password === password
            );

            if (user) {
                this.saveSession({
                    user: {
                        username: user.username,
                        name: user.name,
                        role: user.role
                    },
                    loginTime: new Date().getTime()
                });

                return {
                    success: true,
                    user: {
                        username: user.username,
                        name: user.name,
                        role: user.role
                    }
                };
            } else {
                return {
                    success: false,
                    error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'
                };
            }
        } catch (error) {
            console.error('Login error:', error);
            return {
                success: false,
                error: 'เกิดข้อผิดพลาด'
            };
        }
    },

    /**
     * Logout และลบ session
     */
    logout() {
        localStorage.removeItem(this.keys.SESSION);
        localStorage.removeItem(this.keys.USER);
    },

    /**
     * ดึงข้อมูล user ที่ login อยู่
     * @returns {object|null}
     */
    getCurrentUser() {
        const session = this.loadSession();
        return session ? session.user : null;
    },

    /**
     * บันทึก session ลง localStorage
     * @param {object} sessionData 
     */
    saveSession(sessionData) {
        try {
            localStorage.setItem(this.keys.SESSION, JSON.stringify(sessionData));
        } catch (error) {
            console.error('Error saving session:', error);
        }
    },

    /**
     * โหลด session จาก localStorage
     * @returns {object|null}
     */
    loadSession() {
        try {
            const data = localStorage.getItem(this.keys.SESSION);
            return data ? JSON.parse(data) : null;
        } catch (error) {
            console.error('Error loading session:', error);
            return null;
        }
    },

    /**
     * ตรวจสอบ login และ redirect ไป login page ถ้ายังไม่ได้ login
     * ใช้ใน main app
     */
    requireLogin() {
        if (!this.isLoggedIn()) {
            window.location.href = 'login.html';
            return false;
        }
        return true;
    }
};

// Export for use
window.Auth = Auth;
