/**
 * Storage Module
 * จัดการการเก็บข้อมูลใน localStorage
 */

const Storage = {
    keys: {
        SETTINGS: 'bill_invoice_settings',
        COMPANY: 'bill_invoice_company',
        LOGO: 'bill_invoice_logo',
        CUSTOMERS: 'bill_invoice_customers',
        INVOICES: 'bill_invoice_invoices',
        INVOICE_COUNTER: 'bill_invoice_counter',
        INVOICE_DATE: 'bill_invoice_date'  // เก็บวันที่ล่าสุดสำหรับ reset เลขรัน
    },

    /**
     * บันทึกการตั้งค่าระบบ
     */
    saveSettings(settings) {
        localStorage.setItem(this.keys.SETTINGS, JSON.stringify(settings));
    },

    /**
     * ดึงการตั้งค่าระบบ
     */
    getSettings() {
        const data = localStorage.getItem(this.keys.SETTINGS);
        return data ? JSON.parse(data) : this.getDefaultSettings();
    },

    /**
     * การตั้งค่าเริ่มต้น
     */
    getDefaultSettings() {
        return {
            invoicePrefix: 'INV-WDD',
            invoiceStartNumber: 1,
            invoiceNumberPadding: 4,
            defaultCategory: 'ค่าอาหารและเครื่องดื่ม',
            vatRate: 7,
            sheetsUrl: 'https://docs.google.com/spreadsheets/d/13yguyRVJYM0XaI-HR2iDP-k5qMxRtQyQmiIJYG5XVuc/edit?usp=sharing',
            scriptUrl: 'https://script.google.com/macros/s/AKfycbzQpML7urblAd17vH-43JXG1QuvIEqYaMVvvmvMJrUzJVqXNpTFLL6cp2BBVnUpZp8qQg/exec'
        };
    },

    /**
     * บันทึกข้อมูลบริษัท
     */
    saveCompany(company) {
        localStorage.setItem(this.keys.COMPANY, JSON.stringify(company));
    },

    /**
     * ดึงข้อมูลบริษัท
     */
    getCompany() {
        const data = localStorage.getItem(this.keys.COMPANY);
        return data ? JSON.parse(data) : this.getDefaultCompany();
    },

    /**
     * ข้อมูลบริษัทเริ่มต้น
     */
    getDefaultCompany() {
        return {
            name: 'บริษัท ซิลเวอร์แลนด์ เอสเตท จำกัด (สำนักงานใหญ่)',
            address: '336,336/1,336/3-4 หมู่ที่ 3 ถนนศรีสมาน ตำบลบ้านใหม่ อำเภอปากเกร็ด จังหวัดนนทบุรี 11120',
            taxId: '0125562014407',
            phone: '098-293-5536'
        };
    },

    /**
     * บันทึกโลโก้ (Base64)
     */
    saveLogo(base64) {
        localStorage.setItem(this.keys.LOGO, base64);
    },


    /**
     * ดึงโลโก้
     */
    getLogo() {
        return localStorage.getItem(this.keys.LOGO) || '';
    },

    /**
     * โลโก้เริ่มต้น (ใช้ไฟล์ในโฟลเดอร์โปรเจค)
     */
    getDefaultLogo() {
        return 'logoWDD.jpg';
    },

    /**
     * โหลดโลโก้เริ่มต้นและแปลงเป็น Base64 แล้วเก็บใน localStorage
     */
    async initDefaultLogo() {
        // ถ้ามีโลโก้อยู่แล้วใน localStorage ไม่ต้องโหลดใหม่
        const existingLogo = localStorage.getItem(this.keys.LOGO);
        if (existingLogo) {
            return;
        }

        // โหลดและแปลงโลโก้เริ่มต้นเป็น base64
        try {
            const logoUrl = this.getDefaultLogo();
            const response = await fetch(logoUrl);
            const blob = await response.blob();

            return new Promise((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () => {
                    const base64 = reader.result;
                    this.saveLogo(base64);
                    console.log('Default logo loaded and saved to localStorage');
                    resolve();
                };
                reader.onerror = () => {
                    console.warn('Cannot read logo file');
                    resolve();
                };
                reader.readAsDataURL(blob);
            });
        } catch (error) {
            console.warn('Cannot load default logo:', error);
        }
    },

    /**
     * บันทึกลายเซ็น (Base64)
     */
    saveSignature(base64) {
        localStorage.setItem('bill_invoice_signature', base64);
    },

    /**
     * ดึงลายเซ็น
     */
    getSignature() {
        return localStorage.getItem('bill_invoice_signature') || '';
    },

    /**
     * ลบลายเซ็น
     */
    clearSignature() {
        localStorage.removeItem('bill_invoice_signature');
    },

    /**
     * บันทึกข้อมูลลูกค้า (Cache)
     */
    saveCustomers(customers) {
        localStorage.setItem(this.keys.CUSTOMERS, JSON.stringify(customers));
    },

    /**
     * ดึงข้อมูลลูกค้า
     */
    getCustomers() {
        const data = localStorage.getItem(this.keys.CUSTOMERS);
        return data ? JSON.parse(data) : [];
    },

    /**
     * บันทึกใบกำกับภาษี (Local)
     */
    saveInvoices(invoices) {
        localStorage.setItem(this.keys.INVOICES, JSON.stringify(invoices));
    },

    /**
     * ดึงใบกำกับภาษีทั้งหมด
     */
    getInvoices() {
        const data = localStorage.getItem(this.keys.INVOICES);
        return data ? JSON.parse(data) : [];
    },

    /**
     * เพิ่มใบกำกับภาษีใหม่
     */
    addInvoice(invoice) {
        const invoices = this.getInvoices();
        invoices.unshift(invoice);
        this.saveInvoices(invoices);
        return invoices;
    },

    /**
     * ดึงและเพิ่มเลขที่ใบกำกับ
     * รูปแบบ: YYMMDDXXXX (เช่น 2601070001)
     * - YY = ปี พ.ศ. 2 หลักสุดท้าย (2568 -> 68)
     * - MM = เดือน
     * - DD = วัน
     * - XXXX = เลขรัน 4 หลัก (เริ่มใหม่ทุกวัน)
     */
    getNextInvoiceNumber() {
        const today = new Date();
        const todayStr = this.formatDateForInvoice(today);

        // ตรวจสอบว่าเป็นวันใหม่หรือไม่
        const lastDate = localStorage.getItem(this.keys.INVOICE_DATE);

        let counter;
        if (lastDate !== todayStr) {
            // วันใหม่ - เริ่มนับ 1 ใหม่
            counter = 1;
            localStorage.setItem(this.keys.INVOICE_DATE, todayStr);
        } else {
            // วันเดิม - ใช้ counter ต่อจากเดิม
            counter = parseInt(localStorage.getItem(this.keys.INVOICE_COUNTER)) || 1;
        }

        // สร้างเลขใบกำกับ: YYMMDD + เลขรัน 4 หลัก
        const paddedNumber = counter.toString().padStart(4, '0');
        const invoiceNumber = todayStr + paddedNumber;

        // เพิ่ม counter
        localStorage.setItem(this.keys.INVOICE_COUNTER, counter + 1);

        return invoiceNumber;
    },

    /**
     * สร้าง YYMMDD จากวันที่
     * @param {Date} date
     * @returns {string} YYMMDD (ค.ศ.)
     */
    formatDateForInvoice(date) {
        // ใช้ปี ค.ศ. 2 หลักสุดท้าย (2026 -> 26)
        const yy = date.getFullYear().toString().slice(-2);
        const mm = (date.getMonth() + 1).toString().padStart(2, '0');
        const dd = date.getDate().toString().padStart(2, '0');
        return yy + mm + dd;
    },

    /**
     * รีเซ็ตเลขที่ใบกำกับ
     * จะอัพเดททั้ง counter และ date เพื่อให้ preview ถูกต้อง
     */
    resetInvoiceCounter(startNumber = 1) {
        // อัพเดท counter
        localStorage.setItem(this.keys.INVOICE_COUNTER, startNumber);
        // อัพเดทวันที่เป็นวันนี้ เพื่อให้ counter ถูกใช้งาน
        const today = new Date();
        const todayStr = this.formatDateForInvoice(today);
        localStorage.setItem(this.keys.INVOICE_DATE, todayStr);
    },

    /**
     * ดึงเลขที่ใบกำกับปัจจุบัน (preview)
     * รูปแบบ: YYMMDDXXXX
     */
    previewNextInvoiceNumber() {
        const today = new Date();
        const todayStr = this.formatDateForInvoice(today);

        // ตรวจสอบว่าเป็นวันใหม่หรือไม่
        const lastDate = localStorage.getItem(this.keys.INVOICE_DATE);

        let counter;
        if (lastDate !== todayStr) {
            counter = 1;
        } else {
            counter = parseInt(localStorage.getItem(this.keys.INVOICE_COUNTER)) || 1;
        }

        const paddedNumber = counter.toString().padStart(4, '0');
        return todayStr + paddedNumber;
    },

    /**
     * ล้างข้อมูลทั้งหมด
     */
    clearAll() {
        Object.values(this.keys).forEach(key => {
            localStorage.removeItem(key);
        });
    }
};

// Export for use
window.Storage = Storage;
