/**
 * Google Sheets API Module
 * เชื่อมต่อกับ Google Sheets สำหรับข้อมูลลูกค้าและใบกำกับภาษี
 */

const SheetsAPI = {
    /**
     * แปลง Google Sheets URL เป็น API URL
     */
    parseSheetUrl(url) {
        // Extract spreadsheet ID from URL
        const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        if (!match) return null;
        return match[1];
    },

    /**
     * ดึงข้อมูลจาก Google Sheets ด้วย JSONP (หลีกเลี่ยง CORS)
     */
    async fetchSheetData(sheetUrl, sheetName = 'Customers') {
        const spreadsheetId = this.parseSheetUrl(sheetUrl);
        if (!spreadsheetId) {
            throw new Error('URL ของ Google Sheet ไม่ถูกต้อง');
        }

        return new Promise((resolve, reject) => {
            // สร้าง callback function ชั่วคราว
            const callbackName = 'googleSheetCallback_' + Date.now();

            // ตั้งค่า timeout
            const timeout = setTimeout(() => {
                delete window[callbackName];
                if (script.parentNode) script.parentNode.removeChild(script);
                reject(new Error('การเชื่อมต่อหมดเวลา'));
            }, 10000);

            // สร้าง callback function
            window[callbackName] = (response) => {
                clearTimeout(timeout);
                delete window[callbackName];
                if (script.parentNode) script.parentNode.removeChild(script);

                if (response.status === 'error') {
                    reject(new Error(response.errors?.[0]?.message || 'เกิดข้อผิดพลาด'));
                    return;
                }

                try {
                    const data = this.parseTableData(response);
                    resolve(data);
                } catch (error) {
                    reject(error);
                }
            };

            // สร้าง script tag สำหรับ JSONP
            const script = document.createElement('script');
            const apiUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=responseHandler:${callbackName}&sheet=${encodeURIComponent(sheetName)}`;
            script.src = apiUrl;
            script.onerror = () => {
                clearTimeout(timeout);
                delete window[callbackName];
                reject(new Error('ไม่สามารถเชื่อมต่อ Google Sheets ได้'));
            };

            document.head.appendChild(script);
        });
    },

    /**
     * แปลงข้อมูลจาก Google Visualization format
     */
    parseTableData(data) {
        if (!data.table || !data.table.rows) {
            return [];
        }

        // Get headers from first row or column labels
        const cols = data.table.cols || [];
        const headers = cols.map((col, idx) => col.label || col.id || `col${idx}`);

        console.log('Headers found:', headers);

        const rows = data.table.rows.map(row => {
            const obj = {};
            if (row.c) {
                row.c.forEach((cell, index) => {
                    const header = headers[index] || `col${index}`;
                    // Handle both value formats
                    // ใช้ formatted value (f) สำหรับ taxId และ phone เพื่อรักษาเลข 0 นำหน้า
                    if (cell) {
                        // ถ้ามี formatted value (f) ให้ใช้ก่อน (จะรักษา 0 นำหน้า)
                        // ถ้าไม่มีค่อยใช้ raw value (v)
                        if (cell.f !== null && cell.f !== undefined && cell.f !== '') {
                            obj[header] = String(cell.f);
                        } else if (cell.v !== null && cell.v !== undefined) {
                            obj[header] = cell.v;
                        } else {
                            obj[header] = '';
                        }
                    } else {
                        obj[header] = '';
                    }
                });
            }
            return obj;
        });

        console.log('Parsed rows:', rows.length);
        return rows;
    },

    /**
     * ดึงข้อมูลลูกค้าจาก Google Sheets
     */
    async fetchCustomers() {
        const settings = Storage.getSettings();
        if (!settings.sheetsUrl) {
            throw new Error('กรุณาตั้งค่า Google Sheets URL ก่อน');
        }

        const data = await this.fetchSheetData(settings.sheetsUrl, 'Customers');

        console.log('Raw data from sheet:', data);

        // Map to customer format - support multiple column name formats
        const customers = data.map((row, index) => {
            // Try different possible column names/keys
            const keys = Object.keys(row);
            console.log('Row keys:', keys);

            return {
                id: row['รหัสลูกค้า'] || row['A'] || row['col0'] || row[keys[0]] || `CUST-${index + 1}`,
                name: row['ชื่อลูกค้า'] || row['B'] || row['col1'] || row[keys[1]] || '',
                address: row['ที่อยู่'] || row['C'] || row['col2'] || row[keys[2]] || '',
                taxId: row['เลขผู้เสียภาษี'] || row['D'] || row['col3'] || row[keys[3]] || '',
                phone: row['เบอร์โทร'] || row['E'] || row['col4'] || row[keys[4]] || '',
                email: row['Email'] || row['F'] || row['col5'] || row[keys[5]] || ''
            };
        }).filter(c => c.name); // กรองเฉพาะที่มีชื่อ

        console.log('Parsed customers:', customers);

        // Cache ข้อมูล
        Storage.saveCustomers(customers);

        return customers;
    },

    /**
     * ดึงประวัติใบกำกับภาษีจาก Google Sheets
     */
    async fetchInvoices() {
        const settings = Storage.getSettings();
        if (!settings.sheetsUrl) {
            return Storage.getInvoices(); // Return local data
        }

        try {
            const data = await this.fetchSheetData(settings.sheetsUrl, 'Invoices');

            console.log('Raw invoice data from sheet:', data);

            const invoices = data.map(row => {
                const keys = Object.keys(row);
                console.log('Invoice row keys:', keys, 'values:', row);

                // ลองดึงเลขใบกำกับจากหลายแหล่ง
                const invoiceNumber = row['เลขที่ใบกำกับ'] || row['Invoice No'] ||
                    row['A'] || row['col0'] || row[keys[0]] || '';

                // Helper function to parse number (remove comma separators)
                const parseNumber = (val) => {
                    if (val === null || val === undefined || val === '') return 0;
                    // Convert to string and remove commas
                    const cleanVal = String(val).replace(/,/g, '');
                    const num = parseFloat(cleanVal);
                    return isNaN(num) ? 0 : num;
                };

                // ดึง total จากหลายแหล่ง (column H = index 7 = รวมทั้งสิ้น)
                const subtotalValue = row['ราคารวม'] || row['F'] || row['col5'] || row[keys[5]] || 0;
                const vatValue = row['VAT'] || row['G'] || row['col6'] || row[keys[6]] || 0;
                const totalValue = row['รวมทั้งสิ้น'] || row['Total'] || row['H'] || row['col7'] || row[keys[7]] || 0;

                // ดึงข้อมูล items (อาจเก็บเป็น JSON string)
                let items = [];
                const itemsRaw = row['รายการ'] || row['items'] || row['I'] || row['col8'] || row[keys[8]] || '';
                if (itemsRaw) {
                    try {
                        items = typeof itemsRaw === 'string' ? JSON.parse(itemsRaw) : itemsRaw;
                    } catch (e) {
                        items = [];
                    }
                }

                // ดึงข้อมูล payment (อาจเก็บเป็น JSON string)
                let payment = { cash: false, cashAmount: 0, transfer: false, transferAmount: 0 };
                const paymentRaw = row['payment'] || row['การชำระเงิน'] || row['J'] || row['col9'] || row[keys[9]] || '';
                if (paymentRaw) {
                    try {
                        payment = typeof paymentRaw === 'string' ? JSON.parse(paymentRaw) : paymentRaw;
                    } catch (e) {
                        payment = { cash: false, cashAmount: 0, transfer: false, transferAmount: 0 };
                    }
                }

                // ดึงข้อมูล branch - ต้อง validate ว่า branchType เป็น 'hq' หรือ 'branch' เท่านั้น
                let branchTypeRaw = row['branchType'] || row['ประเภทสาขา'] || '';
                // ถ้าไม่ได้ระบุ branchType หรือไม่ใช่ค่าที่ถูกต้อง ให้ default เป็น 'hq'
                const branchType = (branchTypeRaw === 'hq' || branchTypeRaw === 'branch') ? branchTypeRaw : 'hq';
                const branchNumber = branchType === 'branch' ? (row['branchNumber'] || row['เลขสาขา'] || '') : '';

                return {
                    invoiceNumber: invoiceNumber,
                    date: row['วันที่'] || row['Date'] || row['B'] || row['col1'] || row[keys[1]] || '',
                    customerName: row['ลูกค้า'] || row['Customer'] || row['C'] || row['col2'] || row[keys[2]] || '',
                    customerAddress: row['ที่อยู่'] || row['D'] || row['col3'] || row[keys[3]] || '',
                    customerTaxId: row['เลขผู้เสียภาษี'] || row['E'] || row['col4'] || row[keys[4]] || '',
                    subtotal: parseNumber(subtotalValue),
                    vat: parseNumber(vatValue),
                    total: parseNumber(totalValue),
                    items: items,
                    payment: payment,
                    branchType: branchType,
                    branchNumber: branchNumber
                };
            }).filter(inv => inv.invoiceNumber);

            console.log('Parsed invoices:', invoices);

            return invoices;
        } catch (error) {
            console.warn('Cannot fetch invoices from Sheets, using local data:', error);
            return Storage.getInvoices();
        }
    },

    /**
     * ดึงเลขใบกำกับล่าสุดสำหรับวันที่ที่กำหนด
     * @param {string} date - วันที่ในรูปแบบ YYYY-MM-DD หรือ YYMMDD
     * @returns {Object} { success, datePrefix, latestNumber, nextNumber, nextInvoiceNumber }
     */
    async getLatestInvoiceNumber(date) {
        const settings = Storage.getSettings();
        if (!settings.scriptUrl) {
            console.warn('No script URL, using local counter');
            return { success: false, error: 'No script URL' };
        }

        try {
            const response = await fetch(settings.scriptUrl, {
                method: 'POST',
                redirect: 'follow',
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8'
                },
                body: JSON.stringify({
                    action: 'getLatestInvoiceNumber',
                    data: { date: date }
                })
            });

            try {
                const result = await response.json();
                console.log('Latest invoice number from Sheets:', result);
                return result;
            } catch (e) {
                // CORS issue - try no-cors mode
                console.warn('Cannot parse response, trying no-cors');
                return { success: false, error: 'Cannot parse response' };
            }
        } catch (error) {
            console.error('Error fetching latest invoice number:', error);
            return { success: false, error: error.message };
        }
    },

    /**
     * บันทึกลูกค้าใหม่ไปยัง Google Sheets (ผ่าน Apps Script)
     */
    async saveCustomer(customer, isUpdate = false) {
        const settings = Storage.getSettings();
        if (!settings.scriptUrl) {
            throw new Error('กรุณาตั้งค่า Google Apps Script URL ก่อน');
        }

        // ตรวจสอบว่าลูกค้ามีอยู่แล้วหรือไม่ (ใช้เลขประจำตัวผู้เสียภาษีเป็นตัวเช็ค)
        const customers = Storage.getCustomers();
        const existingIndex = customers.findIndex(c => c.taxId && customer.taxId && c.taxId === customer.taxId);
        const customerExists = existingIndex >= 0;

        // ใช้ action ที่ถูกต้องตามสถานะ
        const action = (isUpdate || customerExists) ? 'updateCustomer' : 'addCustomer';

        try {
            // บังคับให้ taxId และ phone เป็น String เพื่อรักษาเลข 0 นำหน้า
            const customerData = {
                ...customer,
                taxId: String(customer.taxId || ''),
                phone: String(customer.phone || '')
            };

            const response = await fetch(settings.scriptUrl, {
                method: 'POST',
                mode: 'no-cors',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    action: action,
                    data: customerData
                })
            });

            // Update local cache
            if (customerExists) {
                customers[existingIndex] = customer;
            } else {
                customers.push(customer);
            }
            Storage.saveCustomers(customers);

            return { success: true };
        } catch (error) {
            console.error('Error saving customer:', error);
            // Still save to local
            if (customerExists) {
                customers[existingIndex] = customer;
            } else {
                customers.push(customer);
            }
            Storage.saveCustomers(customers);
            throw error;
        }
    },

    /**
     * อัพเดทข้อมูลลูกค้า
     */
    async updateCustomer(customer) {
        const settings = Storage.getSettings();

        // Update local cache first
        const customers = Storage.getCustomers();
        const index = customers.findIndex(c => c.id === customer.id);
        if (index >= 0) {
            customers[index] = customer;
            Storage.saveCustomers(customers);
        }

        if (!settings.scriptUrl) {
            return { success: true, local: true };
        }

        try {
            // บังคับให้ taxId และ phone เป็น String เพื่อรักษาเลข 0 นำหน้า
            const customerData = {
                ...customer,
                taxId: String(customer.taxId || ''),
                phone: String(customer.phone || '')
            };

            await fetch(settings.scriptUrl, {
                method: 'POST',
                mode: 'no-cors',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    action: 'updateCustomer',
                    data: customerData
                })
            });

            return { success: true };
        } catch (error) {
            console.error('Error updating customer to Sheets:', error);
            return { success: true, local: true };
        }
    },

    /**
     * บันทึกใบกำกับภาษีไปยัง Google Sheets
     */
    async saveInvoice(invoice) {
        // Save to local storage first
        Storage.addInvoice(invoice);

        const settings = Storage.getSettings();
        if (!settings.scriptUrl) {
            return { success: true, local: true };
        }

        try {
            await fetch(settings.scriptUrl, {
                method: 'POST',
                mode: 'no-cors',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    action: 'addInvoice',
                    data: {
                        ...invoice,
                        items: JSON.stringify(invoice.items)
                    }
                })
            });

            return { success: true };
        } catch (error) {
            console.error('Error saving invoice to Sheets:', error);
            return { success: true, local: true };
        }
    },

    /**
     * อัปเดตใบกำกับภาษีใน Google Sheets
     */
    async updateInvoice(invoice) {
        const settings = Storage.getSettings();
        if (!settings.scriptUrl) {
            return { success: true, local: true };
        }

        try {
            await fetch(settings.scriptUrl, {
                method: 'POST',
                mode: 'no-cors',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    action: 'updateInvoice',
                    data: {
                        ...invoice,
                        customerTaxId: String(invoice.customerTaxId || '')
                    }
                })
            });

            return { success: true };
        } catch (error) {
            console.error('Error updating invoice to Sheets:', error);
            return { success: true, local: true };
        }
    },

    /**
     * ทดสอบการเชื่อมต่อ
     */
    async testConnection() {
        const settings = Storage.getSettings();

        const result = {
            sheetsOk: false,
            scriptOk: false,
            message: ''
        };

        // Test Sheets connection
        if (settings.sheetsUrl) {
            try {
                await this.fetchSheetData(settings.sheetsUrl, 'Customers');
                result.sheetsOk = true;
            } catch (error) {
                result.message += 'ไม่สามารถเชื่อมต่อ Google Sheets ได้\n';
            }
        }

        // Test Script connection (no-cors doesn't give response)
        if (settings.scriptUrl) {
            result.scriptOk = true; // Assume OK since we can't verify with no-cors
        }

        return result;
    },

    /**
     * ส่งใบกำกับภาษีทาง Email พร้อมแนบ PDF
     * @param {Object} emailData - ข้อมูลสำหรับส่ง email
     * @param {string} emailData.customerEmail - email ลูกค้า
     * @param {string} emailData.customerName - ชื่อลูกค้า
     * @param {string} emailData.invoiceNumber - เลขที่ใบกำกับ
     * @param {string} emailData.invoiceHtml - HTML สำหรับสร้าง PDF
     * @param {string} emailData.companyName - ชื่อบริษัท
     * @param {number} emailData.total - ยอดรวม
     */
    async sendInvoiceEmail(emailData) {
        const settings = Storage.getSettings();
        if (!settings.scriptUrl) {
            throw new Error('กรุณาตั้งค่า Google Apps Script URL ก่อน');
        }

        if (!emailData.customerEmail) {
            throw new Error('กรุณาระบุ email ลูกค้า');
        }

        try {
            // ใช้ fetch แบบ cors เพื่อรับ response
            const response = await fetch(settings.scriptUrl, {
                method: 'POST',
                redirect: 'follow',
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8'
                },
                body: JSON.stringify({
                    action: 'sendInvoiceEmail',
                    data: emailData
                })
            });

            // พยายามอ่าน response
            let result;
            try {
                result = await response.json();
            } catch (e) {
                // ถ้าไม่สามารถ parse JSON ได้ (เพราะ CORS) ถือว่าสำเร็จ
                return {
                    success: true,
                    message: `กำลังส่ง email ไปยัง ${emailData.customerEmail}...`
                };
            }

            return result;

        } catch (error) {
            console.error('Error sending invoice email:', error);

            // ถ้าเกิด error จาก CORS แต่ request ถูกส่งไปแล้ว
            // ลองส่งแบบ no-cors แทน
            try {
                await fetch(settings.scriptUrl, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        action: 'sendInvoiceEmail',
                        data: emailData
                    })
                });

                return {
                    success: true,
                    message: `กำลังส่ง email ไปยัง ${emailData.customerEmail}...`,
                    note: 'ไม่สามารถยืนยันผลได้ กรุณาตรวจสอบ email ปลายทาง'
                };
            } catch (e2) {
                throw new Error('ไม่สามารถส่ง email ได้: ' + error.message);
            }
        }
    }
};

// Export for use
window.SheetsAPI = SheetsAPI;
