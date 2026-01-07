/**
 * Google Apps Script สำหรับเชื่อมต่อกับระบบใบกำกับภาษี
 * 
 * วิธีใช้งาน:
 * 1. ไปที่ Google Sheets ของคุณ
 * 2. เลือก Extensions > Apps Script
 * 3. ลบโค้ดเดิมและ paste โค้ดนี้ทั้งหมด
 * 4. กด Save และตั้งชื่อ Project
 * 5. กด Deploy > New deployment
 * 6. เลือก Type: Web app
 * 7. Execute as: Me
 * 8. Who has access: Anyone
 * 9. กด Deploy และ copy URL ไปใส่ในระบบ
 */

// ===== Configuration =====
const CONFIG = {
    CUSTOMERS_SHEET: 'Customers',  // ชื่อ Sheet สำหรับข้อมูลลูกค้า
    INVOICES_SHEET: 'Invoices',    // ชื่อ Sheet สำหรับใบกำกับภาษี
    USERS_SHEET: 'Users'           // ชื่อ Sheet สำหรับข้อมูล User
};

/**
 * รับ request จากเว็บแอป
 */
function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);
        const action = data.action;
        const payload = data.data;

        let result;

        switch (action) {
            case 'addCustomer':
                result = addCustomer(payload);
                break;
            case 'updateCustomer':
                result = updateCustomer(payload);
                break;
            case 'addInvoice':
                result = addInvoice(payload);
                break;
            case 'getLatestInvoiceNumber':
                result = getLatestInvoiceNumber(payload);
                break;
            case 'updateInvoice':
                result = updateInvoice(payload);
                break;
            case 'sendInvoiceEmail':
                result = sendInvoiceEmail(payload);
                break;
            case 'login':
                result = authenticateUser(payload);
                break;
            default:
                result = { success: false, error: 'Unknown action' };
        }

        return ContentService
            .createTextOutput(JSON.stringify(result))
            .setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
        return ContentService
            .createTextOutput(JSON.stringify({ success: false, error: error.message }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}

/**
 * รองรับ GET request (สำหรับทดสอบ)
 */
function doGet(e) {
    return ContentService
        .createTextOutput(JSON.stringify({
            success: true,
            message: 'Bill Invoice System API is running'
        }))
        .setMimeType(ContentService.MimeType.JSON);
}

/**
 * เพิ่มลูกค้าใหม่
 */
function addCustomer(customer) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.CUSTOMERS_SHEET);

    // สร้าง Sheet ถ้ายังไม่มี
    if (!sheet) {
        sheet = ss.insertSheet(CONFIG.CUSTOMERS_SHEET);
        // เพิ่ม Header
        sheet.appendRow(['รหัสลูกค้า', 'ชื่อลูกค้า', 'ที่อยู่', 'เลขผู้เสียภาษี', 'เบอร์โทร', 'Email']);
        sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
        // ตั้งค่า Column D และ E เป็น Text Format เพื่อรักษาเลข 0 นำหน้า
        sheet.getRange('D:D').setNumberFormat('@');
        sheet.getRange('E:E').setNumberFormat('@');
    }

    // เพิ่มข้อมูลลูกค้า - ใส่ ' นำหน้า taxId และ phone เพื่อบังคับเป็น Text
    const newRow = sheet.getLastRow() + 1;
    const taxIdValue = customer.taxId ? "'" + String(customer.taxId) : '';
    const phoneValue = customer.phone ? "'" + String(customer.phone) : '';

    sheet.getRange(newRow, 1, 1, 6).setValues([[
        customer.id || '',
        customer.name || '',
        customer.address || '',
        taxIdValue,
        phoneValue,
        customer.email || ''
    ]]);

    return { success: true, message: 'เพิ่มลูกค้าสำเร็จ' };
}

/**
 * อัพเดทข้อมูลลูกค้า (ใช้เลขประจำตัวผู้เสียภาษีในการค้นหา)
 */
function updateCustomer(customer) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.CUSTOMERS_SHEET);

    if (!sheet) {
        return { success: false, error: 'ไม่พบ Sheet Customers' };
    }

    const data = sheet.getDataRange().getValues();
    const searchTaxId = String(customer.taxId || '').replace(/^'/, ''); // ลบ ' ออกถ้ามี

    // ค้นหาลูกค้าจากเลขประจำตัวผู้เสียภาษี (column D - index 3)
    for (let i = 1; i < data.length; i++) {
        // ลบ ' ออกจาก stored value ก่อนเปรียบเทียบ
        const storedTaxId = String(data[i][3] || '').replace(/^'/, '');

        if (storedTaxId === searchTaxId && searchTaxId !== '') {
            // อัพเดทแถวที่พบ - ใส่ ' นำหน้า taxId และ phone เพื่อบังคับเป็น Text
            const taxIdValue = customer.taxId ? "'" + String(customer.taxId).replace(/^'/, '') : '';
            const phoneValue = customer.phone ? "'" + String(customer.phone).replace(/^'/, '') : '';

            sheet.getRange(i + 1, 1, 1, 6).setValues([[
                customer.id || data[i][0] || '',
                customer.name || '',
                customer.address || '',
                taxIdValue,
                phoneValue,
                customer.email || data[i][5] || ''
            ]]);

            return { success: true, message: 'อัพเดทลูกค้าสำเร็จ' };
        }
    }

    // ถ้าไม่พบ ให้เพิ่มใหม่
    return addCustomer(customer);
}

/**
 * เพิ่มใบกำกับภาษี
 */
function addInvoice(invoice) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.INVOICES_SHEET);

    // สร้าง Sheet ถ้ายังไม่มี
    if (!sheet) {
        sheet = ss.insertSheet(CONFIG.INVOICES_SHEET);
        // เพิ่ม Header
        sheet.appendRow([
            'เลขที่ใบกำกับ', 'วันที่', 'ลูกค้า', 'ที่อยู่', 'เลขผู้เสียภาษี',
            'ราคารวม', 'VAT', 'รวมทั้งสิ้น', 'รายการ', 'สร้างเมื่อ'
        ]);
        sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    }

    // เพิ่มข้อมูลใบกำกับ - ใส่ ' นำหน้า customerTaxId เพื่อบังคับเป็น Text
    const taxIdValue = invoice.customerTaxId ? "'" + String(invoice.customerTaxId) : '';

    sheet.appendRow([
        invoice.invoiceNumber || '',
        invoice.date || '',
        invoice.customerName || '',
        invoice.customerAddress || '',
        taxIdValue,
        invoice.subtotal || 0,
        invoice.vat || 0,
        invoice.total || 0,
        invoice.items || '[]',
        new Date().toISOString()
    ]);

    return { success: true, message: 'บันทึกใบกำกับภาษีสำเร็จ' };
}

/**
 * ดึงเลขที่ใบกำกับล่าสุด
 * ใช้สำหรับ sync เลขใบกำกับอัตโนมัติ
 */
function getLatestInvoiceNumber(payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.INVOICES_SHEET);

    if (!sheet) {
        return { success: true, latestNumber: 0, datePrefix: '', message: 'ไม่พบ Sheet Invoices' };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
        return { success: true, latestNumber: 0, datePrefix: '', message: 'ยังไม่มีใบกำกับภาษี' };
    }

    // รับ datePrefix จาก payload (รูปแบบ YYMMDD เช่น 260107)
    // ถ้าไม่ส่งมา ใช้วันที่ปัจจุบัน
    let datePrefix = '';
    if (payload && payload.datePrefix) {
        datePrefix = String(payload.datePrefix);
    } else if (payload && payload.date) {
        // แปลง date เป็น YYMMDD
        const d = new Date(payload.date);
        const yy = d.getFullYear().toString().slice(-2);
        const mm = (d.getMonth() + 1).toString().padStart(2, '0');
        const dd = d.getDate().toString().padStart(2, '0');
        datePrefix = yy + mm + dd;
    } else {
        // ใช้วันที่ปัจจุบัน
        const d = new Date();
        const yy = d.getFullYear().toString().slice(-2);
        const mm = (d.getMonth() + 1).toString().padStart(2, '0');
        const dd = d.getDate().toString().padStart(2, '0');
        datePrefix = yy + mm + dd;
    }

    let maxRunningNumber = 0;

    // หาเลขรันสูงสุดสำหรับวันนั้น (column A = เลขที่ใบกำกับ)
    // รูปแบบ: YYMMDDXXXX เช่น 2601070001
    for (let i = 1; i < data.length; i++) {
        const invoiceNumber = String(data[i][0] || '').trim();

        // ตรวจสอบว่าขึ้นต้นด้วย datePrefix หรือไม่
        if (invoiceNumber.startsWith(datePrefix)) {
            // ดึงส่วนเลขรัน (4 หลักท้าย)
            const runningPart = invoiceNumber.substring(datePrefix.length);
            const num = parseInt(runningPart, 10);
            if (!isNaN(num) && num > maxRunningNumber) {
                maxRunningNumber = num;
            }
        }
    }

    Logger.log('Date prefix: ' + datePrefix + ', Max running: ' + maxRunningNumber);

    return {
        success: true,
        datePrefix: datePrefix,
        latestNumber: maxRunningNumber,
        nextNumber: maxRunningNumber + 1,
        nextInvoiceNumber: datePrefix + String(maxRunningNumber + 1).padStart(4, '0'),
        message: `เลขใบกำกับล่าสุดของวันที่ ${datePrefix}: ${maxRunningNumber}`
    };
}

/**
 * อัปเดตใบกำกับภาษี (ค้นหาจากเลขที่ใบกำกับแล้วอัปเดตข้อมูล)
 * รองรับการเปลี่ยนเลขที่ใบกำกับ โดยใช้ originalInvoiceNumber ในการค้นหา
 */
function updateInvoice(invoice) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.INVOICES_SHEET);

    if (!sheet) {
        return { success: false, error: 'ไม่พบ Sheet Invoices' };
    }

    const data = sheet.getDataRange().getValues();

    // ใช้ originalInvoiceNumber ในการค้นหาถ้ามี (กรณีเปลี่ยนเลขใบกำกับ)
    // ถ้าไม่มี ให้ใช้ invoiceNumber ปกติ - แปลงเป็น String เพื่อเปรียบเทียบ
    const searchNumber = String(invoice.originalInvoiceNumber || invoice.invoiceNumber).trim();

    Logger.log('Searching for invoice: ' + searchNumber);

    // ค้นหาใบกำกับจากเลขที่ใบกำกับ (column A - index 0)
    for (let i = 1; i < data.length; i++) {
        // แปลงเป็น String เพื่อเปรียบเทียบโดยไม่สนใจ type
        const rowInvoiceNumber = String(data[i][0] || '').trim();

        if (rowInvoiceNumber === searchNumber) {
            // พบแล้ว - อัปเดตข้อมูล
            const taxIdValue = invoice.customerTaxId ? "'" + String(invoice.customerTaxId) : '';

            sheet.getRange(i + 1, 1, 1, 10).setValues([[
                invoice.invoiceNumber || '',  // ใช้เลขใบกำกับใหม่
                invoice.date || '',
                invoice.customerName || '',
                invoice.customerAddress || '',
                taxIdValue,
                invoice.subtotal || 0,
                invoice.vat || 0,
                invoice.total || 0,
                invoice.items || '[]',
                new Date().toISOString()
            ]]);

            // Log การเปลี่ยนเลขใบกำกับ
            if (invoice.originalInvoiceNumber && invoice.originalInvoiceNumber !== invoice.invoiceNumber) {
                Logger.log(`Invoice number changed: ${invoice.originalInvoiceNumber} -> ${invoice.invoiceNumber}`);
            }

            return {
                success: true,
                message: 'อัปเดตใบกำกับภาษีสำเร็จ',
                oldNumber: invoice.originalInvoiceNumber,
                newNumber: invoice.invoiceNumber
            };
        }
    }

    // ถ้าไม่พบ ให้เพิ่มใหม่
    return addInvoice(invoice);
}

/**
 * ส่งใบกำกับภาษีทาง Email พร้อมแนบ PDF
 * @param {Object} payload - ข้อมูลสำหรับส่ง email
 * @param {string} payload.customerEmail - email ลูกค้า
 * @param {string} payload.customerName - ชื่อลูกค้า
 * @param {string} payload.invoiceNumber - เลขที่ใบกำกับ
 * @param {string} payload.invoiceHtml - HTML ของใบกำกับสำหรับสร้าง PDF
 * @param {string} payload.companyName - ชื่อบริษัทผู้ส่ง
 * @param {number} payload.total - ยอดรวม
 */
function sendInvoiceEmail(payload) {
    try {
        const {
            customerEmail,
            customerName,
            invoiceNumber,
            invoiceHtml,
            companyName,
            total
        } = payload;

        // ตรวจสอบข้อมูลที่จำเป็น
        if (!customerEmail) {
            return { success: false, error: 'กรุณาระบุ email ลูกค้า' };
        }

        if (!invoiceHtml) {
            return { success: false, error: 'ไม่พบข้อมูลใบกำกับภาษี' };
        }

        // สร้าง PDF จาก HTML
        const pdfBlob = createPdfFromHtml(invoiceHtml, invoiceNumber);

        // ชื่อไฟล์ PDF
        const pdfFileName = `ใบกำกับภาษี ${invoiceNumber}.pdf`;
        pdfBlob.setName(pdfFileName);

        // หัวข้อ email
        const subject = `ส่งใบกำกับภาษี ${invoiceNumber} ${companyName || 'ร้านวันดีดี คาเฟ่ เรสเตอร์รองต์'}`;

        // เนื้อหา email (HTML)
        const htmlBody = `
            <div style="font-family: 'Prompt', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
                <h2 style="color: #2c5282; border-bottom: 2px solid #4299e1; padding-bottom: 10px;">
                    📄 ใบกำกับภาษี ${invoiceNumber}
                </h2>
                
                <p style="font-size: 16px; color: #333; line-height: 1.8;">
                    เรียน คุณ${customerName || 'ลูกค้า'},
                </p>
                
                <p style="font-size: 16px; color: #333; line-height: 1.8;">
                    ${companyName || 'ร้านวันดีดี คาเฟ่ เรสเตอร์รองต์'} ขอส่งใบกำกับภาษี รบกวนตรวจสอบความถูกต้องของเอกสาร 
                    หากพบว่าไม่ถูกต้องสามารถติดต่อทางร้าน เพื่อดำเนินการแก้ไข 
                    และขอบคุณที่ท่านใช้บริการ หวังว่าคงจะได้รับใช้ท่านอีกในโอกาสต่อไป
                </p>

                <div style="background: #f7fafc; border-left: 4px solid #4299e1; padding: 15px; margin: 20px 0;">
                    <p style="margin: 0; font-size: 14px;">
                        <strong>เลขที่ใบกำกับ:</strong> ${invoiceNumber}<br>
                        <strong>ยอดรวม:</strong> ฿${Number(total || 0).toLocaleString('th-TH', { minimumFractionDigits: 2 })}
                    </p>
                </div>

                <p style="font-size: 14px; color: #666;">
                    📎 ไฟล์แนบ: ${pdfFileName}
                </p>

                <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 30px 0;">

                <p style="font-size: 14px; color: #666; text-align: center;">
                    ${companyName || 'ร้านวันดีดี คาเฟ่ เรสเตอร์รองต์'}<br>
                    <small>Email นี้ส่งโดยอัตโนมัติจากระบบใบกำกับภาษี</small>
                </p>
            </div>
        `;

        // เนื้อหา plain text (backup)
        const plainBody = `เรียน คุณ${customerName || 'ลูกค้า'},\n\n` +
            `${companyName || 'ร้านวันดีดี คาเฟ่ เรสเตอร์รองต์'} ขอส่งใบกำกับภาษี รบกวนตรวจสอบความถูกต้องของเอกสาร ` +
            `หากพบว่าไม่ถูกต้องสามารถติดต่อทางร้าน เพื่อดำเนินการแก้ไข ` +
            `และขอบคุณที่ท่านใช้บริการ หวังว่าคงจะได้รับใช้ท่านอีกในโอกาสต่อไป\n\n` +
            `เลขที่ใบกำกับ: ${invoiceNumber}\n` +
            `ยอดรวม: ฿${Number(total || 0).toLocaleString('th-TH', { minimumFractionDigits: 2 })}\n\n` +
            `ไฟล์แนบ: ${pdfFileName}\n\n` +
            `${companyName || 'ร้านวันดีดี คาเฟ่ เรสเตอร์รองต์'}`;

        // ส่ง email
        GmailApp.sendEmail(customerEmail, subject, plainBody, {
            htmlBody: htmlBody,
            attachments: [pdfBlob],
            name: companyName || 'ร้านวันดีดี คาเฟ่ เรสเตอร์รองต์'
        });

        return {
            success: true,
            message: `ส่ง email ไปยัง ${customerEmail} สำเร็จ`,
            sentTo: customerEmail,
            invoiceNumber: invoiceNumber
        };

    } catch (error) {
        Logger.log('Error sending email: ' + error.message);
        return {
            success: false,
            error: 'ไม่สามารถส่ง email ได้: ' + error.message
        };
    }
}

/**
 * สร้าง PDF จาก HTML
 * @param {string} html - HTML content
 * @param {string} invoiceNumber - เลขที่ใบกำกับ
 * @returns {Blob} PDF blob
 */
function createPdfFromHtml(html, invoiceNumber) {
    // เพิ่ม CSS สำหรับ PDF
    const fullHtml = `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600&display=swap');
                
                * {
                    margin: 0;
                    padding: 0;
                    box-sizing: border-box;
                }
                
                body {
                    font-family: 'Prompt', sans-serif;
                    font-size: 12px;
                    line-height: 1.4;
                    color: #333;
                }
                
                .invoice-container {
                    width: 210mm;
                    min-height: 297mm;
                    padding: 15mm;
                    background: white;
                }

                table {
                    width: 100%;
                    border-collapse: collapse;
                }

                th, td {
                    padding: 8px;
                    text-align: left;
                    border: 1px solid #ddd;
                }

                th {
                    background-color: #f5f5f5;
                }

                .text-right {
                    text-align: right;
                }

                .text-center {
                    text-align: center;
                }
            </style>
        </head>
        <body>
            ${html}
        </body>
        </html>
    `;

    // สร้าง PDF blob
    const blob = Utilities.newBlob(fullHtml, 'text/html', 'invoice.html');
    const pdf = blob.getAs('application/pdf');

    return pdf;
}

/**
 * ตรวจสอบ User Login
 * @param {Object} credentials - username และ password
 * @returns {Object} ผลลัพธ์การ login
 */
function authenticateUser(credentials) {
    try {
        const { username, password } = credentials;

        if (!username || !password) {
            return { success: false, error: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' };
        }

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName(CONFIG.USERS_SHEET);

        // สร้าง Sheet ถ้ายังไม่มี พร้อม default admin user
        if (!sheet) {
            sheet = ss.insertSheet(CONFIG.USERS_SHEET);
            // เพิ่ม Header
            sheet.appendRow(['Username', 'Password', 'ชื่อ-สกุล', 'Role', 'สถานะ', 'สร้างเมื่อ']);
            sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
            // เพิ่ม default admin user
            sheet.appendRow(['admin', 'admin123', 'ผู้ดูแลระบบ', 'admin', 'active', new Date().toISOString()]);

            // Return success with default admin
            return {
                success: true,
                user: {
                    username: 'admin',
                    name: 'ผู้ดูแลระบบ',
                    role: 'admin'
                },
                message: 'สร้าง Sheet Users และ default admin user สำเร็จ'
            };
        }

        const data = sheet.getDataRange().getValues();

        // ค้นหา user จาก username (column A - index 0)
        for (let i = 1; i < data.length; i++) {
            const storedUsername = String(data[i][0] || '').trim();
            const storedPassword = String(data[i][1] || '');
            const fullName = String(data[i][2] || '');
            const role = String(data[i][3] || 'user');
            const status = String(data[i][4] || 'active');

            if (storedUsername.toLowerCase() === username.toLowerCase()) {
                // พบ username - ตรวจสอบ password
                if (storedPassword === password) {
                    // ตรวจสอบสถานะ
                    if (status.toLowerCase() !== 'active') {
                        return { success: false, error: 'บัญชีถูกระงับการใช้งาน' };
                    }

                    return {
                        success: true,
                        user: {
                            username: storedUsername,
                            name: fullName,
                            role: role
                        },
                        message: 'เข้าสู่ระบบสำเร็จ'
                    };
                } else {
                    return { success: false, error: 'รหัสผ่านไม่ถูกต้อง' };
                }
            }
        }

        // ไม่พบ username
        return { success: false, error: 'ไม่พบชื่อผู้ใช้นี้ในระบบ' };

    } catch (error) {
        Logger.log('Authentication error: ' + error.message);
        return { success: false, error: 'เกิดข้อผิดพลาด: ' + error.message };
    }
}
