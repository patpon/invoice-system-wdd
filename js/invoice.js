/**
 * Invoice Module
 * สร้างและจัดการใบกำกับภาษี
 */

const Invoice = {
    /**
     * แปลงตัวเลขเป็นคำอ่านภาษาไทย
     */
    numberToThaiText(number) {
        const thaiNumbers = ['', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
        const thaiUnits = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน', 'ล้าน'];

        if (number === 0) return 'ศูนย์';

        let result = '';
        let numStr = Math.floor(number).toString();
        let length = numStr.length;

        for (let i = 0; i < length; i++) {
            let digit = parseInt(numStr[i]);
            let position = length - i - 1;
            let unitIndex = position % 7;  // Reset after million

            if (digit === 0) continue;

            // Handle special cases
            if (unitIndex === 1 && digit === 1) {
                result += 'สิบ';
            } else if (unitIndex === 1 && digit === 2) {
                result += 'ยี่สิบ';
            } else if (unitIndex === 0 && digit === 1 && i !== 0) {
                result += 'เอ็ด';
            } else {
                result += thaiNumbers[digit] + thaiUnits[unitIndex];
            }

            // Add million separator
            if (position === 6) {
                result += 'ล้าน';
            }
        }

        return result;
    },

    /**
     * แปลงจำนวนเงินเป็นคำอ่านภาษาไทย (รวมสตางค์)
     */
    amountToThaiText(amount) {
        const baht = Math.floor(amount);
        const satang = Math.round((amount - baht) * 100);

        let result = '';

        if (baht === 0 && satang === 0) {
            return 'ศูนย์บาทถ้วน';
        }

        if (baht > 0) {
            result += this.numberToThaiText(baht) + 'บาท';
        }

        if (satang > 0) {
            result += this.numberToThaiText(satang) + 'สตางค์';
        } else {
            result += 'ถ้วน';
        }

        return result;
    },

    /**
     * คำนวณ VAT แบบ VAT Inclusive (ราคารวม VAT แล้ว)
     * สูตร: VAT = ยอดรวม * 7 / 107
     */
    calculateVATInclusive(total, vatRate = 7) {
        return total * vatRate / (100 + vatRate);
    },

    /**
     * คำนวณยอดรวมทั้งหมด (แบบ VAT Inclusive)
     * ราคาที่กรอกเป็นราคารวม VAT แล้ว
     */
    calculateTotal(items) {
        // ยอดรวม (รวม VAT แล้ว)
        const total = items.reduce((sum, item) => {
            return sum + (parseFloat(item.quantity) || 0) * (parseFloat(item.price) || 0);
        }, 0);

        const settings = Storage.getSettings();
        const vatRate = settings.vatRate || 7;

        // คำนวณ VAT ย้อนกลับจากยอดรวม
        // VAT = ยอดรวม * 7 / 107
        const vat = this.calculateVATInclusive(total, vatRate);

        // ราคาก่อน VAT = ยอดรวม - VAT
        const subtotal = total - vat;

        return {
            subtotal: Math.round(subtotal * 100) / 100,
            vat: Math.round(vat * 100) / 100,
            total: Math.round(total * 100) / 100
        };
    },

    /**
     * Format ตัวเลขเป็นสกุลเงิน
     */
    formatCurrency(amount) {
        return new Intl.NumberFormat('th-TH', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(amount);
    },

    /**
     * Format วันที่เป็นภาษาไทย
     */
    formatDate(dateString) {
        const date = new Date(dateString);
        const thaiMonths = [
            'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
            'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
        ];

        const day = date.getDate();
        const month = thaiMonths[date.getMonth()];
        const year = date.getFullYear() + 543; // Convert to Buddhist Era

        return `${day} ${month} ${year}`;
    },

    /**
     * Format วันที่แบบสั้น
     */
    formatDateShort(dateString) {
        const date = new Date(dateString);
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear() + 543;

        return `${day}/${month}/${year}`;
    },

    /**
     * สร้าง Preview HTML ของใบกำกับภาษี (แบบ A5)
     */
    generatePreviewHTML(data) {
        const { invoiceNumber, date, customer, items, subtotal, vat, total } = data;
        const branchType = data.branchType || 'hq';
        const branchNumber = data.branchNumber || '';
        const payment = data.payment || { cash: false, cashAmount: 0, transfer: false, transferAmount: 0 };

        const company = Storage.getCompany();
        const logo = Storage.getLogo();
        const signature = Storage.getSignature ? Storage.getSignature() : '';

        // Branch display
        const branchDisplay = branchType === 'hq'
            ? '<span class="info-value">☑ สำนักงานใหญ่</span>'
            : `<span class="info-value">☑ สาขาที่ ${branchNumber || '-'}</span>`;

        // Item amount should show amount before VAT (price / 1.07)
        const itemsHTML = items.map((item, index) => {
            const priceWithVat = (item.quantity || 0) * (item.price || 0);
            const priceBeforeVat = priceWithVat / 1.07;  // ราคาก่อน VAT
            return `
            <tr>
                <td class="text-center">${index + 1}</td>
                <td>${item.description || ''}</td>
                <td class="text-center">${item.quantity || 0}</td>
                <td class="text-right">${this.formatCurrency(priceBeforeVat)}</td>
            </tr>
        `}).join('');

        // Payment display
        const cashChecked = payment.cash ? '☑' : '☐';
        const transferChecked = payment.transfer ? '☑' : '☐';
        const cashAmountDisplay = payment.cash && payment.cashAmount > 0 ? this.formatCurrency(payment.cashAmount) : '________';
        const transferAmountDisplay = payment.transfer && payment.transferAmount > 0 ? this.formatCurrency(payment.transferAmount) : '________';

        return `
            <div class="invoice-a5">
                <!-- Header -->
                <div class="invoice-header">
                    <h1 class="invoice-title">ใบกำกับภาษี / ใบเสร็จรับเงิน (Tax Invoice / Receipt)</h1>
                </div>

                <!-- Company Info -->
                <div class="company-section">
                    <div class="company-logo">
                        ${logo ? `<img src="${logo}" alt="Logo">` : ''}
                    </div>
                    <div class="company-info">
                        <div class="company-name">${company.name || 'ชื่อบริษัท/ร้านค้า'}</div>
                        <div class="company-address">${company.address || 'ที่อยู่'}</div>
                        <div>โทร. ${company.phone || '-'}</div>
                        <div>Tax ID : ${company.taxId || '-'}</div>
                    </div>
                </div>

                <!-- Customer & Invoice Info -->
                <div class="info-section">
                    <div class="customer-box">
                        <div class="info-label">ชื่อผู้รับบริการ/Customer:</div>
                        <div class="info-value customer-name">${customer.name || '-'}</div>
                        <div class="info-label">ที่อยู่/Address:</div>
                        <div class="info-value">${customer.address || '-'}</div>
                        <div class="info-row">
                            <span class="info-label">เลขประจำตัวผู้เสียภาษี/Tax ID:</span>
                            <span class="info-value">${customer.taxId || '-'}</span>
                        </div>
                        <div class="info-row">
                            ${branchDisplay}
                        </div>
                    </div>
                    <div class="invoice-box">
                        <div class="info-row">
                            <span class="info-label">วันที่/Date:</span>
                            <span class="info-value date-value">${this.formatDateShort(date)}</span>
                        </div>
                        <div class="info-row">
                            <span class="info-label">เลขที่/No.:</span>
                            <span class="info-value invoice-no">${invoiceNumber}</span>
                        </div>
                    </div>
                </div>

                <!-- Items Table -->
                <table class="items-table">
                    <thead>
                        <tr>
                            <th class="col-no">ลำดับ<br>Items</th>
                            <th class="col-desc">รายละเอียด<br>Description</th>
                            <th class="col-qty">จำนวน/หน่วย<br>No of Trans.</th>
                            <th class="col-amount">จำนวนเงิน<br>Amount</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${itemsHTML || '<tr><td colspan="4" class="text-center">-</td></tr>'}
                        <tr class="empty-row"><td colspan="4">&nbsp;</td></tr>
                    </tbody>
                </table>

                <!-- Summary Section -->
                <div class="summary-section">
                    <div class="summary-left">
                        <div class="thai-text-row">
                            <span class="label">ตัวอักษร.</span>
                            <span class="value">( ${this.amountToThaiText(total)} )</span>
                        </div>
                        
                        <div class="payment-row">
                            <span class="checkbox">${cashChecked}</span> เงินสด
                            <span class="label">จำนวนเงิน</span>
                            <span class="amount-box">${cashAmountDisplay}</span>
                            <span>บาท</span>
                        </div>
                        
                        <div class="payment-row">
                            <span class="checkbox">${transferChecked}</span> เงินโอน
                            <span class="label">จำนวนเงิน</span>
                            <span class="amount-box">${transferAmountDisplay}</span>
                            <span>บาท</span>
                        </div>
                    </div>
                    
                    <div class="summary-right">
                        <div class="total-row">
                            <span>จำนวนเงินรวม<br>Total Amount</span>
                            <span class="amount">${this.formatCurrency(subtotal)}</span>
                        </div>
                        <div class="total-row">
                            <span>ภาษีมูลค่าเพิ่ม 7%<br>Vat</span>
                            <span class="amount">${this.formatCurrency(vat)}</span>
                        </div>
                        <div class="total-row grand">
                            <span>จำนวนเงินรวมทั้งสิ้น<br>Grand Total</span>
                            <span class="amount">${this.formatCurrency(total)}</span>
                        </div>
                    </div>
                </div>

                <!-- Signature Section - No Image in Preview, Only Line and Label -->
                <div class="signature-section">
                    <div class="signature-box">
                        <div class="signature-line"></div>
                        <div class="signature-label">ผู้รับเงิน</div>
                    </div>
                </div>
            </div>
        `;
    },

    /**
     * สร้าง HTML สำหรับพิมพ์ (A4) - รองรับต้นฉบับ + สำเนา
     */
    generatePrintHTML(data) {
        const company = Storage.getCompany();
        const logo = Storage.getLogo();
        const includeCopy = data.includeCopy !== false; // default true

        // Generate invoice content for original
        const originalContent = this.generateInvoiceContent(data, 'ต้นฉบับ');
        const copyContent = includeCopy ? this.generateInvoiceContent(data, 'สำเนา') : '';

        return `
            <!DOCTYPE html>
            <html lang="th">
            <head>
                <meta charset="UTF-8">
                <title>ใบกำกับภาษี ${data.invoiceNumber}</title>
                <link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600&display=swap" rel="stylesheet">
                <style>
                    * { margin: 0; padding: 0; box-sizing: border-box; }
                    @page { size: A4; margin: 0; }
                    body { 
                        font-family: 'Prompt', 'Sarabun', sans-serif; 
                        font-size: ${includeCopy ? '8pt' : '9pt'};
                        background: white;
                        color: #333;
                    }
                    .print-container {
                        width: 210mm;
                        height: 297mm;
                        display: flex;
                        flex-direction: column;
                    }
                    .invoice-wrapper {
                        height: 148.5mm;
                        padding: ${includeCopy ? '5mm' : '8mm'};
                        position: relative;
                        overflow: hidden;
                        border-bottom: 1px dashed #999;
                    }
                    .invoice-wrapper:last-child {
                        border-bottom: 1px dashed #999;
                    }
                    .copy-label {
                        position: absolute;
                        top: 5mm;
                        right: 5mm;
                        background: #1a5490;
                        color: white;
                        padding: 1px 8px;
                        border-radius: 3px;
                        font-size: ${includeCopy ? '8pt' : '9pt'};
                        font-weight: 600;
                    }
                    .invoice-a5 {
                        font-size: ${includeCopy ? '8pt' : '9pt'};
                    }
                    .invoice-header {
                        text-align: center;
                        border-bottom: 2px solid #1a5490;
                        padding-bottom: ${includeCopy ? '2px' : '5px'};
                        margin-bottom: ${includeCopy ? '4px' : '8px'};
                    }
                    .invoice-title {
                        font-size: ${includeCopy ? '12pt' : '14pt'};
                        font-weight: 600;
                        color: #1a5490;
                    }
                    .company-section {
                        display: flex;
                        gap: ${includeCopy ? '6px' : '10px'};
                        margin-bottom: ${includeCopy ? '4px' : '8px'};
                        padding: ${includeCopy ? '3px' : '5px'};
                        background: #f8f9fa;
                        border-radius: 3px;
                    }
                    .company-logo img {
                        max-width: ${includeCopy ? '35px' : '50px'};
                        max-height: ${includeCopy ? '35px' : '50px'};
                    }
                    .company-info {
                        flex: 1;
                        font-size: ${includeCopy ? '7pt' : '9pt'};
                    }
                    .company-name {
                        font-size: ${includeCopy ? '10pt' : '12pt'};
                        font-weight: 600;
                        color: #1a5490;
                    }
                    .info-section {
                        display: flex;
                        gap: ${includeCopy ? '5px' : '8px'};
                        margin-bottom: ${includeCopy ? '4px' : '8px'};
                    }
                    .customer-box, .invoice-box {
                        border: 1px solid #ddd;
                        padding: ${includeCopy ? '3px' : '5px'};
                        border-radius: 3px;
                    }
                    .customer-box { flex: 2; }
                    .invoice-box { flex: 1; }
                    .info-label {
                        color: #666;
                        font-size: ${includeCopy ? '7pt' : '8pt'};
                    }
                    .info-value { font-weight: 500; }
                    .customer-name {
                        font-size: ${includeCopy ? '10pt' : '11pt'};
                        color: #1a5490;
                        font-weight: 600;
                    }
                    .info-row {
                        display: flex;
                        gap: 5px;
                        align-items: center;
                        margin-top: 2px;
                    }
                    .date-value, .invoice-no {
                        background: #e8f4fd;
                        padding: 1px 6px;
                        border-radius: 3px;
                        font-weight: 600;
                    }
                    .items-table {
                        width: 100%;
                        border-collapse: collapse;
                        margin-bottom: ${includeCopy ? '4px' : '8px'};
                        font-size: ${includeCopy ? '8pt' : '9pt'};
                    }
                    .items-table th {
                        background: #1a5490;
                        color: white;
                        padding: ${includeCopy ? '3px' : '5px'};
                        text-align: center;
                        font-size: ${includeCopy ? '7pt' : '8pt'};
                    }
                    .items-table td {
                        border: 1px solid #ddd;
                        padding: ${includeCopy ? '2px 3px' : '4px'};
                    }
                    .col-no { width: 8%; }
                    .col-desc { width: 50%; }
                    .col-qty { width: 15%; }
                    .col-amount { width: 27%; }
                    .text-center { text-align: center; }
                    .text-right { text-align: right; }
                    .empty-row td { height: ${includeCopy ? '8px' : '15px'}; }
                    .summary-section {
                        display: flex;
                        gap: ${includeCopy ? '8px' : '15px'};
                        margin-bottom: ${includeCopy ? '4px' : '8px'};
                    }
                    .summary-left {
                        flex: 1;
                        font-size: ${includeCopy ? '7pt' : '9pt'};
                    }
                    .summary-right {
                        width: ${includeCopy ? '110px' : '140px'};
                    }
                    .thai-text-row { margin-bottom: ${includeCopy ? '2px' : '5px'}; }
                    .payment-row { margin: ${includeCopy ? '2px' : '3px'} 0; }
                    .checkbox { display: inline-block; width: 12px; }
                    .amount-box {
                        display: inline-block;
                        border-bottom: 1px solid #333;
                        min-width: ${includeCopy ? '55px' : '70px'};
                        text-align: center;
                        font-weight: 600;
                    }
                    .total-row {
                        display: flex;
                        justify-content: space-between;
                        padding: ${includeCopy ? '2px 5px' : '3px 6px'};
                        border: 1px solid #ddd;
                        font-size: ${includeCopy ? '8pt' : '9pt'};
                    }
                    .total-row.grand {
                        background: #1a5490;
                        color: white;
                        font-weight: 600;
                    }
                    .total-row .amount {
                        font-weight: 600;
                        min-width: ${includeCopy ? '55px' : '65px'};
                        text-align: right;
                    }
                    .signature-section {
                        text-align: center;
                        margin-top: ${includeCopy ? '6px' : '12px'};
                    }
                    .signature-box {
                        display: inline-block;
                        min-width: ${includeCopy ? '80px' : '110px'};
                        position: relative;
                    }
                    .signature-line {
                        border-bottom: 1px solid #333;
                        height: ${includeCopy ? '25px' : '40px'};
                        margin-bottom: 2px;
                    }
                    .signature-img {
                        max-width: ${includeCopy ? '55px' : '80px'};
                        max-height: ${includeCopy ? '22px' : '35px'};
                        position: absolute;
                        top: 0;
                        left: 50%;
                        transform: translateX(-50%);
                    }
                    .signature-label {
                        font-size: ${includeCopy ? '8pt' : '9pt'};
                        text-align: center;
                    }
                    @media print {
                        body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                    }
                </style>
            </head>
            <body>
                <div class="print-container">
                    <div class="invoice-wrapper">
                        <div class="copy-label">ต้นฉบับ</div>
                        ${originalContent}
                    </div>
                    ${includeCopy ? `
                    <div class="invoice-wrapper">
                        <div class="copy-label" style="background: #666;">สำเนา</div>
                        ${copyContent}
                    </div>
                    ` : ''}
                </div>
                <script>
                    window.onload = function() { 
                        window.print(); 
                        // ปิดหน้าพิมพ์หลังพิมพ์เสร็จหรือยกเลิก
                        window.onafterprint = function() {
                            window.close();
                        };
                        // สำหรับ browser ที่ไม่ support onafterprint
                        setTimeout(function() {
                            window.close();
                        }, 1000);
                    }
                </script>
            </body>
            </html>
        `;
    },

    /**
     * สร้างเนื้อหาใบกำกับภาษี (ใช้สำหรับทั้งต้นฉบับและสำเนา)
     */
    generateInvoiceContent(data, copyType) {
        const { invoiceNumber, date, customer, items, subtotal, vat, total } = data;
        const branchType = data.branchType || 'hq';
        const branchNumber = data.branchNumber || '';
        // ถ้า payment ไม่มีข้อมูล หรือ ไม่ได้เลือกช่องทางชำระเงินใดๆ ให้ default เป็นเงินสดเต็มจำนวน
        let payment = data.payment;
        if (!payment || (!payment.cash && !payment.transfer)) {
            payment = { cash: true, cashAmount: total || 0, transfer: false, transferAmount: 0 };
        }

        const company = Storage.getCompany();
        const logo = Storage.getLogo();
        const signature = Storage.getSignature ? Storage.getSignature() : '';

        const branchDisplay = branchType === 'hq'
            ? '<span class="info-value">☑ สำนักงานใหญ่</span>'
            : `<span class="info-value">☑ สาขาที่ ${branchNumber || '-'}</span>`;

        const itemsHTML = items.map((item, index) => {
            const priceWithVat = (item.quantity || 0) * (item.price || 0);
            const priceBeforeVat = priceWithVat / 1.07;
            return `
            <tr>
                <td class="text-center">${index + 1}</td>
                <td>${item.description || ''}</td>
                <td class="text-center">${item.quantity || 0}</td>
                <td class="text-right">${this.formatCurrency(priceBeforeVat)}</td>
            </tr>
        `}).join('');

        const cashChecked = payment.cash ? '☑' : '☐';
        const transferChecked = payment.transfer ? '☑' : '☐';
        const cashAmountDisplay = payment.cash && payment.cashAmount > 0 ? this.formatCurrency(payment.cashAmount) : '________';
        const transferAmountDisplay = payment.transfer && payment.transferAmount > 0 ? this.formatCurrency(payment.transferAmount) : '________';

        return `
            <div class="invoice-a5">
                <div class="invoice-header">
                    <h1 class="invoice-title">ใบกำกับภาษี / ใบเสร็จรับเงิน (Tax Invoice / Receipt)</h1>
                </div>
                <div class="company-section">
                    <div class="company-logo">
                        ${logo ? `<img src="${logo}" alt="Logo">` : ''}
                    </div>
                    <div class="company-info">
                        <div class="company-name">${company.name || 'ชื่อบริษัท/ร้านค้า'}</div>
                        <div>${company.address || 'ที่อยู่'}</div>
                        <div>โทร. ${company.phone || '-'}</div>
                        <div>Tax ID : ${company.taxId || '-'}</div>
                    </div>
                </div>
                <div class="info-section">
                    <div class="customer-box">
                        <div class="info-label">ชื่อผู้รับบริการ/Customer:</div>
                        <div class="info-value customer-name">${customer.name || '-'}</div>
                        <div class="info-label">ที่อยู่/Address:</div>
                        <div class="info-value">${customer.address || '-'}</div>
                        <div class="info-row">
                            <span class="info-label">Tax ID:</span>
                            <span class="info-value">${customer.taxId || '-'}</span>
                        </div>
                        <div class="info-row">${branchDisplay}</div>
                    </div>
                    <div class="invoice-box">
                        <div class="info-row">
                            <span class="info-label">วันที่:</span>
                            <span class="info-value date-value">${this.formatDateShort(date)}</span>
                        </div>
                        <div class="info-row">
                            <span class="info-label">เลขที่:</span>
                            <span class="info-value invoice-no">${invoiceNumber}</span>
                        </div>
                    </div>
                </div>
                <table class="items-table">
                    <thead>
                        <tr>
                            <th class="col-no">ลำดับ</th>
                            <th class="col-desc">รายละเอียด</th>
                            <th class="col-qty">จำนวน</th>
                            <th class="col-amount">จำนวนเงิน</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${itemsHTML || '<tr><td colspan="4" class="text-center">-</td></tr>'}
                    </tbody>
                </table>
                <div class="summary-section">
                    <div class="summary-left">
                        <div class="thai-text-row">
                            <span class="label">ตัวอักษร.</span>
                            <span class="value">( ${this.amountToThaiText(total)} )</span>
                        </div>
                        <div class="payment-row">
                            <span class="checkbox">${cashChecked}</span> เงินสด
                            <span class="amount-box">${cashAmountDisplay}</span> บาท
                        </div>
                        <div class="payment-row">
                            <span class="checkbox">${transferChecked}</span> เงินโอน
                            <span class="amount-box">${transferAmountDisplay}</span> บาท
                        </div>
                    </div>
                    <div class="summary-right">
                        <div class="total-row">
                            <span>จำนวนเงินรวม</span>
                            <span class="amount">${this.formatCurrency(subtotal)}</span>
                        </div>
                        <div class="total-row">
                            <span>VAT 7%</span>
                            <span class="amount">${this.formatCurrency(vat)}</span>
                        </div>
                        <div class="total-row grand">
                            <span>รวมทั้งสิ้น</span>
                            <span class="amount">${this.formatCurrency(total)}</span>
                        </div>
                    </div>
                </div>
                <div class="signature-section">
                    <div class="signature-box">
                        <div class="signature-line"></div>
                        <div class="signature-label">ผู้รับเงิน</div>
                    </div>
                </div>
            </div>
        `;
    },

    /**
     * พิมพ์ใบกำกับภาษี
     */
    print(data) {
        const printWindow = window.open('', '_blank');
        printWindow.document.write(this.generatePrintHTML(data));
        printWindow.document.close();
    },

    /**
     * แปลง URL รูปภาพเป็น Base64 (เพื่อให้ html2pdf render ได้)
     */
    async imageToBase64(url) {
        return new Promise((resolve) => {
            if (!url) {
                resolve('');
                return;
            }
            // ถ้าเป็น base64 อยู่แล้ว ให้ใช้เลย
            if (url.startsWith('data:')) {
                resolve(url);
                return;
            }
            const img = new Image();
            img.crossOrigin = 'anonymous';
            img.onload = function () {
                try {
                    const canvas = document.createElement('canvas');
                    canvas.width = img.width;
                    canvas.height = img.height;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0);
                    const dataURL = canvas.toDataURL('image/png');
                    resolve(dataURL);
                } catch (e) {
                    console.warn('Cannot convert image to base64:', e);
                    resolve(url);
                }
            };
            img.onerror = function () {
                console.warn('Cannot load image:', url);
                resolve('');
            };
            img.src = url;
        });
    },

    /**
     * สร้าง PDF และดาวน์โหลด (A4) - ใช้ inline styles เพื่อให้ html2pdf render ได้ถูกต้อง
     */
    async generatePDF(data) {
        const includeCopy = data.includeCopy === true;
        const company = Storage.getCompany();
        const logoPath = Storage.getLogo();
        const signaturePath = Storage.getSignature ? Storage.getSignature() : '';

        // แปลง logo และ signature เป็น base64 ก่อน เพื่อให้ html2pdf render ได้
        const logo = await this.imageToBase64(logoPath);
        const signature = await this.imageToBase64(signaturePath);

        // Generate invoice HTML with full inline styles
        const pdfContent = this.generatePDFInlineContent(data, company, logo, signature);
        const copyContent = includeCopy ? this.generatePDFInlineContent(data, company, logo, signature) : '';

        // Create container
        const element = document.createElement('div');
        element.innerHTML = `
            <div style="width: 794px; min-height: 1123px; background: white; font-family: 'Prompt', -apple-system, sans-serif; margin: 0; padding: 0; color: #333;">
                <div style="width: 794px; height: 561px; padding: 25px; position: relative; border-bottom: 2px dashed #999; box-sizing: border-box; overflow: hidden;">
                    <div style="position: absolute; top: 12px; right: 18px; background: #1a5490; color: white; padding: 3px 12px; border-radius: 4px; font-size: 11px; font-weight: 600;">ต้นฉบับ</div>
                    ${pdfContent}
                </div>
                ${includeCopy ? `
                <div style="width: 794px; height: 561px; padding: 25px; position: relative; border-bottom: 2px dashed #999; box-sizing: border-box; overflow: hidden;">
                    <div style="position: absolute; top: 12px; right: 18px; background: #666; color: white; padding: 3px 12px; border-radius: 4px; font-size: 11px; font-weight: 600;">สำเนา</div>
                    ${copyContent}
                </div>
                ` : `
                <div style="width: 794px; height: 561px; border-bottom: 2px dashed #999; box-sizing: border-box;"></div>
                `}
            </div>
        `;

        // Set element position for rendering
        element.style.cssText = 'position: absolute; left: -9999px; top: 0;';
        document.body.appendChild(element);

        const options = {
            margin: 0,
            filename: `ใบกำกับภาษี-${data.invoiceNumber}.pdf`,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: {
                scale: 2,
                useCORS: true,
                letterRendering: true,
                logging: false
            },
            jsPDF: {
                unit: 'mm',
                format: 'a4',
                orientation: 'portrait'
            }
        };

        try {
            await html2pdf().set(options).from(element).save();
        } catch (error) {
            console.error('Error generating PDF:', error);
            throw error;
        } finally {
            document.body.removeChild(element);
        }
    },

    /**
     * สร้างเนื้อหา PDF พร้อม inline styles ทั้งหมด
     */
    generatePDFInlineContent(data, company, logo, signature) {
        const { invoiceNumber, date, customer, items, subtotal, vat, total } = data;
        const branchType = data.branchType || 'hq';
        const branchNumber = data.branchNumber || '';
        // ถ้า payment ไม่มีข้อมูล หรือ ไม่ได้เลือกช่องทางชำระเงินใดๆ ให้ default เป็นเงินสดเต็มจำนวน
        let payment = data.payment;
        if (!payment || (!payment.cash && !payment.transfer)) {
            payment = { cash: true, cashAmount: total || 0, transfer: false, transferAmount: 0 };
        }

        const branchDisplay = branchType === 'hq' ? '☑ สำนักงานใหญ่' : `☑ สาขาที่ ${branchNumber || '-'}`;

        // Generate items rows with inline styles
        const itemsHTML = items.map((item, index) => {
            const priceWithVat = (item.quantity || 0) * (item.price || 0);
            const priceBeforeVat = priceWithVat / 1.07;
            return `
                <tr>
                    <td style="border: 1px solid #ddd; padding: 4px; text-align: center; font-size: 10px;">${index + 1}</td>
                    <td style="border: 1px solid #ddd; padding: 4px; font-size: 10px;">${item.description || ''}</td>
                    <td style="border: 1px solid #ddd; padding: 4px; text-align: center; font-size: 10px;">${item.quantity || 0}</td>
                    <td style="border: 1px solid #ddd; padding: 4px; text-align: right; font-size: 10px;">${this.formatCurrency(priceBeforeVat)}</td>
                </tr>
            `;
        }).join('');

        const cashChecked = payment.cash ? '☑' : '☐';
        const transferChecked = payment.transfer ? '☑' : '☐';
        const cashAmountDisplay = payment.cash && payment.cashAmount > 0 ? this.formatCurrency(payment.cashAmount) : '________';
        const transferAmountDisplay = payment.transfer && payment.transferAmount > 0 ? this.formatCurrency(payment.transferAmount) : '________';

        return `
            <div style="font-size: 10px;">
                <!-- Header -->
                <div style="text-align: center; border-bottom: 2px solid #1a5490; padding-bottom: 5px; margin-bottom: 8px;">
                    <h1 style="font-size: 14px; font-weight: 600; color: #1a5490; margin: 0;">ใบกำกับภาษี / ใบเสร็จรับเงิน (Tax Invoice / Receipt)</h1>
                </div>

                <!-- Company Section -->
                <div style="display: flex; gap: 10px; margin-bottom: 8px; padding: 6px; background: #f8f9fa; border-radius: 4px;">
                    ${logo ? `<div style="width: 45px; height: 45px; display: flex; align-items: center; justify-content: center;"><img src="${logo}" style="max-width: 45px; max-height: 45px; object-fit: contain;"></div>` : ''}
                    <div style="flex: 1; font-size: 9px;">
                        <div style="font-size: 12px; font-weight: 600; color: #1a5490;">${company.name || 'ชื่อบริษัท/ร้านค้า'}</div>
                        <div>${company.address || 'ที่อยู่'}</div>
                        <div>โทร. ${company.phone || '-'}</div>
                        <div>Tax ID : ${company.taxId || '-'}</div>
                    </div>
                </div>

                <!-- Customer & Invoice Info -->
                <div style="display: flex; gap: 8px; margin-bottom: 8px;">
                    <div style="flex: 2; border: 1px solid #ddd; padding: 6px; border-radius: 4px; font-size: 9px;">
                        <div style="color: #666; font-size: 8px;">ชื่อผู้รับบริการ/Customer:</div>
                        <div style="font-weight: 600; color: #1a5490; font-size: 11px;">${customer.name || '-'}</div>
                        <div style="color: #666; font-size: 8px;">ที่อยู่/Address:</div>
                        <div>${customer.address || '-'}</div>
                        <div style="margin-top: 3px;"><span style="color: #666;">Tax ID:</span> ${customer.taxId || '-'}</div>
                        <div>${branchDisplay}</div>
                    </div>
                    <div style="flex: 1; border: 1px solid #ddd; padding: 6px; border-radius: 4px; font-size: 9px;">
                        <div style="margin-bottom: 5px;">
                            <span style="color: #666;">วันที่:</span>
                            <span style="background: #e8f4fd; padding: 2px 6px; border-radius: 3px; font-weight: 600;">${this.formatDateShort(date)}</span>
                        </div>
                        <div>
                            <span style="color: #666;">เลขที่:</span>
                            <span style="background: #e8f4fd; padding: 2px 6px; border-radius: 3px; font-weight: 600;">${invoiceNumber}</span>
                        </div>
                    </div>
                </div>

                <!-- Items Table -->
                <table style="width: 100%; border-collapse: collapse; margin-bottom: 8px; font-size: 10px;">
                    <thead>
                        <tr style="background: #1a5490; color: white;">
                            <th style="padding: 5px; text-align: center; font-size: 9px; width: 8%;">ลำดับ</th>
                            <th style="padding: 5px; text-align: center; font-size: 9px; width: 50%;">รายละเอียด</th>
                            <th style="padding: 5px; text-align: center; font-size: 9px; width: 15%;">จำนวน</th>
                            <th style="padding: 5px; text-align: center; font-size: 9px; width: 27%;">จำนวนเงิน</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${itemsHTML || '<tr><td colspan="4" style="text-align: center; padding: 8px;">-</td></tr>'}
                    </tbody>
                </table>

                <!-- Summary Section -->
                <div style="display: flex; gap: 15px; margin-bottom: 8px;">
                    <div style="flex: 1; font-size: 9px;">
                        <div style="margin-bottom: 4px;"><span>ตัวอักษร.</span> ( ${this.amountToThaiText(total)} )</div>
                        <div style="margin: 3px 0;">${cashChecked} เงินสด <span style="border-bottom: 1px solid #333; min-width: 60px; display: inline-block; text-align: center; font-weight: 600;">${cashAmountDisplay}</span> บาท</div>
                        <div style="margin: 3px 0;">${transferChecked} เงินโอน <span style="border-bottom: 1px solid #333; min-width: 60px; display: inline-block; text-align: center; font-weight: 600;">${transferAmountDisplay}</span> บาท</div>
                    </div>
                    <div style="width: 140px;">
                        <div style="display: flex; justify-content: space-between; padding: 3px 6px; border: 1px solid #ddd; font-size: 9px;">
                            <span>จำนวนเงินรวม</span>
                            <span style="font-weight: 600;">${this.formatCurrency(subtotal)}</span>
                        </div>
                        <div style="display: flex; justify-content: space-between; padding: 3px 6px; border: 1px solid #ddd; border-top: none; font-size: 9px;">
                            <span>VAT 7%</span>
                            <span style="font-weight: 600;">${this.formatCurrency(vat)}</span>
                        </div>
                        <div style="display: flex; justify-content: space-between; padding: 3px 6px; background: #1a5490; color: white; font-weight: 600; font-size: 10px;">
                            <span>รวมทั้งสิ้น</span>
                            <span>${this.formatCurrency(total)}</span>
                        </div>
                    </div>
                </div>

                <!-- Signature -->
                <div style="text-align: center; margin-top: 15px;">
                    <div style="display: inline-block; min-width: 100px; position: relative;">
                        ${signature ? `<img src="${signature}" style="max-width: 70px; max-height: 30px; position: absolute; top: -5px; left: 50%; transform: translateX(-50%);">` : ''}
                        <div style="border-bottom: 1px solid #333; height: 30px; margin-bottom: 3px;"></div>
                        <div style="font-size: 9px;">ผู้รับเงิน</div>
                    </div>
                </div>
            </div>
        `;
    },

    /**
     * สร้าง HTML สำหรับ PDF ที่ส่ง Email (ใช้รูปแบบเดียวกับ Print/Preview)
     * รวม CSS styles ไว้ใน document เพื่อให้ Google Apps Script สร้าง PDF ได้ถูกต้อง
     */
    generateEmailPDFHTML(data, company, logo, signature) {
        const { invoiceNumber, date, customer, items, subtotal, vat, total } = data;
        // Validate branchType - ต้องเป็น 'hq' หรือ 'branch' เท่านั้น ถ้าไม่ใช่ให้ default เป็น 'hq'
        const branchTypeRaw = data.branchType || '';
        const branchType = (branchTypeRaw === 'hq' || branchTypeRaw === 'branch') ? branchTypeRaw : 'hq';
        const branchNumber = branchType === 'branch' ? (data.branchNumber || '') : '';

        // ถ้า payment ไม่มีข้อมูล หรือ ไม่ได้เลือกช่องทางชำระเงินใดๆ ให้ default เป็นเงินสดเต็มจำนวน
        let payment = data.payment;
        if (!payment || (!payment.cash && !payment.transfer)) {
            payment = { cash: true, cashAmount: total || 0, transfer: false, transferAmount: 0 };
        }

        const branchDisplay = branchType === 'hq' ? '☑ สำนักงานใหญ่' : `☑ สาขาที่ ${branchNumber || '-'}`;

        // Generate items rows
        const itemsHTML = items.map((item, index) => {
            const priceWithVat = (item.quantity || 0) * (item.price || 0);
            const priceBeforeVat = priceWithVat / 1.07;
            return `
                <tr>
                    <td class="text-center">${index + 1}</td>
                    <td>${item.description || ''}</td>
                    <td class="text-center">${item.quantity || 0}</td>
                    <td class="text-right">${this.formatCurrency(priceBeforeVat)}</td>
                </tr>
            `;
        }).join('');

        const cashChecked = payment.cash ? '☑' : '☐';
        const transferChecked = payment.transfer ? '☑' : '☐';
        const cashAmountDisplay = payment.cash && payment.cashAmount > 0 ? this.formatCurrency(payment.cashAmount) : '________';
        const transferAmountDisplay = payment.transfer && payment.transferAmount > 0 ? this.formatCurrency(payment.transferAmount) : '________';

        return `
            <!DOCTYPE html>
            <html lang="th">
            <head>
                <meta charset="UTF-8">
                <title>ใบกำกับภาษี ${invoiceNumber}</title>
                <link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600&display=swap" rel="stylesheet">
                <style>
                    * { margin: 0; padding: 0; box-sizing: border-box; }
                    @page { size: A5; margin: 5mm; }
                    body { 
                        font-family: 'Prompt', 'Sarabun', sans-serif; 
                        font-size: 9pt;
                        background: white;
                        color: #333;
                        padding: 10px;
                    }
                    .invoice-a5 {
                        font-size: 9pt;
                    }
                    .invoice-header {
                        text-align: center;
                        border-bottom: 2px solid #1a5490;
                        padding-bottom: 5px;
                        margin-bottom: 8px;
                    }
                    .invoice-title {
                        font-size: 14pt;
                        font-weight: 600;
                        color: #1a5490;
                    }
                    .company-section {
                        display: flex;
                        gap: 10px;
                        margin-bottom: 8px;
                        padding: 5px;
                        background: #f8f9fa;
                        border-radius: 3px;
                    }
                    .company-logo img {
                        max-width: 50px;
                        max-height: 50px;
                    }
                    .company-info {
                        flex: 1;
                        font-size: 9pt;
                    }
                    .company-name {
                        font-size: 12pt;
                        font-weight: 600;
                        color: #1a5490;
                    }
                    .info-section {
                        display: flex;
                        gap: 8px;
                        margin-bottom: 8px;
                    }
                    .customer-box, .invoice-box {
                        border: 1px solid #ddd;
                        padding: 5px;
                        border-radius: 3px;
                    }
                    .customer-box { flex: 2; }
                    .invoice-box { flex: 1; }
                    .info-label {
                        color: #666;
                        font-size: 8pt;
                    }
                    .info-value { font-weight: 500; }
                    .customer-name {
                        font-size: 11pt;
                        color: #1a5490;
                        font-weight: 600;
                    }
                    .info-row {
                        display: flex;
                        gap: 5px;
                        align-items: center;
                        margin-top: 2px;
                    }
                    .date-value, .invoice-no {
                        background: #e8f4fd;
                        padding: 1px 6px;
                        border-radius: 3px;
                        font-weight: 600;
                    }
                    .items-table {
                        width: 100%;
                        border-collapse: collapse;
                        margin-bottom: 8px;
                        font-size: 9pt;
                    }
                    .items-table th {
                        background: #1a5490;
                        color: white;
                        padding: 5px;
                        text-align: center;
                        font-size: 8pt;
                    }
                    .items-table td {
                        border: 1px solid #ddd;
                        padding: 4px;
                    }
                    .col-no { width: 8%; }
                    .col-desc { width: 50%; }
                    .col-qty { width: 15%; }
                    .col-amount { width: 27%; }
                    .text-center { text-align: center; }
                    .text-right { text-align: right; }
                    .summary-section {
                        display: flex;
                        gap: 15px;
                        margin-bottom: 8px;
                    }
                    .summary-left {
                        flex: 1;
                        font-size: 9pt;
                    }
                    .summary-right {
                        width: 140px;
                    }
                    .thai-text-row { margin-bottom: 5px; }
                    .payment-row { margin: 3px 0; }
                    .checkbox { display: inline-block; width: 12px; }
                    .amount-box {
                        display: inline-block;
                        border-bottom: 1px solid #333;
                        min-width: 70px;
                        text-align: center;
                        font-weight: 600;
                    }
                    .total-row {
                        display: flex;
                        justify-content: space-between;
                        padding: 3px 6px;
                        border: 1px solid #ddd;
                        font-size: 9pt;
                    }
                    .total-row.grand {
                        background: #333333;
                        color: white;
                        font-weight: 600;
                    }
                    .total-row .amount {
                        font-weight: 600;
                        min-width: 65px;
                        text-align: right;
                    }
                    .signature-section {
                        text-align: center;
                        margin-top: 15px;
                    }
                    .signature-box {
                        display: inline-flex;
                        flex-direction: column;
                        align-items: center;
                        min-width: 150px;
                    }
                    .signature-img {
                        max-width: 120px;
                        max-height: 50px;
                        margin-bottom: 5px;
                    }
                    .signature-line {
                        width: 150px;
                        border-bottom: 1px solid #333;
                        margin-bottom: 3px;
                    }
                    .signature-label {
                        font-size: 9pt;
                        text-align: center;
                    }
                </style>
            </head>
            <body>
                <div class="invoice-a5">
                    <div class="invoice-header">
                        <h1 class="invoice-title">ใบกำกับภาษี / ใบเสร็จรับเงิน (Tax Invoice / Receipt)</h1>
                    </div>
                    <div class="company-section">
                        <div class="company-logo">
                            ${logo ? `<img src="${logo}" alt="Logo">` : ''}
                        </div>
                        <div class="company-info">
                            <div class="company-name">${company.name || 'ชื่อบริษัท/ร้านค้า'}</div>
                            <div>${company.address || 'ที่อยู่'}</div>
                            <div>โทร. ${company.phone || '-'}</div>
                            <div>Tax ID : ${company.taxId || '-'}</div>
                        </div>
                    </div>
                    <div class="info-section">
                        <div class="customer-box">
                            <div class="info-label">ชื่อผู้รับบริการ/Customer:</div>
                            <div class="info-value customer-name">${customer.name || '-'}</div>
                            <div class="info-label">ที่อยู่/Address:</div>
                            <div class="info-value">${customer.address || '-'}</div>
                            <div class="info-row">
                                <span class="info-label">Tax ID:</span>
                                <span class="info-value">${customer.taxId || '-'}</span>
                            </div>
                            <div class="info-row">${branchDisplay}</div>
                        </div>
                        <div class="invoice-box">
                            <div class="info-row">
                                <span class="info-label">วันที่:</span>
                                <span class="info-value date-value">${this.formatDateShort(date)}</span>
                            </div>
                            <div class="info-row">
                                <span class="info-label">เลขที่:</span>
                                <span class="info-value invoice-no">${invoiceNumber}</span>
                            </div>
                        </div>
                    </div>
                    <table class="items-table">
                        <thead>
                            <tr>
                                <th class="col-no">ลำดับ</th>
                                <th class="col-desc">รายละเอียด</th>
                                <th class="col-qty">จำนวน</th>
                                <th class="col-amount">จำนวนเงิน</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${itemsHTML || '<tr><td colspan="4" class="text-center">-</td></tr>'}
                        </tbody>
                    </table>
                    <div class="summary-section">
                        <div class="summary-left">
                            <div class="thai-text-row">
                                <span class="label">ตัวอักษร.</span>
                                <span class="value">( ${this.amountToThaiText(total)} )</span>
                            </div>
                            <div class="payment-row">
                                <span class="checkbox">${cashChecked}</span> เงินสด
                                <span class="amount-box">${cashAmountDisplay}</span> บาท
                            </div>
                            <div class="payment-row">
                                <span class="checkbox">${transferChecked}</span> เงินโอน
                                <span class="amount-box">${transferAmountDisplay}</span> บาท
                            </div>
                        </div>
                        <div class="summary-right" style="width: 140px;">
                            <div style="display: flex; justify-content: space-between; padding: 3px 6px; border: 1px solid #ddd; font-size: 9pt;">
                                <span>จำนวนเงินรวม</span>
                                <span style="font-weight: 600; min-width: 65px; text-align: right;">${this.formatCurrency(subtotal)}</span>
                            </div>
                            <div style="display: flex; justify-content: space-between; padding: 3px 6px; border: 1px solid #ddd; border-top: none; font-size: 9pt;">
                                <span>VAT 7%</span>
                                <span style="font-weight: 600; min-width: 65px; text-align: right;">${this.formatCurrency(vat)}</span>
                            </div>
                            <div style="display: flex; justify-content: space-between; padding: 3px 6px; background-color: #333333; color: white; font-weight: 600; font-size: 9pt;">
                                <span>รวมทั้งสิ้น</span>
                                <span style="font-weight: 600; min-width: 65px; text-align: right;">${this.formatCurrency(total)}</span>
                            </div>
                        </div>
                    </div>
                    <div class="signature-section">
                        <div class="signature-box">
                            ${signature ? `<img class="signature-img" src="${signature}" alt="Signature">` : '<div style="height: 50px;"></div>'}
                            <div class="signature-line"></div>
                            <div class="signature-label">ผู้รับเงิน</div>
                        </div>
                    </div>
                </div>
            </body>
            </html>
        `;
    },

    /**
     * Styles สำหรับ PDF (ใช้รูปแบบเดียวกับการพิมพ์)
     */
    getPDFStyles(includeCopy = false) {
        return `
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { font-family: 'Prompt', sans-serif; background: white; }
            .print-container {
                width: 210mm;
                min-height: 297mm;
                background: white;
            }
            .invoice-wrapper {
                width: 210mm;
                height: 148.5mm;
                min-height: 148.5mm;
                max-height: 148.5mm;
                padding: ${includeCopy ? '5mm' : '8mm'};
                position: relative;
                overflow: hidden;
                border-bottom: 1px dashed #999;
                background: white;
                page-break-inside: avoid;
            }
            .invoice-wrapper:last-child {
                border-bottom: 1px dashed #999;
            }
            .empty-half {
                height: 148.5mm;
                min-height: 148.5mm;
            }
            .copy-label {
                position: absolute;
                top: 5mm;
                right: 5mm;
                background: #1a5490;
                color: white;
                padding: 2px 10px;
                border-radius: 3px;
                font-size: ${includeCopy ? '8pt' : '9pt'};
                font-weight: 600;
            }
            .copy-label-gray {
                background: #666;
            }
            .invoice-a5 { font-size: ${includeCopy ? '8pt' : '9pt'}; color: #333; }
            .invoice-header { text-align: center; border-bottom: 2px solid #1a5490; padding-bottom: ${includeCopy ? '2px' : '5px'}; margin-bottom: ${includeCopy ? '4px' : '8px'}; }
            .invoice-title { font-size: ${includeCopy ? '12pt' : '14pt'}; font-weight: 600; color: #1a5490; }
            .company-section { display: flex; gap: ${includeCopy ? '6px' : '10px'}; margin-bottom: ${includeCopy ? '4px' : '8px'}; padding: ${includeCopy ? '3px' : '5px'}; background: #f8f9fa; border-radius: 3px; }
            .company-logo img { max-width: ${includeCopy ? '35px' : '50px'}; max-height: ${includeCopy ? '35px' : '50px'}; }
            .company-info { flex: 1; font-size: ${includeCopy ? '7pt' : '9pt'}; }
            .company-name { font-size: ${includeCopy ? '10pt' : '12pt'}; font-weight: 600; color: #1a5490; }
            .info-section { display: flex; gap: ${includeCopy ? '5px' : '8px'}; margin-bottom: ${includeCopy ? '4px' : '8px'}; }
            .customer-box, .invoice-box { border: 1px solid #ddd; padding: ${includeCopy ? '3px' : '5px'}; border-radius: 3px; }
            .customer-box { flex: 2; }
            .invoice-box { flex: 1; }
            .info-label { color: #666; font-size: ${includeCopy ? '7pt' : '8pt'}; }
            .info-value { font-weight: 500; }
            .info-row { display: flex; gap: 5px; align-items: center; margin-top: 2px; }
            .customer-name { font-size: ${includeCopy ? '10pt' : '11pt'}; color: #1a5490; font-weight: 600; }
            .date-value, .invoice-no { background: #e8f4fd; padding: 1px 6px; border-radius: 3px; font-weight: 600; }
            .items-table { width: 100%; border-collapse: collapse; margin-bottom: ${includeCopy ? '4px' : '8px'}; font-size: ${includeCopy ? '8pt' : '9pt'}; }
            .items-table th { background: #1a5490; color: white; padding: ${includeCopy ? '3px' : '5px'}; text-align: center; font-size: ${includeCopy ? '7pt' : '8pt'}; }
            .items-table td { border: 1px solid #ddd; padding: ${includeCopy ? '2px 3px' : '4px'}; }
            .col-no { width: 8%; }
            .col-desc { width: 50%; }
            .col-qty { width: 15%; }
            .col-amount { width: 27%; }
            .text-center { text-align: center; }
            .text-right { text-align: right; }
            .empty-row td { height: ${includeCopy ? '8px' : '15px'}; }
            .summary-section { display: flex; gap: ${includeCopy ? '8px' : '15px'}; margin-bottom: ${includeCopy ? '4px' : '8px'}; }
            .summary-left { flex: 1; font-size: ${includeCopy ? '7pt' : '9pt'}; }
            .summary-right { width: ${includeCopy ? '110px' : '140px'}; }
            .thai-text-row { margin-bottom: ${includeCopy ? '2px' : '5px'}; }
            .payment-row { margin: ${includeCopy ? '2px' : '3px'} 0; }
            .checkbox { display: inline-block; width: 12px; }
            .amount-box { display: inline-block; border-bottom: 1px solid #333; min-width: ${includeCopy ? '55px' : '70px'}; text-align: center; font-weight: 600; }
            .total-row { display: flex; justify-content: space-between; padding: ${includeCopy ? '2px 5px' : '3px 6px'}; border: 1px solid #ddd; font-size: ${includeCopy ? '8pt' : '9pt'}; }
            .total-row.grand { background: #1a5490; color: white; font-weight: 600; }
            .total-row .amount { font-weight: 600; min-width: ${includeCopy ? '55px' : '65px'}; text-align: right; }
            .signature-section { text-align: center; margin-top: ${includeCopy ? '6px' : '12px'}; }
            .signature-box { display: inline-block; min-width: ${includeCopy ? '80px' : '110px'}; position: relative; }
            .signature-line { border-bottom: 1px solid #333; height: ${includeCopy ? '25px' : '40px'}; margin-bottom: 2px; }
            .signature-img { max-width: ${includeCopy ? '55px' : '80px'}; max-height: ${includeCopy ? '22px' : '35px'}; position: absolute; top: 0; left: 50%; transform: translateX(-50%); }
            .signature-label { font-size: ${includeCopy ? '8pt' : '9pt'}; text-align: center; }
        `;
    },

    /**
     * สร้างข้อมูลใบกำกับภาษีสำหรับบันทึก
     */
    createInvoiceData(formData) {
        const totals = this.calculateTotal(formData.items);

        return {
            invoiceNumber: formData.invoiceNumber,
            date: formData.date,
            customer: {
                name: formData.customerName,
                address: formData.customerAddress,
                taxId: formData.customerTaxId,
                phone: formData.customerPhone
            },
            items: formData.items,
            subtotal: totals.subtotal,
            vat: totals.vat,
            total: totals.total,
            thaiText: this.amountToThaiText(totals.total),
            createdAt: new Date().toISOString()
        };
    }
};

// Export for use
window.Invoice = Invoice;
