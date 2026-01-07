/**
 * Main Application
 * ระบบพิมพ์บิลและใบกำกับภาษี
 */

const App = {
    // State
    currentTab: 'invoice',
    items: [],
    selectedCustomer: null,
    customers: [],
    chart: null,

    /**
     * เริ่มต้นแอปพลิเคชัน
     */
    init() {
        // ตรวจสอบ login ก่อน
        if (!Auth.requireLogin()) {
            return; // Redirect ไป login page
        }

        // แสดงข้อมูล user ที่ login
        this.displayCurrentUser();

        // Setup logout button
        this.initLogout();

        this.loadSettings();
        this.loadCustomers();
        this.initEventListeners();
        this.initTabs();
        this.initInvoice();
        this.initLogo();
        this.initSignature();
        this.updatePreview();
        this.loadDashboard();
        console.log('App initialized');
    },

    /**
     * แสดงข้อมูล user ที่ login อยู่
     */
    displayCurrentUser() {
        const user = Auth.getCurrentUser();
        if (user) {
            const displayName = document.getElementById('userDisplayName');
            if (displayName) {
                displayName.textContent = `👤 ${user.name || user.username}`;
            }
        }
    },

    /**
     * เริ่มต้น Logout
     */
    initLogout() {
        const logoutBtn = document.getElementById('logoutBtn');
        if (logoutBtn) {
            logoutBtn.addEventListener('click', () => {
                if (confirm('ต้องการออกจากระบบหรือไม่?')) {
                    Auth.logout();
                    window.location.href = 'login.html';
                }
            });
        }
    },

    /**
     * โหลดการตั้งค่า
     */
    loadSettings() {
        const settings = Storage.getSettings();
        const company = Storage.getCompany();

        // Settings form
        document.getElementById('invoicePrefix').value = settings.invoicePrefix || 'INV-';
        document.getElementById('invoiceStartNumber').value = settings.invoiceStartNumber || 1;
        document.getElementById('invoiceNumberPadding').value = settings.invoiceNumberPadding || 4;
        document.getElementById('defaultCategory').value = settings.defaultCategory || 'ค่าอาหารและเครื่องดื่ม';
        document.getElementById('sheetsUrl').value = settings.sheetsUrl || '';
        document.getElementById('scriptUrl').value = settings.scriptUrl || '';

        // โหลดเลขรันปัจจุบัน
        const currentCounter = parseInt(localStorage.getItem('bill_invoice_counter')) || 1;
        document.getElementById('invoiceCurrentNumber').value = currentCounter;

        // Company form
        document.getElementById('companyName').value = company.name || '';
        document.getElementById('companyAddress').value = company.address || '';
        document.getElementById('companyTaxId').value = company.taxId || '';
        document.getElementById('companyPhone').value = company.phone || '';
    },

    /**
     * โหลดข้อมูลลูกค้า
     */
    async loadCustomers() {
        try {
            // Try to load from cache first
            this.customers = Storage.getCustomers();

            // Then try to fetch from Sheets
            const settings = Storage.getSettings();
            if (settings.sheetsUrl) {
                try {
                    this.customers = await SheetsAPI.fetchCustomers();

                    // Auto-sync invoice number from latest invoice
                    await this.syncInvoiceNumber();

                    this.showToast('โหลดข้อมูลลูกค้าสำเร็จ', 'success');
                } catch (error) {
                    console.warn('Could not fetch from Sheets, using cache');
                }
            }

            this.renderCustomersTable();
        } catch (error) {
            console.error('Error loading customers:', error);
        }
    },

    /**
     * Sync เลขที่ใบกำกับจาก Google Sheets
     * ดึงเลขล่าสุดสำหรับวันที่ที่เลือกจาก Sheets
     * @param {string} date - วันที่ในรูปแบบ YYYY-MM-DD (optional, default = today)
     */
    async syncInvoiceNumber(date) {
        try {
            // ใช้วันที่จาก form หรือวันที่ปัจจุบัน
            const invoiceDate = date || document.getElementById('invoiceDate')?.value || new Date().toISOString().split('T')[0];

            console.log('Syncing invoice number for date:', invoiceDate);

            // เรียก API ดึงเลขล่าสุดจาก Google Sheets
            const result = await SheetsAPI.getLatestInvoiceNumber(invoiceDate);

            if (result && result.success) {
                // อัพเดท counter ใน localStorage
                const nextNumber = result.nextNumber || 1;
                Storage.resetInvoiceCounter(nextNumber);

                // อัพเดท UI
                const invoiceNumberField = document.getElementById('invoiceNumber');
                if (invoiceNumberField) {
                    invoiceNumberField.value = result.nextInvoiceNumber || Storage.previewNextInvoiceNumber();
                }

                console.log(`Invoice number synced from Sheets: ${result.nextInvoiceNumber} (next running: ${nextNumber})`);
                return result;
            } else {
                // ถ้าไม่สำเร็จ ใช้ local counter
                console.warn('Could not sync from Sheets, using local counter');
                const invoiceNumberField = document.getElementById('invoiceNumber');
                if (invoiceNumberField) {
                    invoiceNumberField.value = Storage.previewNextInvoiceNumber();
                }
            }
        } catch (error) {
            console.warn('Could not sync invoice number:', error);
            // Fallback to local
            const invoiceNumberField = document.getElementById('invoiceNumber');
            if (invoiceNumberField) {
                invoiceNumberField.value = Storage.previewNextInvoiceNumber();
            }
        }
    },


    /**
     * เริ่มต้น Event Listeners
     */
    initEventListeners() {
        // Tabs
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => this.switchTab(e.target.dataset.tab));
        });

        // Header buttons
        document.getElementById('dashboardBtn').addEventListener('click', () => this.switchTab('dashboard'));
        document.getElementById('historyBtn').addEventListener('click', () => this.switchTab('history'));
        document.getElementById('settingsBtn').addEventListener('click', () => this.openModal('settingsModal'));

        // Logo upload
        document.getElementById('logoContainer').addEventListener('click', () => {
            document.getElementById('logoInput').click();
        });
        document.getElementById('logoInput').addEventListener('change', (e) => this.handleLogoUpload(e));
        document.getElementById('deleteLogoBtn').addEventListener('click', (e) => {
            e.stopPropagation(); // ป้องกันไม่ให้เปิด file dialog
            this.deleteLogo();
        });

        // Customer search
        document.getElementById('customerSearch').addEventListener('input', (e) => this.searchCustomers(e.target.value));
        document.getElementById('customerSearch').addEventListener('focus', () => {
            if (this.customers.length > 0) {
                this.showCustomerResults(this.customers.slice(0, 5));
            }
        });
        document.getElementById('newCustomerBtn').addEventListener('click', () => this.openCustomerModal());

        // Customer form change events
        ['customerName', 'customerAddress', 'customerTaxId', 'customerPhone'].forEach(id => {
            document.getElementById(id).addEventListener('input', () => this.updatePreview());
        });

        // Branch type toggle
        document.querySelectorAll('input[name="branchType"]').forEach(radio => {
            radio.addEventListener('change', (e) => {
                const branchNumberGroup = document.getElementById('branchNumberGroup');
                branchNumberGroup.style.display = e.target.value === 'branch' ? 'block' : 'none';
                this.updatePreview();
            });
        });
        document.getElementById('branchNumber').addEventListener('input', () => this.updatePreview());

        // Payment method checkboxes
        document.getElementById('paymentCash').addEventListener('change', (e) => {
            const amountInput = document.getElementById('paymentCashAmount');
            const transferCheck = document.getElementById('paymentTransfer');
            const transferInput = document.getElementById('paymentTransferAmount');
            amountInput.disabled = !e.target.checked;

            if (e.target.checked) {
                const totals = Invoice.calculateTotal(this.items);
                if (transferCheck.checked && parseFloat(transferInput.value) > 0) {
                    // If transfer already has amount, calculate remaining for cash
                    const remaining = totals.total - parseFloat(transferInput.value);
                    amountInput.value = remaining > 0 ? remaining.toFixed(2) : '0.00';
                } else {
                    amountInput.value = totals.total.toFixed(2);
                }
            } else {
                amountInput.value = '';
            }
            this.validatePayment();
            this.updatePreview();
        });
        document.getElementById('paymentTransfer').addEventListener('change', (e) => {
            const amountInput = document.getElementById('paymentTransferAmount');
            const cashCheck = document.getElementById('paymentCash');
            const cashInput = document.getElementById('paymentCashAmount');
            amountInput.disabled = !e.target.checked;

            if (e.target.checked) {
                const totals = Invoice.calculateTotal(this.items);
                if (cashCheck.checked && parseFloat(cashInput.value) > 0) {
                    // If cash already has amount, calculate remaining for transfer
                    const remaining = totals.total - parseFloat(cashInput.value);
                    amountInput.value = remaining > 0 ? remaining.toFixed(2) : '0.00';
                } else {
                    amountInput.value = totals.total.toFixed(2);
                }
            } else {
                amountInput.value = '';
            }
            this.validatePayment();
            this.updatePreview();
        });
        ['paymentCashAmount', 'paymentTransferAmount'].forEach(id => {
            document.getElementById(id).addEventListener('input', (e) => {
                // ถ้าแก้ไขยอดเงินโอน ให้คำนวณยอดเงินสดใหม่
                if (id === 'paymentTransferAmount') {
                    const totals = Invoice.calculateTotal(this.items);
                    const cashCheckbox = document.getElementById('paymentCash');
                    const cashAmountInput = document.getElementById('paymentCashAmount');
                    const transferAmount = parseFloat(e.target.value) || 0;

                    if (cashCheckbox && cashCheckbox.checked) {
                        const cashAmount = totals.total - transferAmount;
                        cashAmountInput.value = cashAmount >= 0 ? cashAmount.toFixed(2) : '0.00';
                    }
                }
                this.validatePayment();
                this.updatePreview();
            });
        });

        // Invoice date
        document.getElementById('invoiceDate').valueAsDate = new Date();
        document.getElementById('invoiceDate').addEventListener('change', () => this.updatePreview());

        // Add item button
        document.getElementById('addItemBtn').addEventListener('click', () => this.addItem());

        // Print, PDF & Save buttons
        document.getElementById('printBtn').addEventListener('click', () => this.printInvoice());
        document.getElementById('pdfBtn').addEventListener('click', () => this.exportPDF());
        document.getElementById('saveBtn').addEventListener('click', () => this.saveInvoice());
        document.getElementById('saveAndPrintBtn').addEventListener('click', () => this.saveAndPrintInvoice());

        // Email button
        document.getElementById('emailBtn').addEventListener('click', () => this.openEmailModal());

        // Settings modal
        document.getElementById('closeSettingsBtn').addEventListener('click', () => this.closeModal('settingsModal'));
        document.getElementById('cancelSettingsBtn').addEventListener('click', () => this.closeModal('settingsModal'));
        document.getElementById('saveSettingsBtn').addEventListener('click', () => this.saveSettings());
        document.getElementById('testConnectionBtn').addEventListener('click', () => this.testConnection());
        document.getElementById('syncDataBtn').addEventListener('click', () => this.syncData());
        document.getElementById('setInvoiceCounterBtn').addEventListener('click', () => this.setInvoiceCounter());

        // Signature upload
        document.getElementById('signatureUploadArea').addEventListener('click', () => {
            document.getElementById('signatureInput').click();
        });
        document.getElementById('signatureInput').addEventListener('change', (e) => this.handleSignatureUpload(e));
        document.getElementById('clearSignatureBtn').addEventListener('click', () => this.clearSignature());

        // Settings tabs
        document.querySelectorAll('.settings-tab').forEach(tab => {
            tab.addEventListener('click', (e) => this.switchSettingsTab(e.target.dataset.settings));
        });

        // Customer modal
        document.getElementById('addCustomerBtn').addEventListener('click', () => this.openCustomerModal());
        document.getElementById('closeCustomerBtn').addEventListener('click', () => this.closeModal('customerModal'));
        document.getElementById('cancelCustomerBtn').addEventListener('click', () => this.closeModal('customerModal'));
        document.getElementById('saveCustomerBtn').addEventListener('click', () => this.saveCustomer());

        // Customer list
        document.getElementById('customerListSearch').addEventListener('input', (e) => this.filterCustomersList(e.target.value));
        document.getElementById('refreshCustomersBtn').addEventListener('click', () => this.loadCustomers());

        // Invoice history
        document.getElementById('searchInvoiceBtn').addEventListener('click', () => this.searchInvoices());
        document.getElementById('invoiceSearch').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') this.searchInvoices();
        });
        document.getElementById('syncHistoryBtn').addEventListener('click', () => this.syncInvoiceHistory());
        document.getElementById('clearHistoryBtn').addEventListener('click', () => this.clearInvoiceHistory());

        // Export history buttons
        document.getElementById('exportHistoryPdfBtn').addEventListener('click', () => this.exportHistoryToPDF());
        document.getElementById('exportHistoryExcelBtn').addEventListener('click', () => this.exportHistoryToExcel());

        // Invoice detail modal
        document.getElementById('closeDetailBtn').addEventListener('click', () => this.closeModal('invoiceDetailModal'));
        document.getElementById('closeDetailModal').addEventListener('click', () => this.closeModal('invoiceDetailModal'));
        document.getElementById('reprintBtn').addEventListener('click', () => this.reprintInvoice());

        // Email modal
        document.getElementById('closeEmailBtn').addEventListener('click', () => this.closeModal('emailModal'));
        document.getElementById('cancelEmailBtn').addEventListener('click', () => this.closeModal('emailModal'));
        document.getElementById('confirmSendEmailBtn').addEventListener('click', () => this.sendInvoiceEmailFromHistory());
        document.getElementById('emailSignatureInput').addEventListener('change', (e) => this.handleEmailSignatureUpload(e));

        // Close modals on backdrop click
        document.querySelectorAll('.modal').forEach(modal => {
            modal.addEventListener('click', (e) => {
                if (e.target === modal) this.closeModal(modal.id);
            });
        });

        // Close customer results on outside click
        document.addEventListener('click', (e) => {
            const searchContainer = document.querySelector('.search-container');
            if (!searchContainer.contains(e.target)) {
                document.getElementById('customerResults').classList.remove('active');
            }
        });
    },

    /**
     * เริ่มต้น Tabs
     */
    initTabs() {
        this.switchTab('invoice');
    },

    /**
     * สลับ Tab
     */
    switchTab(tabName) {
        this.currentTab = tabName;

        // Update tab buttons
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tab === tabName);
        });

        // Update tab contents
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.toggle('active', content.id === tabName + 'Tab');
        });

        // Load tab-specific data
        if (tabName === 'dashboard') {
            this.loadDashboard();
        } else if (tabName === 'history') {
            this.loadHistory();
        }
    },

    /**
     * เริ่มต้นใบกำกับภาษี
     */
    initInvoice() {
        // Set invoice number
        document.getElementById('invoiceNumber').value = Storage.previewNextInvoiceNumber();

        // Add default item
        this.addItem();
    },

    /**
     * เริ่มต้นโลโก้
     */
    initLogo() {
        const logo = Storage.getLogo();
        if (logo) {
            document.getElementById('logoPreview').src = logo;
            document.getElementById('logoPreview').classList.remove('hidden');
            document.getElementById('logoPlaceholder').classList.add('hidden');
            document.getElementById('deleteLogoBtn').classList.remove('hidden');
        }
    },

    /**
     * ลบโลโก้
     */
    deleteLogo() {
        Storage.saveLogo(''); // Clear logo from storage
        document.getElementById('logoPreview').src = '';
        document.getElementById('logoPreview').classList.add('hidden');
        document.getElementById('logoPlaceholder').classList.remove('hidden');
        document.getElementById('deleteLogoBtn').classList.add('hidden');
        this.updatePreview();
        this.showToast('ลบโลโก้แล้ว', 'success');
    },

    /**
     * อัปโหลดโลโก้
     */
    handleLogoUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (event) => {
            const base64 = event.target.result;
            Storage.saveLogo(base64);

            document.getElementById('logoPreview').src = base64;
            document.getElementById('logoPreview').classList.remove('hidden');
            document.getElementById('logoPlaceholder').classList.add('hidden');
            document.getElementById('deleteLogoBtn').classList.remove('hidden');

            this.updatePreview();
            this.showToast('อัปโหลดโลโก้สำเร็จ', 'success');
        };
        reader.readAsDataURL(file);
    },

    /**
     * อัปโหลดลายเซ็น
     */
    handleSignatureUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (event) => {
            const base64 = event.target.result;
            Storage.saveSignature(base64);

            document.getElementById('signaturePreview').src = base64;
            document.getElementById('signaturePreview').classList.remove('hidden');
            document.getElementById('signaturePlaceholder').classList.add('hidden');

            this.updatePreview();
            this.showToast('อัปโหลดลายเซ็นสำเร็จ', 'success');
        };
        reader.readAsDataURL(file);
    },

    /**
     * ลบลายเซ็น
     */
    clearSignature() {
        Storage.clearSignature();
        document.getElementById('signaturePreview').src = '';
        document.getElementById('signaturePreview').classList.add('hidden');
        document.getElementById('signaturePlaceholder').classList.remove('hidden');
        this.updatePreview();
        this.showToast('ลบลายเซ็นแล้ว', 'success');
    },

    /**
     * เริ่มต้นลายเซ็น (โหลดจาก storage)
     */
    initSignature() {
        const signature = Storage.getSignature();
        if (signature) {
            document.getElementById('signaturePreview').src = signature;
            document.getElementById('signaturePreview').classList.remove('hidden');
            document.getElementById('signaturePlaceholder').classList.add('hidden');
        }
    },

    /**
     * ค้นหาลูกค้า
     */
    searchCustomers(query) {
        if (!query) {
            document.getElementById('customerResults').classList.remove('active');
            return;
        }

        const queryLower = query.toLowerCase().trim();

        const results = this.customers.filter(c => {
            const name = (c.name || '').toLowerCase();
            const id = String(c.id || '').toLowerCase();
            const taxId = String(c.taxId || '');

            return name.includes(queryLower) ||
                id.includes(queryLower) ||
                taxId.includes(query);  // เลขผู้เสียภาษีไม่ต้อง toLowerCase
        }).slice(0, 10);  // เพิ่มจำนวนผลลัพธ์

        this.showCustomerResults(results);
    },

    /**
     * แสดงผลการค้นหาลูกค้า
     */
    showCustomerResults(customers) {
        const container = document.getElementById('customerResults');

        if (customers.length === 0) {
            container.classList.remove('active');
            return;
        }

        container.innerHTML = customers.map(c => `
            <div class="search-result-item" data-customer-id="${String(c.id).replace(/"/g, '&quot;')}">
                <div class="search-result-name">${c.name || '-'}</div>
                <div class="search-result-info">${c.id} | เลขผู้เสียภาษี: ${c.taxId || '-'}</div>
            </div>
        `).join('');

        // ใช้ event delegation แทน
        container.onclick = (e) => {
            const item = e.target.closest('.search-result-item');
            if (item) {
                const customerId = item.dataset.customerId;
                this.selectCustomer(customerId);
            }
        };

        container.classList.add('active');
    },

    /**
     * เลือกลูกค้า
     */
    selectCustomer(customerId) {
        // รองรับทั้ง id เป็นตัวเลขและ string
        const customer = this.customers.find(c => String(c.id) === String(customerId));
        if (!customer) {
            console.error('Customer not found:', customerId);
            return;
        }

        this.selectedCustomer = customer;

        document.getElementById('customerSearch').value = customer.name || '';
        document.getElementById('customerName').value = customer.name || '';
        document.getElementById('customerAddress').value = customer.address || '';
        document.getElementById('customerTaxId').value = String(customer.taxId || '');
        document.getElementById('customerPhone').value = String(customer.phone || '');

        document.getElementById('customerResults').classList.remove('active');
        this.updatePreview();

        this.showToast(`เลือกลูกค้า: ${customer.name}`, 'success');
    },

    /**
     * เพิ่มรายการสินค้า
     */
    addItem() {
        const settings = Storage.getSettings();
        const item = {
            id: Date.now(),
            description: settings.defaultCategory || 'ค่าอาหารและเครื่องดื่ม',
            quantity: 1,
            price: 0
        };

        this.items.push(item);
        this.renderItems();
    },

    /**
     * ลบรายการสินค้า
     */
    removeItem(itemId) {
        this.items = this.items.filter(item => item.id !== itemId);
        this.renderItems();
        this.updateTotals();
        this.updatePreview();
    },

    /**
     * Render รายการสินค้า
     */
    renderItems() {
        const tbody = document.getElementById('itemsBody');

        tbody.innerHTML = this.items.map((item, index) => `
            <tr data-item-id="${item.id}">
                <td class="text-center">${index + 1}</td>
                <td class="item-desc">
                    <input type="text" value="${item.description}" placeholder="รายละเอียด" 
                           onchange="App.updateItem(${item.id}, 'description', this.value)">
                </td>
                <td class="item-qty">
                    <input type="number" value="${item.quantity}" min="0" step="1"
                           onchange="App.updateItem(${item.id}, 'quantity', this.value)">
                </td>
                <td class="item-price">
                    <input type="number" value="${item.price}" min="0" step="0.01"
                           onchange="App.updateItem(${item.id}, 'price', this.value)">
                </td>
                <td class="item-total text-right">
                    ${Invoice.formatCurrency(item.quantity * item.price)}
                </td>
                <td class="item-action">
                    <button class="btn-remove" onclick="App.removeItem(${item.id})">×</button>
                </td>
            </tr>
        `).join('');
    },

    /**
     * อัปเดตรายการสินค้า
     */
    updateItem(itemId, field, value) {
        const item = this.items.find(i => i.id === itemId);
        if (!item) return;

        if (field === 'quantity' || field === 'price') {
            item[field] = parseFloat(value) || 0;
        } else {
            item[field] = value;
        }

        this.renderItems();
        this.updateTotals();
        this.updatePreview();
    },

    /**
     * อัปเดตยอดรวม
     */
    updateTotals() {
        const totals = Invoice.calculateTotal(this.items);

        document.getElementById('subtotal').textContent = Invoice.formatCurrency(totals.subtotal) + ' บาท';
        document.getElementById('vatAmount').textContent = Invoice.formatCurrency(totals.vat) + ' บาท';
        document.getElementById('grandTotal').textContent = Invoice.formatCurrency(totals.total) + ' บาท';
        document.getElementById('totalText').textContent = '(' + Invoice.amountToThaiText(totals.total) + ')';

        // อัปเดตยอดเงินสดอัตโนมัติ
        const cashCheckbox = document.getElementById('paymentCash');
        const cashAmountInput = document.getElementById('paymentCashAmount');
        const transferCheckbox = document.getElementById('paymentTransfer');
        const transferAmountInput = document.getElementById('paymentTransferAmount');

        if (cashCheckbox && cashCheckbox.checked) {
            if (transferCheckbox && transferCheckbox.checked) {
                // ถ้าเลือกทั้งเงินสดและเงินโอน: เงินสด = รวมทั้งสิ้น - เงินโอน
                const transferAmount = parseFloat(transferAmountInput.value) || 0;
                cashAmountInput.value = (totals.total - transferAmount).toFixed(2);
            } else {
                // ถ้าเลือกเฉพาะเงินสด: เงินสด = รวมทั้งสิ้น
                cashAmountInput.value = totals.total.toFixed(2);
            }
        }
    },

    /**
     * ตรวจสอบยอดชำระเงิน (เงินสด + เงินโอน = ยอดรวม)
     */
    validatePayment() {
        const totals = Invoice.calculateTotal(this.items);
        const cashCheck = document.getElementById('paymentCash').checked;
        const transferCheck = document.getElementById('paymentTransfer').checked;
        const cashAmount = parseFloat(document.getElementById('paymentCashAmount').value) || 0;
        const transferAmount = parseFloat(document.getElementById('paymentTransferAmount').value) || 0;

        let warningEl = document.getElementById('paymentWarning');
        if (!warningEl) {
            warningEl = document.createElement('div');
            warningEl.id = 'paymentWarning';
            warningEl.className = 'payment-warning';
            const paymentSection = document.querySelector('.payment-options');
            if (paymentSection) {
                paymentSection.appendChild(warningEl);
            }
        }

        // Only validate if at least one payment method is selected
        if (cashCheck || transferCheck) {
            const paymentTotal = cashAmount + transferAmount;
            const diff = Math.abs(paymentTotal - totals.total);

            if (diff > 0.01) { // Allow small rounding difference
                warningEl.innerHTML = `⚠️ เงินสด + เงินโอน (${Invoice.formatCurrency(paymentTotal)}) ≠ ยอดรวม (${Invoice.formatCurrency(totals.total)})`;
                warningEl.style.display = 'block';
            } else {
                warningEl.innerHTML = '✅ ยอดถูกต้อง';
                warningEl.style.display = 'block';
                warningEl.className = 'payment-warning success';
                setTimeout(() => {
                    warningEl.style.display = 'none';
                    warningEl.className = 'payment-warning';
                }, 2000);
            }
        } else {
            warningEl.style.display = 'none';
        }
    },

    /**
     * อัปเดต Preview ใบกำกับภาษี
     */
    updatePreview() {
        const totals = Invoice.calculateTotal(this.items);

        // Get branch type
        const branchType = document.querySelector('input[name="branchType"]:checked')?.value || 'hq';
        const branchNumber = document.getElementById('branchNumber').value || '';

        // Get payment info
        const paymentCash = document.getElementById('paymentCash').checked;
        const paymentCashAmount = parseFloat(document.getElementById('paymentCashAmount').value) || 0;
        const paymentTransfer = document.getElementById('paymentTransfer').checked;
        const paymentTransferAmount = parseFloat(document.getElementById('paymentTransferAmount').value) || 0;

        const data = {
            invoiceNumber: document.getElementById('invoiceNumber').value,
            date: document.getElementById('invoiceDate').value,
            customer: {
                name: document.getElementById('customerName').value,
                address: document.getElementById('customerAddress').value,
                taxId: document.getElementById('customerTaxId').value,
                phone: document.getElementById('customerPhone').value
            },
            branchType: branchType,
            branchNumber: branchNumber,
            payment: {
                cash: paymentCash,
                cashAmount: paymentCashAmount,
                transfer: paymentTransfer,
                transferAmount: paymentTransferAmount
            },
            items: this.items,
            subtotal: totals.subtotal,
            vat: totals.vat,
            total: totals.total
        };

        document.getElementById('invoicePreview').innerHTML = Invoice.generatePreviewHTML(data);
        this.updateTotals();
    },

    /**
     * พิมพ์ใบกำกับภาษี
     */
    printInvoice() {
        const totals = Invoice.calculateTotal(this.items);

        // Get branch type
        const branchType = document.querySelector('input[name="branchType"]:checked')?.value || 'hq';
        const branchNumber = document.getElementById('branchNumber').value || '';

        // Get payment info
        const paymentCash = document.getElementById('paymentCash').checked;
        const paymentCashAmount = parseFloat(document.getElementById('paymentCashAmount').value) || 0;
        const paymentTransfer = document.getElementById('paymentTransfer').checked;
        const paymentTransferAmount = parseFloat(document.getElementById('paymentTransferAmount').value) || 0;

        // Get copy option
        const includeCopy = document.getElementById('includeCopy').checked;

        const data = {
            invoiceNumber: document.getElementById('invoiceNumber').value,
            date: document.getElementById('invoiceDate').value,
            customer: {
                name: document.getElementById('customerName').value,
                address: document.getElementById('customerAddress').value,
                taxId: document.getElementById('customerTaxId').value,
                phone: document.getElementById('customerPhone').value
            },
            branchType: branchType,
            branchNumber: branchNumber,
            payment: {
                cash: paymentCash,
                cashAmount: paymentCashAmount,
                transfer: paymentTransfer,
                transferAmount: paymentTransferAmount
            },
            items: this.items,
            subtotal: totals.subtotal,
            vat: totals.vat,
            total: totals.total,
            includeCopy: includeCopy
        };

        Invoice.print(data);
    },

    /**
     * ส่งออก PDF
     */
    async exportPDF() {
        const totals = Invoice.calculateTotal(this.items);

        // Get branch type
        const branchType = document.querySelector('input[name="branchType"]:checked')?.value || 'hq';
        const branchNumber = document.getElementById('branchNumber').value || '';

        // Get payment info
        const paymentCash = document.getElementById('paymentCash').checked;
        const paymentCashAmount = parseFloat(document.getElementById('paymentCashAmount').value) || 0;
        const paymentTransfer = document.getElementById('paymentTransfer').checked;
        const paymentTransferAmount = parseFloat(document.getElementById('paymentTransferAmount').value) || 0;

        const data = {
            invoiceNumber: document.getElementById('invoiceNumber').value,
            date: document.getElementById('invoiceDate').value,
            customer: {
                name: document.getElementById('customerName').value,
                address: document.getElementById('customerAddress').value,
                taxId: document.getElementById('customerTaxId').value,
                phone: document.getElementById('customerPhone').value
            },
            branchType: branchType,
            branchNumber: branchNumber,
            payment: {
                cash: paymentCash,
                cashAmount: paymentCashAmount,
                transfer: paymentTransfer,
                transferAmount: paymentTransferAmount
            },
            items: this.items,
            subtotal: totals.subtotal,
            vat: totals.vat,
            total: totals.total,
            includeCopy: document.getElementById('includeCopy').checked
        };

        this.showToast('กำลังสร้างไฟล์ PDF...', 'info');

        try {
            await Invoice.generatePDF(data);
            this.showToast('สร้างไฟล์ PDF สำเร็จ', 'success');
        } catch (error) {
            console.error('Error generating PDF:', error);
            this.showToast('ไม่สามารถสร้าง PDF ได้', 'error');
        }
    },

    /**
     * บันทึกใบกำกับภาษี
     */
    async saveInvoice() {
        const customerName = document.getElementById('customerName').value;
        if (!customerName) {
            this.showToast('กรุณากรอกชื่อลูกค้า', 'error');
            return;
        }

        if (this.items.length === 0 || this.items.every(i => !i.description)) {
            this.showToast('กรุณาเพิ่มรายการสินค้า', 'error');
            return;
        }

        const totals = Invoice.calculateTotal(this.items);

        // Get branch info
        const branchType = document.querySelector('input[name="branchType"]:checked')?.value || 'hq';
        const branchNumber = document.getElementById('branchNumber').value || '';

        // Get payment info
        const paymentCash = document.getElementById('paymentCash').checked;
        const paymentCashAmount = parseFloat(document.getElementById('paymentCashAmount').value) || 0;
        const paymentTransfer = document.getElementById('paymentTransfer').checked;
        const paymentTransferAmount = parseFloat(document.getElementById('paymentTransferAmount').value) || 0;

        const invoiceData = {
            invoiceNumber: Storage.getNextInvoiceNumber(),
            date: document.getElementById('invoiceDate').value,
            customerName: document.getElementById('customerName').value,
            customerAddress: document.getElementById('customerAddress').value,
            customerTaxId: document.getElementById('customerTaxId').value,
            customerPhone: document.getElementById('customerPhone').value,
            branchType: branchType,
            branchNumber: branchNumber,
            payment: {
                cash: paymentCash,
                cashAmount: paymentCashAmount,
                transfer: paymentTransfer,
                transferAmount: paymentTransferAmount
            },
            items: this.items.filter(i => i.description),
            subtotal: totals.subtotal,
            vat: totals.vat,
            total: totals.total,
            thaiText: Invoice.amountToThaiText(totals.total),
            createdAt: new Date().toISOString()
        };

        try {
            await SheetsAPI.saveInvoice(invoiceData);
            this.showToast('บันทึกใบกำกับภาษีสำเร็จ', 'success');

            // Reset form
            this.resetInvoiceForm();

            // Refresh dashboard
            this.loadDashboard();
        } catch (error) {
            console.error('Error saving invoice:', error);
            this.showToast('บันทึกใบกำกับภาษีสำเร็จ (เก็บในเครื่อง)', 'success');
            this.resetInvoiceForm();
        }
    },

    /**
     * บันทึกและพิมพ์ใบกำกับภาษี
     */
    async saveAndPrintInvoice() {
        // ป้องกันการคลิกซ้ำ
        const saveAndPrintBtn = document.getElementById('saveAndPrintBtn');
        if (saveAndPrintBtn.disabled || this.isSaving) {
            console.log('กำลังบันทึกอยู่... กรุณารอสักครู่');
            return;
        }

        // ตั้ง flag และ disable ปุ่มทันที
        this.isSaving = true;
        const originalBtnText = saveAndPrintBtn.innerHTML;
        saveAndPrintBtn.disabled = true;
        saveAndPrintBtn.innerHTML = '⏳ กำลังบันทึก...';

        const customerName = document.getElementById('customerName').value;
        if (!customerName) {
            this.showToast('กรุณากรอกชื่อลูกค้า', 'error');
            saveAndPrintBtn.disabled = false;
            saveAndPrintBtn.innerHTML = originalBtnText;
            this.isSaving = false;
            return;
        }

        if (this.items.length === 0 || this.items.every(i => !i.description)) {
            this.showToast('กรุณาเพิ่มรายการสินค้า', 'error');
            saveAndPrintBtn.disabled = false;
            saveAndPrintBtn.innerHTML = originalBtnText;
            this.isSaving = false;
            return;
        }

        const totals = Invoice.calculateTotal(this.items);

        // Get branch info
        const branchType = document.querySelector('input[name="branchType"]:checked')?.value || 'hq';
        const branchNumber = document.getElementById('branchNumber').value || '';

        // Get payment info
        const paymentCash = document.getElementById('paymentCash').checked;
        const paymentCashAmount = parseFloat(document.getElementById('paymentCashAmount').value) || 0;
        const paymentTransfer = document.getElementById('paymentTransfer').checked;
        const paymentTransferAmount = parseFloat(document.getElementById('paymentTransferAmount').value) || 0;

        // Get copy option
        const includeCopy = document.getElementById('includeCopy').checked;

        // ดึงเลขใบกำกับหลังจาก validate และ disable ปุ่มแล้ว
        const invoiceNumber = Storage.getNextInvoiceNumber();

        // Get customer data
        const customerTaxId = document.getElementById('customerTaxId').value;
        const customerAddress = document.getElementById('customerAddress').value;
        const customerPhone = document.getElementById('customerPhone').value;

        const invoiceData = {
            invoiceNumber: invoiceNumber,
            date: document.getElementById('invoiceDate').value,
            customerName: customerName,
            customerAddress: customerAddress,
            customerTaxId: customerTaxId,
            customerPhone: customerPhone,
            branchType: branchType,
            branchNumber: branchNumber,
            payment: {
                cash: paymentCash,
                cashAmount: paymentCashAmount,
                transfer: paymentTransfer,
                transferAmount: paymentTransferAmount
            },
            items: this.items.filter(i => i.description),
            subtotal: totals.subtotal,
            vat: totals.vat,
            total: totals.total,
            thaiText: Invoice.amountToThaiText(totals.total),
            createdAt: new Date().toISOString()
        };

        // Data for printing
        const printData = {
            invoiceNumber: invoiceNumber,
            date: document.getElementById('invoiceDate').value,
            customer: {
                name: customerName,
                address: customerAddress,
                taxId: customerTaxId,
                phone: customerPhone
            },
            branchType: branchType,
            branchNumber: branchNumber,
            payment: {
                cash: paymentCash,
                cashAmount: paymentCashAmount,
                transfer: paymentTransfer,
                transferAmount: paymentTransferAmount
            },
            items: this.items,
            subtotal: totals.subtotal,
            vat: totals.vat,
            total: totals.total,
            includeCopy: includeCopy
        };

        // แสดง Loading
        this.showLoading('กำลังบันทึกข้อมูล และเตรียมพิมพ์...');

        try {
            await SheetsAPI.saveInvoice(invoiceData);
            this.showToast('บันทึกใบกำกับภาษีสำเร็จ', 'success');
        } catch (error) {
            console.error('Error saving invoice:', error);
            this.showToast('บันทึกใบกำกับภาษีสำเร็จ (เก็บในเครื่อง)', 'success');
        }

        // ส่งข้อมูลลูกค้าไปอัพเดทที่ Google Sheets (ถ้ามีการเลือกลูกค้าหรือกรอก taxId)
        if (this.selectedCustomer || customerTaxId) {
            // เช็คว่ามีการแก้ไขข้อมูลลูกค้าหรือไม่
            const originalCustomer = this.selectedCustomer;
            // ใช้ String() เพื่อเปรียบเทียบให้ type ตรงกัน
            const hasChanges = !originalCustomer ||
                String(originalCustomer.name || '') !== customerName ||
                String(originalCustomer.address || '') !== customerAddress ||
                String(originalCustomer.taxId || '') !== customerTaxId ||
                String(originalCustomer.phone || '') !== customerPhone;

            console.log('Customer change detection:', {
                originalCustomer,
                currentData: { customerName, customerAddress, customerTaxId, customerPhone },
                hasChanges
            });

            if (hasChanges) {
                try {
                    const customerData = {
                        id: originalCustomer?.id || 'CUST-' + Date.now(),
                        name: customerName,
                        address: customerAddress,
                        taxId: customerTaxId,
                        phone: customerPhone,
                        email: originalCustomer?.email || ''
                    };
                    await SheetsAPI.saveCustomer(customerData);
                    console.log('Customer data updated to Google Sheets');
                    this.showToast('อัพเดทข้อมูลลูกค้าสำเร็จ', 'info');
                } catch (error) {
                    console.error('Error updating customer:', error);
                    // ไม่ต้องแสดง error เพราะไม่ใช่ action หลัก
                }
            } else {
                console.log('No customer changes detected, skipping update');
            }
        }

        // Print invoice
        Invoice.print(printData);

        // ซ่อน Loading และ enable ปุ่ม
        this.hideLoading();
        saveAndPrintBtn.disabled = false;
        saveAndPrintBtn.innerHTML = originalBtnText;
        this.isSaving = false;

        // Reset form
        this.resetInvoiceForm();

        // Refresh dashboard
        this.loadDashboard();

        // กลับไปหน้าสร้างใบกำกับ
        this.switchTab('invoice');
    },

    /**
     * รีเซ็ตฟอร์มใบกำกับภาษี
     */
    resetInvoiceForm() {
        document.getElementById('invoiceNumber').value = Storage.previewNextInvoiceNumber();
        document.getElementById('invoiceDate').valueAsDate = new Date();
        document.getElementById('customerSearch').value = '';
        document.getElementById('customerName').value = '';
        document.getElementById('customerAddress').value = '';
        document.getElementById('customerTaxId').value = '';
        document.getElementById('customerPhone').value = '';

        // Reset payment fields
        document.getElementById('paymentCash').checked = true;
        document.getElementById('paymentCashAmount').value = '';
        document.getElementById('paymentCashAmount').disabled = false;
        document.getElementById('paymentTransfer').checked = false;
        document.getElementById('paymentTransferAmount').value = '';
        document.getElementById('paymentTransferAmount').disabled = true;

        // Reset branch type
        document.getElementById('branchTypeHQ').checked = true;
        document.getElementById('branchNumber').value = '';
        document.getElementById('branchNumberGroup').style.display = 'none';

        this.selectedCustomer = null;
        this.items = [];
        this.addItem();
        this.updatePreview();
    },

    /**
     * สลับ Settings Tab
     */
    switchSettingsTab(tabName) {
        document.querySelectorAll('.settings-tab').forEach(tab => {
            tab.classList.toggle('active', tab.dataset.settings === tabName);
        });
        document.querySelectorAll('.settings-content').forEach(content => {
            content.classList.toggle('active', content.id === tabName + 'Settings');
        });
    },

    /**
     * บันทึกการตั้งค่า
     */
    saveSettings() {
        const settings = {
            invoicePrefix: document.getElementById('invoicePrefix').value,
            invoiceStartNumber: parseInt(document.getElementById('invoiceStartNumber').value) || 1,
            invoiceNumberPadding: parseInt(document.getElementById('invoiceNumberPadding').value) || 4,
            defaultCategory: document.getElementById('defaultCategory').value,
            vatRate: 7,
            sheetsUrl: document.getElementById('sheetsUrl').value,
            scriptUrl: document.getElementById('scriptUrl').value
        };

        const company = {
            name: document.getElementById('companyName').value,
            address: document.getElementById('companyAddress').value,
            taxId: document.getElementById('companyTaxId').value,
            phone: document.getElementById('companyPhone').value
        };

        Storage.saveSettings(settings);
        Storage.saveCompany(company);

        // รีเซ็ต invoice counter ให้ตรงกับเลขที่เริ่มต้นใหม่
        Storage.resetInvoiceCounter(settings.invoiceStartNumber);

        // Update invoice number preview
        document.getElementById('invoiceNumber').value = Storage.previewNextInvoiceNumber();

        this.updatePreview();
        this.closeModal('settingsModal');
        this.showToast('บันทึกการตั้งค่าสำเร็จ', 'success');
    },

    /**
     * ตั้งค่าเลขรันใบกำกับด้วยตนเอง (สำหรับ Admin)
     */
    setInvoiceCounter() {
        const newNumber = parseInt(document.getElementById('invoiceCurrentNumber').value) || 1;

        if (newNumber < 1) {
            this.showToast('เลขรันต้องมากกว่า 0', 'error');
            return;
        }

        // ตั้งค่าเลขรันใหม่
        Storage.resetInvoiceCounter(newNumber);

        // อัพเดทเลขในฟอร์มหลักและ preview
        const nextNumber = Storage.previewNextInvoiceNumber();
        document.getElementById('invoiceNumber').value = nextNumber;
        document.getElementById('invoiceCurrentNumber').value = newNumber;

        // อัพเดท preview ใบกำกับ
        this.updatePreview();

        // ปิด modal และแจ้งผู้ใช้
        this.closeModal('settingsModal');
        this.showToast(`ตั้งค่าเลขรันเป็น ${newNumber} สำเร็จ (ใบถัดไป: ${nextNumber})`, 'success');
    },

    /**
     * ทดสอบการเชื่อมต่อ
     */
    async testConnection() {
        const statusEl = document.getElementById('connectionStatus');
        statusEl.textContent = 'กำลังทดสอบการเชื่อมต่อ...';
        statusEl.className = 'connection-status';

        // Save current values first
        const settings = {
            sheetsUrl: document.getElementById('sheetsUrl').value,
            scriptUrl: document.getElementById('scriptUrl').value,
            ...Storage.getSettings()
        };
        settings.sheetsUrl = document.getElementById('sheetsUrl').value;
        settings.scriptUrl = document.getElementById('scriptUrl').value;
        Storage.saveSettings(settings);

        try {
            const result = await SheetsAPI.testConnection();

            if (result.sheetsOk) {
                statusEl.textContent = '✅ เชื่อมต่อ Google Sheets สำเร็จ';
                statusEl.className = 'connection-status success';
            } else {
                statusEl.textContent = '❌ ' + result.message;
                statusEl.className = 'connection-status error';
            }
        } catch (error) {
            statusEl.textContent = '❌ ไม่สามารถเชื่อมต่อได้: ' + error.message;
            statusEl.className = 'connection-status error';
        }
    },

    /**
     * ดึงข้อมูลจาก Sheets
     */
    async syncData() {
        const statusEl = document.getElementById('connectionStatus');
        statusEl.textContent = 'กำลังดึงข้อมูล...';
        statusEl.className = 'connection-status';

        try {
            await this.loadCustomers();
            statusEl.textContent = '✅ ดึงข้อมูลสำเร็จ พบลูกค้า ' + this.customers.length + ' ราย';
            statusEl.className = 'connection-status success';
        } catch (error) {
            statusEl.textContent = '❌ ไม่สามารถดึงข้อมูลได้: ' + error.message;
            statusEl.className = 'connection-status error';
        }
    },

    /**
     * เปิด Modal ลูกค้า
     */
    openCustomerModal(customerId = null) {
        const modal = document.getElementById('customerModal');
        const title = document.getElementById('customerModalTitle');

        // Convert to string for comparison
        const searchId = customerId ? String(customerId) : null;

        if (searchId) {
            // Find customer by ID (compare as strings)
            const customer = this.customers.find(c => String(c.id) === searchId);

            if (customer) {
                title.textContent = '✏️ แก้ไขข้อมูลลูกค้า';
                document.getElementById('editCustomerId').value = customer.id;
                document.getElementById('modalCustomerId').value = customer.id;
                document.getElementById('modalCustomerName').value = customer.name || '';
                document.getElementById('modalCustomerAddress').value = customer.address || '';
                document.getElementById('modalCustomerTaxId').value = String(customer.taxId || '');
                document.getElementById('modalCustomerPhone').value = String(customer.phone || '');
                document.getElementById('modalCustomerEmail').value = customer.email || '';
            } else {
                // Customer not found - show add new form
                console.warn('Customer not found:', searchId);
                title.textContent = '👤 เพิ่มลูกค้าใหม่';
                this.clearCustomerForm();
            }
        } else {
            title.textContent = '👤 เพิ่มลูกค้าใหม่';
            this.clearCustomerForm();
        }

        modal.classList.add('active');
    },

    /**
     * ล้างฟอร์มลูกค้า
     */
    clearCustomerForm() {
        document.getElementById('editCustomerId').value = '';
        document.getElementById('modalCustomerId').value = 'CUST-' + Date.now();
        document.getElementById('modalCustomerName').value = '';
        document.getElementById('modalCustomerAddress').value = '';
        document.getElementById('modalCustomerTaxId').value = '';
        document.getElementById('modalCustomerPhone').value = '';
        document.getElementById('modalCustomerEmail').value = '';
    },

    /**
     * บันทึกลูกค้า
     */
    async saveCustomer() {
        const name = document.getElementById('modalCustomerName').value;
        if (!name) {
            this.showToast('กรุณากรอกชื่อลูกค้า', 'error');
            return;
        }

        const customer = {
            id: document.getElementById('modalCustomerId').value,
            name: name,
            address: document.getElementById('modalCustomerAddress').value,
            taxId: String(document.getElementById('modalCustomerTaxId').value || ''),
            phone: String(document.getElementById('modalCustomerPhone').value || ''),
            email: document.getElementById('modalCustomerEmail').value
        };

        const editId = document.getElementById('editCustomerId').value;

        // แสดง Loading และ disable ปุ่มบันทึก
        const saveBtn = document.getElementById('saveCustomerBtn');
        const originalBtnText = saveBtn.innerHTML;
        saveBtn.disabled = true;
        saveBtn.innerHTML = '⏳ กำลังบันทึก...';
        this.showLoading('รอสักครู่ กำลังบันทึกข้อมูล...');

        try {
            if (editId) {
                await SheetsAPI.updateCustomer(customer);
                this.showToast('อัพเดทข้อมูลลูกค้าสำเร็จ', 'success');
            } else {
                await SheetsAPI.saveCustomer(customer);
                this.showToast('เพิ่มลูกค้าใหม่สำเร็จ', 'success');
            }

            // เก็บข้อมูลลูกค้าไว้สำหรับ auto-fill
            const savedCustomer = { ...customer };
            const isNewCustomer = !editId;

            // Reload customers จาก Google Sheets ใหม่
            this.showLoading('กำลังดึงข้อมูลลูกค้าใหม่...');
            await this.loadCustomers();
            this.closeModal('customerModal');

            // Auto-fill ข้อมูลลูกค้าใหม่ลงฟอร์ม (เฉพาะกรณีเพิ่มใหม่)
            if (isNewCustomer) {
                this.selectedCustomer = savedCustomer;
                document.getElementById('customerSearch').value = savedCustomer.name || '';
                document.getElementById('customerName').value = savedCustomer.name || '';
                document.getElementById('customerAddress').value = savedCustomer.address || '';
                document.getElementById('customerTaxId').value = String(savedCustomer.taxId || '');
                document.getElementById('customerPhone').value = String(savedCustomer.phone || '');
                this.updatePreview();
                this.showToast(`เลือกลูกค้า: ${savedCustomer.name}`, 'success');
            }
        } catch (error) {
            console.error('Error saving customer:', error);
            this.showToast('บันทึกลูกค้าสำเร็จ (เก็บในเครื่อง)', 'success');
            this.customers = Storage.getCustomers();
            this.renderCustomersTable();
            this.closeModal('customerModal');
        } finally {
            // ซ่อน Loading และ enable ปุ่มบันทึก
            this.hideLoading();
            saveBtn.disabled = false;
            saveBtn.innerHTML = originalBtnText;
        }
    },

    /**
     * Render ตารางลูกค้า
     */
    renderCustomersTable() {
        const tbody = document.getElementById('customersBody');

        if (this.customers.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" class="text-center">ไม่พบข้อมูลลูกค้า</td></tr>';
            return;
        }

        tbody.innerHTML = this.customers.map(c => `
            <tr>
                <td>${c.id}</td>
                <td>${c.name}</td>
                <td>${c.taxId || '-'}</td>
                <td>${c.phone || '-'}</td>
                <td class="actions">
                    <button class="btn btn-secondary btn-sm" onclick="App.openCustomerModal('${c.id}')">✏️ แก้ไข</button>
                </td>
            </tr>
        `).join('');
    },

    /**
     * กรองรายการลูกค้า
     */
    filterCustomersList(query) {
        if (!query || query.trim() === '') {
            // ถ้าไม่มีคำค้นหา ให้แสดงทั้งหมด
            this.renderCustomersTable();
            return;
        }

        const queryLower = query.toLowerCase().trim();

        const filtered = this.customers.filter(c => {
            const name = String(c.name || '').toLowerCase();
            const id = String(c.id || '').toLowerCase();
            const taxId = String(c.taxId || '');

            return name.includes(queryLower) ||
                id.includes(queryLower) ||
                taxId.includes(query);  // เลขผู้เสียภาษีไม่ต้อง toLowerCase
        });

        const tbody = document.getElementById('customersBody');

        if (filtered.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" class="text-center">ไม่พบข้อมูลลูกค้า</td></tr>';
            return;
        }

        tbody.innerHTML = filtered.map(c => `
            <tr>
                <td>${c.id || '-'}</td>
                <td>${c.name || '-'}</td>
                <td>${c.taxId || '-'}</td>
                <td>${c.phone || '-'}</td>
                <td class="actions">
                    <button class="btn btn-secondary btn-sm" onclick="App.openCustomerModal('${c.id}')">✏️ แก้ไข</button>
                </td>
            </tr>
        `).join('');
    },


    /**
     * โหลด Dashboard
     */
    loadDashboard() {
        // Initialize dashboard filters
        this.initDashboardFilters();

        // Load dashboard with current filter
        this.updateDashboardStats();
    },

    /**
     * Initialize dashboard filter controls
     */
    initDashboardFilters() {
        const yearSelect = document.getElementById('dashboardYear');
        const monthSelect = document.getElementById('dashboardMonth');

        // ตรวจสอบว่า event listener ถูกเพิ่มแล้วหรือยัง
        if (yearSelect && !yearSelect.dataset.initialized) {
            // Populate year dropdown (current year and 2 previous years)
            const currentYear = new Date().getFullYear();
            yearSelect.innerHTML = '';
            for (let y = currentYear; y >= currentYear - 2; y--) {
                const option = document.createElement('option');
                option.value = y;
                option.textContent = y + 543; // Convert to Buddhist year
                if (y === currentYear) option.selected = true;
                yearSelect.appendChild(option);
            }

            // Add event listeners
            yearSelect.addEventListener('change', () => this.updateDashboardStats());
            yearSelect.dataset.initialized = 'true';
        }

        if (monthSelect && !monthSelect.dataset.initialized) {
            // Set current month as default
            const currentMonth = new Date().getMonth() + 1;
            monthSelect.value = currentMonth;

            monthSelect.addEventListener('change', () => this.updateDashboardStats());
            monthSelect.dataset.initialized = 'true';
        }

        // Apply date range button
        const applyBtn = document.getElementById('applyDateRangeBtn');
        if (applyBtn && !applyBtn.dataset.initialized) {
            applyBtn.addEventListener('click', () => this.applyCustomDateRange());
            applyBtn.dataset.initialized = 'true';
        }

        // Reset filter button
        const resetBtn = document.getElementById('resetFilterBtn');
        if (resetBtn && !resetBtn.dataset.initialized) {
            resetBtn.addEventListener('click', () => this.resetDashboardFilters());
            resetBtn.dataset.initialized = 'true';
        }
    },

    /**
     * Apply custom date range filter
     */
    applyCustomDateRange() {
        const dateFrom = document.getElementById('dashboardDateFrom').value;
        const dateTo = document.getElementById('dashboardDateTo').value;

        if (!dateFrom || !dateTo) {
            this.showToast('กรุณาเลือกวันที่เริ่มต้นและสิ้นสุด', 'error');
            return;
        }

        if (dateFrom > dateTo) {
            this.showToast('วันที่เริ่มต้นต้องน้อยกว่าวันที่สิ้นสุด', 'error');
            return;
        }

        // Clear month/year selection when using custom date range
        document.getElementById('dashboardMonth').value = 'all';

        this.updateDashboardStats(dateFrom, dateTo);
    },

    /**
     * Reset dashboard filters to current month
     */
    resetDashboardFilters() {
        const currentYear = new Date().getFullYear();
        const currentMonth = new Date().getMonth() + 1;

        document.getElementById('dashboardYear').value = currentYear;
        document.getElementById('dashboardMonth').value = currentMonth;
        document.getElementById('dashboardDateFrom').value = '';
        document.getElementById('dashboardDateTo').value = '';

        this.updateDashboardStats();
        this.showToast('รีเซ็ตตัวกรองแล้ว', 'success');
    },

    /**
     * Update dashboard statistics based on filter
     */
    updateDashboardStats(customDateFrom = null, customDateTo = null) {
        const invoices = Storage.getInvoices();
        const customers = Storage.getCustomers();

        // Get filter values
        const selectedYear = parseInt(document.getElementById('dashboardYear')?.value) || new Date().getFullYear();
        const selectedMonth = document.getElementById('dashboardMonth')?.value || 'all';

        // แยก invoice ที่ active และที่ยกเลิก
        let activeInvoices = invoices.filter(inv => inv.status !== 'cancelled');
        let cancelledInvoices = invoices.filter(inv => inv.status === 'cancelled');

        // Apply date filter
        if (customDateFrom && customDateTo) {
            // Custom date range filter
            activeInvoices = activeInvoices.filter(inv => {
                const invDate = inv.date;
                return invDate >= customDateFrom && invDate <= customDateTo;
            });
            cancelledInvoices = cancelledInvoices.filter(inv => {
                const invDate = inv.date;
                return invDate >= customDateFrom && invDate <= customDateTo;
            });
        } else {
            // Month/Year filter
            activeInvoices = activeInvoices.filter(inv => {
                if (!inv.date) return false;
                const invDate = new Date(inv.date);
                const invYear = invDate.getFullYear();
                const invMonth = invDate.getMonth() + 1;

                if (selectedMonth === 'all') {
                    return invYear === selectedYear;
                }
                return invYear === selectedYear && invMonth === parseInt(selectedMonth);
            });
            cancelledInvoices = cancelledInvoices.filter(inv => {
                if (!inv.date) return false;
                const invDate = new Date(inv.date);
                const invYear = invDate.getFullYear();
                const invMonth = invDate.getMonth() + 1;

                if (selectedMonth === 'all') {
                    return invYear === selectedYear;
                }
                return invYear === selectedYear && invMonth === parseInt(selectedMonth);
            });
        }

        // Stats - นับเฉพาะบิลที่ยังไม่ถูกยกเลิก
        const totalInvoices = activeInvoices.length + cancelledInvoices.length;
        const totalSales = activeInvoices.reduce((sum, inv) => sum + (inv.total || 0), 0);
        const today = new Date().toISOString().split('T')[0];
        const todayInvoices = activeInvoices.filter(inv => inv.date === today).length;
        const totalCustomers = customers.length;

        // คำนวณ VAT รวม
        const totalVat = activeInvoices.reduce((sum, inv) => sum + (inv.vat || 0), 0);

        // Stats ของบิลยกเลิก
        const cancelledCount = cancelledInvoices.length;
        const cancelledTotal = cancelledInvoices.reduce((sum, inv) => sum + (inv.total || 0), 0);

        document.getElementById('totalInvoices').textContent = totalInvoices;
        document.getElementById('totalSales').textContent = '฿' + Invoice.formatCurrency(totalSales);
        document.getElementById('todayInvoices').textContent = todayInvoices;
        document.getElementById('totalCustomers').textContent = totalCustomers;

        // แสดง VAT รวม
        const vatEl = document.getElementById('totalVat');
        if (vatEl) vatEl.textContent = '฿' + Invoice.formatCurrency(totalVat);

        // แสดงสถิติบิลยกเลิก (ถ้ามี element)
        const cancelledEl = document.getElementById('cancelledInvoices');
        const cancelledAmountEl = document.getElementById('cancelledAmount');
        if (cancelledEl) cancelledEl.textContent = cancelledCount;
        if (cancelledAmountEl) cancelledAmountEl.textContent = '฿' + Invoice.formatCurrency(cancelledTotal);

        // Recent invoices - แสดงเฉพาะที่ active
        const recentBody = document.getElementById('recentInvoicesBody');
        const recent = activeInvoices.slice(0, 5);

        if (recent.length === 0) {
            recentBody.innerHTML = '<tr><td colspan="4" class="text-center">ไม่มีใบกำกับภาษีในช่วงเวลาที่เลือก</td></tr>';
        } else {
            recentBody.innerHTML = recent.map(inv => `
                <tr>
                    <td>${inv.invoiceNumber}</td>
                    <td>${Invoice.formatDateShort(inv.date)}</td>
                    <td>${inv.customerName}</td>
                    <td class="text-right">${Invoice.formatCurrency(inv.total)} บาท</td>
                </tr>
            `).join('');
        }

        // Chart - ใช้เฉพาะ active invoices จากปีที่เลือก
        const yearInvoices = invoices.filter(inv => {
            if (inv.status === 'cancelled' || !inv.date) return false;
            const invDate = new Date(inv.date);
            return invDate.getFullYear() === selectedYear;
        });
        this.renderChart(yearInvoices);
    },

    /**
     * Render กราฟยอดขาย
     */
    renderChart(invoices) {
        const ctx = document.getElementById('salesChart');
        if (!ctx) return;

        // Group by month
        const monthlyData = {};
        const thaiMonths = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

        // Initialize last 6 months
        for (let i = 5; i >= 0; i--) {
            const date = new Date();
            date.setMonth(date.getMonth() - i);
            const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
            monthlyData[key] = { month: thaiMonths[date.getMonth()], total: 0 };
        }

        // Sum invoices by month
        invoices.forEach(inv => {
            if (!inv.date) return;
            const key = inv.date.substring(0, 7);
            if (monthlyData[key]) {
                monthlyData[key].total += inv.total || 0;
            }
        });

        const labels = Object.values(monthlyData).map(d => d.month);
        const data = Object.values(monthlyData).map(d => d.total);

        // Destroy existing chart
        if (this.chart) {
            this.chart.destroy();
        }

        this.chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: 'ยอดขาย (บาท)',
                    data: data,
                    backgroundColor: 'rgba(99, 102, 241, 0.8)',
                    borderColor: 'rgb(99, 102, 241)',
                    borderWidth: 1,
                    borderRadius: 8
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            callback: function (value) {
                                return '฿' + value.toLocaleString();
                            }
                        }
                    }
                }
            }
        });
    },

    /**
     * โหลดประวัติใบกำกับภาษี
     */
    loadHistory() {
        const invoices = Storage.getInvoices();
        this.renderHistoryTable(invoices);
    },

    /**
     * ค้นหาใบกำกับภาษี
     */
    searchInvoices() {
        const query = document.getElementById('invoiceSearch').value.toLowerCase();
        const fromDate = document.getElementById('searchDateFrom').value;
        const toDate = document.getElementById('searchDateTo').value;

        let invoices = Storage.getInvoices();

        if (query) {
            invoices = invoices.filter(inv =>
                inv.invoiceNumber.toLowerCase().includes(query) ||
                inv.customerName.toLowerCase().includes(query)
            );
        }

        if (fromDate) {
            invoices = invoices.filter(inv => inv.date >= fromDate);
        }

        if (toDate) {
            invoices = invoices.filter(inv => inv.date <= toDate);
        }

        this.renderHistoryTable(invoices);
    },

    /**
     * Get filtered invoices based on current search criteria
     */
    getFilteredInvoices() {
        const query = document.getElementById('invoiceSearch').value.toLowerCase();
        const fromDate = document.getElementById('searchDateFrom').value;
        const toDate = document.getElementById('searchDateTo').value;

        let invoices = Storage.getInvoices();

        if (query) {
            invoices = invoices.filter(inv =>
                inv.invoiceNumber.toLowerCase().includes(query) ||
                inv.customerName.toLowerCase().includes(query)
            );
        }

        if (fromDate) {
            invoices = invoices.filter(inv => inv.date >= fromDate);
        }

        if (toDate) {
            invoices = invoices.filter(inv => inv.date <= toDate);
        }

        return invoices;
    },

    /**
     * Export history to PDF
     */
    exportHistoryToPDF() {
        const invoices = this.getFilteredInvoices();

        if (invoices.length === 0) {
            this.showToast('ไม่มีข้อมูลใบกำกับภาษีสำหรับส่งออก', 'error');
            return;
        }

        // สร้าง HTML สำหรับ PDF
        const now = new Date();
        const dateStr = Invoice.formatDateShort(now.toISOString().split('T')[0]);

        // คำนวณยอดรวม
        const activeInvoices = invoices.filter(inv => inv.status !== 'cancelled');
        const totalAmount = activeInvoices.reduce((sum, inv) => sum + (inv.total || 0), 0);
        const totalVat = activeInvoices.reduce((sum, inv) => sum + (inv.vat || 0), 0);

        const content = document.createElement('div');
        content.innerHTML = `
            <style>
                body { font-family: 'Prompt', sans-serif; font-size: 10pt; }
                .header { text-align: center; margin-bottom: 20px; background: linear-gradient(135deg, #1a5490 0%, #2563eb 100%); padding: 20px; border-radius: 8px; }
                .header h1 { color: white; font-size: 18pt; margin: 0; }
                .header p { color: rgba(255,255,255,0.9); margin: 8px 0 0 0; font-size: 11pt; }
                .summary { background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
                .summary-row { display: flex; justify-content: space-between; margin: 5px 0; }
                table { width: 100%; border-collapse: collapse; margin-top: 15px; }
                th { background: #1a5490; color: white; padding: 10px; text-align: left; font-size: 9pt; }
                td { padding: 8px 10px; border-bottom: 1px solid #ddd; font-size: 9pt; }
                tr:nth-child(even) { background: #f9f9f9; }
                .text-right { text-align: right; }
                .status-active { color: #16a34a; }
                .status-cancelled { color: #dc2626; text-decoration: line-through; }
                .footer { margin-top: 20px; text-align: center; font-size: 8pt; color: #666; }
            </style>
            <div class="header">
                <h1>📋 รายงานประวัติใบกำกับภาษี</h1>
                <p>วันที่พิมพ์: ${dateStr}</p>
            </div>
            <div class="summary">
                <div class="summary-row">
                    <span>จำนวนใบกำกับทั้งหมด:</span>
                    <strong>${invoices.length} รายการ</strong>
                </div>
                <div class="summary-row">
                    <span>ยอดขายรวม (ไม่รวมบิลยกเลิก):</span>
                    <strong>฿${Invoice.formatCurrency(totalAmount)}</strong>
                </div>
                <div class="summary-row">
                    <span>ยอดรวม VAT:</span>
                    <strong>฿${Invoice.formatCurrency(totalVat)}</strong>
                </div>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>เลขที่ใบกำกับ</th>
                        <th>วันที่</th>
                        <th>ลูกค้า</th>
                        <th class="text-right">ยอดรวม</th>
                        <th class="text-right">VAT</th>
                        <th>สถานะ</th>
                    </tr>
                </thead>
                <tbody>
                    ${invoices.map(inv => `
                        <tr class="${inv.status === 'cancelled' ? 'status-cancelled' : ''}">
                            <td>${inv.invoiceNumber || '-'}</td>
                            <td>${Invoice.formatDateShort(inv.date) || '-'}</td>
                            <td>${inv.customerName || '-'}</td>
                            <td class="text-right">${Invoice.formatCurrency(inv.total || 0)}</td>
                            <td class="text-right">${Invoice.formatCurrency(inv.vat || 0)}</td>
                            <td class="${inv.status === 'cancelled' ? 'status-cancelled' : 'status-active'}">
                                ${inv.status === 'cancelled' ? '❌ ยกเลิก' : '✅ ปกติ'}
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            <div class="footer">
                <p>รายงานนี้สร้างโดยระบบพิมพ์บิลและใบกำกับภาษี</p>
            </div>
        `;

        // ใช้ html2pdf
        const opt = {
            margin: [10, 10, 10, 10],
            filename: `invoice_history_${now.toISOString().split('T')[0]}.pdf`,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
        };

        html2pdf().set(opt).from(content).save();
        this.showToast('📄 กำลังสร้างไฟล์ PDF...', 'success');
    },

    /**
     * Export history to Excel (CSV format with BOM for Thai support)
     */
    exportHistoryToExcel() {
        const invoices = this.getFilteredInvoices();

        if (invoices.length === 0) {
            this.showToast('ไม่มีข้อมูลใบกำกับภาษีสำหรับส่งออก', 'error');
            return;
        }

        // สร้าง CSV content
        const headers = ['เลขที่ใบกำกับ', 'วันที่', 'ลูกค้า', 'ที่อยู่', 'เลขผู้เสียภาษี', 'ยอดก่อน VAT', 'VAT', 'ยอดรวม', 'สถานะ'];
        const rows = invoices.map(inv => {
            // ใช้รูปแบบ ="value" เพื่อบังคับให้ Excel เก็บเป็น text (รักษาเลข 0 นำหน้า)
            const taxId = inv.customerTaxId ? `="${inv.customerTaxId}"` : '';
            const invoiceNum = inv.invoiceNumber ? `="${inv.invoiceNumber}"` : '';

            return [
                invoiceNum,
                inv.date || '',
                inv.customerName || '',
                (inv.customerAddress || '').replace(/,/g, ' ').replace(/\n/g, ' '),
                taxId,
                inv.subtotal || 0,
                inv.vat || 0,
                inv.total || 0,
                inv.status === 'cancelled' ? 'ยกเลิก' : 'ปกติ'
            ];
        });

        // สร้าง CSV string
        let csv = '\uFEFF'; // BOM for UTF-8
        csv += headers.join(',') + '\n';
        rows.forEach(row => {
            csv += row.map(cell => {
                // Escape double quotes and wrap in quotes if contains comma
                const str = String(cell);
                if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                    return '"' + str.replace(/"/g, '""') + '"';
                }
                return str;
            }).join(',') + '\n';
        });

        // Download file
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const now = new Date();
        link.href = URL.createObjectURL(blob);
        link.download = `invoice_history_${now.toISOString().split('T')[0]}.csv`;
        link.style.display = 'none';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        this.showToast('📊 ดาวน์โหลดไฟล์ Excel (CSV) สำเร็จ', 'success');
    },

    /**
     * Sync ประวัติใบกำกับจาก Google Sheets
     * รวมข้อมูลโดยใช้ localStorage เป็นหลัก (ถ้ามี invoice number ซ้ำ)
     */
    async syncInvoiceHistory() {
        const syncBtn = document.getElementById('syncHistoryBtn');
        const originalBtnText = syncBtn.innerHTML;

        syncBtn.disabled = true;
        syncBtn.innerHTML = '⏳ กำลัง Sync...';
        this.showLoading('กำลังดึงข้อมูลจาก Google Sheets...');

        try {
            // ดึงข้อมูลจาก Google Sheets
            const sheetsInvoices = await SheetsAPI.fetchInvoices();

            // ดึงข้อมูลจาก localStorage
            const localInvoices = Storage.getInvoices();

            // สร้าง Set ของเลขใบกำกับจาก localStorage
            const localInvoiceNumbers = new Set(localInvoices.map(inv => inv.invoiceNumber));

            // เพิ่มเฉพาะใบกำกับที่ไม่มีใน localStorage
            let addedCount = 0;
            sheetsInvoices.forEach(sheetInv => {
                if (sheetInv.invoiceNumber && !localInvoiceNumbers.has(sheetInv.invoiceNumber)) {
                    localInvoices.push(sheetInv);
                    addedCount++;
                }
            });

            // เรียงลำดับตามเลขใบกำกับ (ล่าสุดก่อน)
            localInvoices.sort((a, b) => {
                // เปรียบเทียบเป็น string เพราะอาจมี prefix
                return String(b.invoiceNumber || '').localeCompare(String(a.invoiceNumber || ''));
            });

            // บันทึกกลับเข้า localStorage
            Storage.saveInvoices(localInvoices);

            // แสดงผลในตาราง
            this.renderHistoryTable(localInvoices);

            if (addedCount > 0) {
                this.showToast(`✅ Sync สำเร็จ! เพิ่มใบกำกับใหม่ ${addedCount} รายการ`, 'success');
            } else {
                this.showToast('✅ Sync สำเร็จ! ข้อมูลทั้งหมดเป็นปัจจุบันแล้ว', 'success');
            }

        } catch (error) {
            console.error('Error syncing invoice history:', error);
            this.showToast('❌ ไม่สามารถ Sync ได้: ' + error.message, 'error');
        } finally {
            this.hideLoading();
            syncBtn.disabled = false;
            syncBtn.innerHTML = originalBtnText;
        }
    },

    /**
     * ล้างข้อมูลประวัติใบกำกับทั้งหมดใน localStorage
     */
    clearInvoiceHistory() {
        const invoices = Storage.getInvoices();
        const count = invoices.length;

        if (count === 0) {
            this.showToast('ไม่มีข้อมูลใบกำกับให้ลบ', 'info');
            return;
        }

        if (!confirm(`⚠️ ยืนยันการลบประวัติใบกำกับทั้งหมด ${count} รายการ?\n\n⛔ การลบจะไม่สามารถกู้คืนได้\n(ข้อมูลใน Google Sheets จะยังคงอยู่)`)) {
            return;
        }

        // ยืนยันอีกครั้ง
        if (!confirm('❗ ยืนยันอีกครั้ง: ต้องการลบข้อมูลทั้งหมดจริงๆ?')) {
            return;
        }

        // ล้างข้อมูล
        Storage.saveInvoices([]);

        this.loadHistory();
        this.showToast(`🗑️ ลบประวัติใบกำกับ ${count} รายการสำเร็จ`, 'success');
    },

    /**
     * Render ตารางประวัติ
     */
    renderHistoryTable(invoices) {
        const tbody = document.getElementById('historyBody');

        if (invoices.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" class="text-center">ไม่พบใบกำกับภาษี</td></tr>';
            return;
        }

        tbody.innerHTML = invoices.map(inv => {
            const status = inv.status || 'active';
            const isActive = status === 'active';
            const statusBadge = isActive
                ? '<span class="status-badge status-active">✅ ปกติ</span>'
                : '<span class="status-badge status-cancelled">❌ ยกเลิก</span>';
            const rowClass = isActive ? '' : 'row-cancelled';

            return `
            <tr class="${rowClass}">
                <td>${inv.invoiceNumber}</td>
                <td>${Invoice.formatDateShort(inv.date)}</td>
                <td>${inv.customerName}</td>
                <td class="text-right">${Invoice.formatCurrency(inv.total)} บาท</td>
                <td>${statusBadge}</td>
                <td class="actions">
                    <button class="btn btn-warning btn-sm" onclick="App.editInvoice('${inv.invoiceNumber}')" ${!isActive ? 'disabled' : ''}>✏️ แก้ไข</button>
                    <button class="btn btn-primary btn-sm" onclick="App.reprintFromHistory('${inv.invoiceNumber}')">🖨️ พิมพ์</button>
                    ${isActive ? `<button class="btn btn-info btn-sm" onclick="App.openEmailModalFromHistory('${inv.invoiceNumber}')">📧 Email</button>` : ''}
                    ${isActive
                    ? `<button class="btn btn-danger btn-sm" onclick="App.cancelInvoice('${inv.invoiceNumber}')">🚫 ยกเลิก</button>`
                    : `<button class="btn btn-success btn-sm" onclick="App.restoreInvoice('${inv.invoiceNumber}')">♻️ คืนสถานะ</button>`
                }
                    <button class="btn btn-dark btn-sm" onclick="App.deleteInvoice('${inv.invoiceNumber}')" title="ลบถาวร">🗑️ ลบ</button>
                </td>
            </tr>
        `}).join('');
    },

    /**
     * ลบใบกำกับภาษี
     */
    deleteInvoice(invoiceNumber) {
        if (!confirm(`ยืนยันการลบใบกำกับภาษี ${invoiceNumber}?\n\nการลบจะไม่สามารถกู้คืนได้`)) {
            return;
        }

        const invoices = Storage.getInvoices();
        const filtered = invoices.filter(inv => inv.invoiceNumber !== invoiceNumber);

        Storage.saveInvoices(filtered);

        this.loadHistory();
        this.showToast(`ลบใบกำกับภาษี ${invoiceNumber} สำเร็จ`, 'success');
    },

    /**
     * ยกเลิกใบกำกับภาษี (เปลี่ยนสถานะเป็น cancelled)
     */
    async cancelInvoice(invoiceNumber) {
        if (!confirm(`ยืนยันการยกเลิกใบกำกับภาษี ${invoiceNumber}?\n\nใบกำกับจะถูกทำเครื่องหมายว่ายกเลิก`)) {
            return;
        }

        const invoices = Storage.getInvoices();
        const index = invoices.findIndex(inv => inv.invoiceNumber === invoiceNumber);

        if (index === -1) {
            this.showToast('ไม่พบใบกำกับภาษี', 'error');
            return;
        }

        // เปลี่ยนสถานะเป็น cancelled
        invoices[index].status = 'cancelled';
        invoices[index].cancelledAt = new Date().toISOString();

        // บันทึกลง localStorage
        Storage.saveInvoices(invoices);

        // อัปเดต Google Sheets
        try {
            const settings = Storage.getSettings();
            if (settings.scriptUrl) {
                await SheetsAPI.updateInvoice({
                    ...invoices[index],
                    items: JSON.stringify(invoices[index].items || [])
                });
            }
        } catch (error) {
            console.warn('Could not sync to Google Sheets:', error);
        }

        this.loadHistory();
        this.showToast(`ยกเลิกใบกำกับภาษี ${invoiceNumber} สำเร็จ`, 'success');
    },

    /**
     * คืนสถานะใบกำกับภาษี (เปลี่ยนสถานะกลับเป็น active)
     */
    async restoreInvoice(invoiceNumber) {
        if (!confirm(`ยืนยันการคืนสถานะใบกำกับภาษี ${invoiceNumber}?`)) {
            return;
        }

        const invoices = Storage.getInvoices();
        const index = invoices.findIndex(inv => inv.invoiceNumber === invoiceNumber);

        if (index === -1) {
            this.showToast('ไม่พบใบกำกับภาษี', 'error');
            return;
        }

        // เปลี่ยนสถานะกลับเป็น active
        invoices[index].status = 'active';
        delete invoices[index].cancelledAt;

        // บันทึกลง localStorage
        Storage.saveInvoices(invoices);

        // อัปเดต Google Sheets
        try {
            const settings = Storage.getSettings();
            if (settings.scriptUrl) {
                await SheetsAPI.updateInvoice({
                    ...invoices[index],
                    items: JSON.stringify(invoices[index].items || [])
                });
            }
        } catch (error) {
            console.warn('Could not sync to Google Sheets:', error);
        }

        this.loadHistory();
        this.showToast(`คืนสถานะใบกำกับภาษี ${invoiceNumber} สำเร็จ`, 'success');
    },

    /**
     * แก้ไขใบกำกับภาษี
     */
    editInvoice(invoiceNumber) {
        const invoices = Storage.getInvoices();
        const invoice = invoices.find(inv => inv.invoiceNumber === invoiceNumber);

        if (!invoice) {
            this.showToast('ไม่พบใบกำกับภาษี', 'error');
            return;
        }

        // Open edit invoice modal
        this.openEditInvoiceModal(invoice);
    },

    /**
     * เปิด Modal แก้ไขใบกำกับภาษี
     */
    openEditInvoiceModal(invoice) {
        // Store current editing invoice
        this.editingInvoice = JSON.parse(JSON.stringify(invoice));

        // Create modal if not exists
        let modal = document.getElementById('editInvoiceModal');
        if (!modal) {
            modal = document.createElement('div');
            modal.id = 'editInvoiceModal';
            modal.className = 'modal';
            modal.innerHTML = `
                <div class="modal-content" style="max-width: 800px; max-height: 90vh; overflow-y: auto;">
                    <div class="modal-header">
                        <h3>✏️ แก้ไขใบกำกับภาษี</h3>
                        <button class="modal-close" onclick="App.closeModal('editInvoiceModal')">&times;</button>
                    </div>
                    <div class="modal-body">
                        <input type="hidden" id="editInvoiceNumber">
                        <div class="form-row">
                            <div class="form-group">
                                <label>เลขที่ใบกำกับภาษี <span style="font-size: 10px; color: #f59e0b;">⚠️ แก้ไขได้</span></label>
                                <input type="text" id="editInvoiceNo" style="border: 2px solid #f59e0b;">
                            </div>
                            <div class="form-group">
                                <label>วันที่</label>
                                <input type="date" id="editInvoiceDate">
                            </div>
                        </div>
                        <div class="form-group">
                            <label>ชื่อลูกค้า</label>
                            <input type="text" id="editInvoiceCustomerName">
                        </div>
                        <div class="form-group">
                            <label>ที่อยู่</label>
                            <textarea id="editInvoiceCustomerAddress" rows="2"></textarea>
                        </div>
                        <div class="form-row">
                            <div class="form-group">
                                <label>เลขประจำตัวผู้เสียภาษี</label>
                                <input type="text" id="editInvoiceCustomerTaxId">
                            </div>
                            <div class="form-group">
                                <label>เบอร์โทร</label>
                                <input type="text" id="editInvoiceCustomerPhone">
                            </div>
                        </div>
                        
                        <h4 style="margin: 15px 0 10px; color: var(--primary);">📦 รายการสินค้า/บริการ</h4>
                        <div id="editInvoiceItemsContainer"></div>
                        <button class="btn btn-secondary btn-sm" onclick="App.addEditInvoiceItem()" style="margin-top: 10px;">➕ เพิ่มรายการ</button>
                        
                        <div class="form-row" style="margin-top: 15px;">
                            <div class="form-group">
                                <label>ยอดก่อน VAT</label>
                                <input type="number" id="editInvoiceSubtotal" step="0.01" readonly>
                            </div>
                            <div class="form-group">
                                <label>VAT 7%</label>
                                <input type="number" id="editInvoiceVat" step="0.01" readonly>
                            </div>
                            <div class="form-group">
                                <label>ยอดรวมทั้งสิ้น</label>
                                <input type="number" id="editInvoiceTotal" step="0.01" readonly>
                            </div>
                        </div>
                        
                        <h4 style="margin: 15px 0 10px; color: var(--primary);">💳 การรับชำระเงิน</h4>
                        <div class="payment-options">
                            <div class="payment-option">
                                <label class="checkbox-label">
                                    <input type="checkbox" id="editPaymentCash" onchange="App.toggleEditPaymentCash()">
                                    <span>💵 เงินสด</span>
                                </label>
                                <input type="number" id="editPaymentCashAmount" class="payment-amount" placeholder="จำนวนเงิน" step="0.01" disabled>
                                <span class="payment-unit">บาท</span>
                            </div>
                            <div class="payment-option">
                                <label class="checkbox-label">
                                    <input type="checkbox" id="editPaymentTransfer" onchange="App.toggleEditPaymentTransfer()">
                                    <span>💳 เงินโอน</span>
                                </label>
                                <input type="number" id="editPaymentTransferAmount" class="payment-amount" placeholder="จำนวนเงิน" step="0.01" disabled>
                                <span class="payment-unit">บาท</span>
                            </div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-secondary" onclick="App.closeModal('editInvoiceModal')">ยกเลิก</button>
                        <button class="btn btn-primary" onclick="App.saveAndPrintEditedInvoice()">🖨️ บันทึก + พิมพ์</button>
                    </div>
                </div>
            `;
            document.body.appendChild(modal);
        }

        // Fill form
        document.getElementById('editInvoiceNumber').value = invoice.invoiceNumber;
        document.getElementById('editInvoiceNo').value = invoice.invoiceNumber;
        document.getElementById('editInvoiceDate').value = invoice.date || '';
        document.getElementById('editInvoiceCustomerName').value = invoice.customerName || '';
        document.getElementById('editInvoiceCustomerAddress').value = invoice.customerAddress || '';
        document.getElementById('editInvoiceCustomerTaxId').value = invoice.customerTaxId || '';
        document.getElementById('editInvoiceCustomerPhone').value = invoice.customerPhone || '';

        // Load payment info
        const payment = invoice.payment || { cash: false, cashAmount: 0, transfer: false, transferAmount: 0 };
        document.getElementById('editPaymentCash').checked = payment.cash;
        document.getElementById('editPaymentCashAmount').value = payment.cashAmount || '';
        document.getElementById('editPaymentCashAmount').disabled = !payment.cash;
        document.getElementById('editPaymentTransfer').checked = payment.transfer;
        document.getElementById('editPaymentTransferAmount').value = payment.transferAmount || '';
        document.getElementById('editPaymentTransferAmount').disabled = !payment.transfer;

        // Load items
        this.renderEditInvoiceItems();

        modal.classList.add('active');
    },

    /**
     * Toggle เงินสดใน Modal แก้ไข
     */
    toggleEditPaymentCash() {
        const checkbox = document.getElementById('editPaymentCash');
        const amountInput = document.getElementById('editPaymentCashAmount');
        amountInput.disabled = !checkbox.checked;
        if (checkbox.checked && !amountInput.value) {
            amountInput.value = this.editingInvoice.total || 0;
        }
    },

    /**
     * Toggle เงินโอนใน Modal แก้ไข
     */
    toggleEditPaymentTransfer() {
        const checkbox = document.getElementById('editPaymentTransfer');
        const amountInput = document.getElementById('editPaymentTransferAmount');
        amountInput.disabled = !checkbox.checked;
        if (checkbox.checked && !amountInput.value) {
            amountInput.value = this.editingInvoice.total || 0;
        }
    },

    /**
     * Render รายการสินค้าใน Modal แก้ไข
     */
    renderEditInvoiceItems() {
        const container = document.getElementById('editInvoiceItemsContainer');
        if (!container) return;

        const items = this.editingInvoice.items || [];

        if (items.length === 0) {
            // Add default item
            items.push({
                id: 1,
                description: 'รายการสินค้า/บริการ',
                quantity: 1,
                price: (this.editingInvoice.total || 0) / 1.07
            });
            this.editingInvoice.items = items;
        }

        container.innerHTML = items.map((item, index) => `
            <div class="edit-item-row" style="display: flex; gap: 10px; margin-bottom: 8px; align-items: center;">
                <span style="width: 30px; text-align: center;">${index + 1}</span>
                <input type="text" value="${item.description || ''}" 
                    placeholder="รายละเอียด" 
                    style="flex: 2;"
                    onchange="App.updateEditItem(${index}, 'description', this.value)">
                <input type="number" value="${item.quantity || 1}" 
                    placeholder="จำนวน" 
                    style="width: 80px;"
                    onchange="App.updateEditItem(${index}, 'quantity', this.value)">
                <input type="number" value="${(item.price || 0).toFixed(2)}" 
                    placeholder="ราคา/หน่วย" 
                    style="width: 120px;"
                    step="0.01"
                    onchange="App.updateEditItem(${index}, 'price', this.value)">
                <span style="width: 100px; text-align: right;">${Invoice.formatCurrency((item.quantity || 1) * (item.price || 0))}</span>
                <button class="btn btn-danger btn-sm" onclick="App.removeEditItem(${index})" ${items.length <= 1 ? 'disabled' : ''}>✕</button>
            </div>
        `).join('');

        this.calculateEditInvoiceTotals();
    },

    /**
     * อัปเดตรายการสินค้าใน Modal แก้ไข
     */
    updateEditItem(index, field, value) {
        if (!this.editingInvoice.items[index]) return;

        if (field === 'quantity' || field === 'price') {
            this.editingInvoice.items[index][field] = parseFloat(value) || 0;
        } else {
            this.editingInvoice.items[index][field] = value;
        }

        this.renderEditInvoiceItems();
    },

    /**
     * เพิ่มรายการสินค้าใน Modal แก้ไข
     */
    addEditInvoiceItem() {
        if (!this.editingInvoice.items) this.editingInvoice.items = [];

        this.editingInvoice.items.push({
            id: Date.now(),
            description: '',
            quantity: 1,
            price: 0
        });

        this.renderEditInvoiceItems();
    },

    /**
     * ลบรายการสินค้าใน Modal แก้ไข
     */
    removeEditItem(index) {
        if (this.editingInvoice.items.length <= 1) return;

        this.editingInvoice.items.splice(index, 1);
        this.renderEditInvoiceItems();
    },

    /**
     * คำนวณยอดรวมใน Modal แก้ไข (ใช้ VAT Inclusive เหมือนหน้าหลัก)
     */
    calculateEditInvoiceTotals() {
        const items = this.editingInvoice.items || [];

        // ใช้ Invoice.calculateTotal ซึ่งคำนวณแบบ VAT Inclusive
        // ราคาที่กรอกเป็นราคารวม VAT แล้ว
        const totals = Invoice.calculateTotal(items);

        document.getElementById('editInvoiceSubtotal').value = totals.subtotal.toFixed(2);
        document.getElementById('editInvoiceVat').value = totals.vat.toFixed(2);
        document.getElementById('editInvoiceTotal').value = totals.total.toFixed(2);

        this.editingInvoice.subtotal = totals.subtotal;
        this.editingInvoice.vat = totals.vat;
        this.editingInvoice.total = totals.total;
    },

    /**
     * บันทึกใบกำกับภาษีที่แก้ไข (รวมถึงส่งข้อมูลไป Google Sheets)
     */
    async saveEditedInvoice(shouldPrint = false) {
        const originalInvoiceNumber = document.getElementById('editInvoiceNumber').value;
        const newInvoiceNumber = document.getElementById('editInvoiceNo').value.trim();

        // Validate new invoice number
        if (!newInvoiceNumber) {
            this.showToast('กรุณากรอกเลขที่ใบกำกับภาษี', 'error');
            return;
        }

        const invoices = Storage.getInvoices();
        const index = invoices.findIndex(inv => inv.invoiceNumber === originalInvoiceNumber);

        if (index === -1) {
            this.showToast('ไม่พบใบกำกับภาษี', 'error');
            return;
        }

        // Check if new invoice number already exists (if changed)
        if (newInvoiceNumber !== originalInvoiceNumber) {
            const duplicateIndex = invoices.findIndex(inv => inv.invoiceNumber === newInvoiceNumber);
            if (duplicateIndex !== -1) {
                this.showToast('เลขที่ใบกำกับนี้มีอยู่แล้ว กรุณาใช้เลขอื่น', 'error');
                return;
            }
        }

        // Show loading popup
        this.showLoading(shouldPrint ? 'กำลังบันทึก และเตรียมพิมพ์...' : 'กำลังบันทึกข้อมูล...');

        // Get payment info from form
        const paymentCash = document.getElementById('editPaymentCash').checked;
        const paymentCashAmount = parseFloat(document.getElementById('editPaymentCashAmount').value) || 0;
        const paymentTransfer = document.getElementById('editPaymentTransfer').checked;
        const paymentTransferAmount = parseFloat(document.getElementById('editPaymentTransferAmount').value) || 0;

        // Get customer info
        const customerName = document.getElementById('editInvoiceCustomerName').value;
        const customerAddress = document.getElementById('editInvoiceCustomerAddress').value;
        const customerTaxId = document.getElementById('editInvoiceCustomerTaxId').value;
        const customerPhone = document.getElementById('editInvoiceCustomerPhone').value;
        const invoiceDate = document.getElementById('editInvoiceDate').value;

        // Update items from form
        this.updateEditItemsFromForm();

        // Update invoice data (including new invoice number)
        invoices[index].invoiceNumber = newInvoiceNumber;  // Updated to new invoice number
        invoices[index].date = invoiceDate;
        invoices[index].customerName = customerName;
        invoices[index].customerAddress = customerAddress;
        invoices[index].customerTaxId = customerTaxId;
        invoices[index].customerPhone = customerPhone;
        invoices[index].items = this.editingInvoice.items || [];
        invoices[index].subtotal = this.editingInvoice.subtotal || 0;
        invoices[index].vat = this.editingInvoice.vat || 0;
        invoices[index].total = this.editingInvoice.total || 0;
        invoices[index].thaiText = Invoice.amountToThaiText(invoices[index].total);
        invoices[index].payment = {
            cash: paymentCash,
            cashAmount: paymentCashAmount,
            transfer: paymentTransfer,
            transferAmount: paymentTransferAmount
        };
        invoices[index].updatedAt = new Date().toISOString();

        // Store original invoice number if it was changed (for Sheets update)
        if (newInvoiceNumber !== originalInvoiceNumber) {
            invoices[index].originalInvoiceNumber = originalInvoiceNumber;
        }

        // Save back to localStorage
        Storage.saveInvoices(invoices);

        // Sync customer data if there are changes
        this.syncCustomerFromInvoice(customerName, customerAddress, customerTaxId, customerPhone);

        // Send update to Google Sheets
        try {
            const settings = Storage.getSettings();
            if (settings.scriptUrl) {
                // Update customer in Google Sheets
                await SheetsAPI.updateCustomer({
                    id: 'CUST-' + Date.now(),
                    name: customerName,
                    address: customerAddress,
                    taxId: customerTaxId,
                    phone: customerPhone,
                    email: ''
                });

                // Update invoice in Google Sheets
                await SheetsAPI.updateInvoice({
                    invoiceNumber: newInvoiceNumber,
                    originalInvoiceNumber: originalInvoiceNumber,  // For finding the row to update
                    date: invoiceDate,
                    customerName: customerName,
                    customerAddress: customerAddress,
                    customerTaxId: customerTaxId,
                    subtotal: invoices[index].subtotal,
                    vat: invoices[index].vat,
                    total: invoices[index].total,
                    items: JSON.stringify(invoices[index].items)
                });

                if (newInvoiceNumber !== originalInvoiceNumber) {
                    this.showToast(`✅ อัปเดตเลขใบกำกับจาก ${originalInvoiceNumber} เป็น ${newInvoiceNumber} สำเร็จ`, 'success');
                }
            }
        } catch (error) {
            console.warn('Could not sync to Google Sheets:', error);
        }

        // Hide loading
        this.hideLoading();

        // Print if requested
        if (shouldPrint) {
            const data = {
                invoiceNumber: newInvoiceNumber,
                date: invoiceDate,
                customer: {
                    name: customerName,
                    address: customerAddress,
                    taxId: customerTaxId,
                    phone: customerPhone
                },
                branchType: this.editingInvoice.branchType || 'hq',
                branchNumber: this.editingInvoice.branchNumber || '',
                payment: {
                    cash: paymentCash,
                    cashAmount: paymentCashAmount,
                    transfer: paymentTransfer,
                    transferAmount: paymentTransferAmount
                },
                items: this.editingInvoice.items || [],
                subtotal: this.editingInvoice.subtotal || 0,
                vat: this.editingInvoice.vat || 0,
                total: this.editingInvoice.total || 0,
                includeCopy: false
            };

            Invoice.print(data);
        }

        this.closeModal('editInvoiceModal');
        this.loadHistory();
        this.showToast('บันทึกการแก้ไขสำเร็จ', 'success');
    },

    /**
     * บันทึก + พิมพ์ใบกำกับภาษีที่แก้ไข
     */
    async saveAndPrintEditedInvoice() {
        await this.saveEditedInvoice(true);
    },

    /**
     * Sync ข้อมูลลูกค้าจากใบกำกับภาษี (อัปเดตหรือเพิ่มใหม่)
     */
    syncCustomerFromInvoice(name, address, taxId, phone) {
        if (!name) return;

        const customers = Storage.getCustomers();

        // Try to find existing customer by taxId or name
        let existingIndex = -1;
        if (taxId) {
            existingIndex = customers.findIndex(c => c.taxId === taxId);
        }
        if (existingIndex === -1) {
            existingIndex = customers.findIndex(c => c.name === name);
        }

        if (existingIndex !== -1) {
            // Update existing customer
            customers[existingIndex].name = name;
            customers[existingIndex].address = address;
            customers[existingIndex].taxId = taxId;
            customers[existingIndex].phone = phone;
        } else {
            // Add new customer
            customers.push({
                id: 'CUST-' + Date.now(),
                name: name,
                address: address,
                taxId: taxId,
                phone: phone,
                email: ''
            });
        }

        localStorage.setItem('customers', JSON.stringify(customers));
        this.customers = customers;
    },

    /**
     * พิมพ์ใบกำกับภาษีจาก Modal แก้ไข (ใช้ข้อมูลที่แก้ไขแล้ว)
     */
    printEditedInvoice() {
        if (!this.editingInvoice) {
            this.showToast('ไม่พบข้อมูลใบกำกับภาษี', 'error');
            return;
        }

        // Force update items from form inputs before printing
        this.updateEditItemsFromForm();

        // Get current form values
        const invoiceNumber = document.getElementById('editInvoiceNo').value;
        const date = document.getElementById('editInvoiceDate').value;
        const customerName = document.getElementById('editInvoiceCustomerName').value;
        const customerAddress = document.getElementById('editInvoiceCustomerAddress').value;
        const customerTaxId = document.getElementById('editInvoiceCustomerTaxId').value;
        const customerPhone = document.getElementById('editInvoiceCustomerPhone').value;

        // Get items and recalculate totals
        const items = this.editingInvoice.items || [];
        const subtotal = this.editingInvoice.subtotal || 0;
        const vat = this.editingInvoice.vat || 0;
        const total = this.editingInvoice.total || 0;

        // Get payment from form (current values)
        const paymentCash = document.getElementById('editPaymentCash').checked;
        const paymentCashAmount = parseFloat(document.getElementById('editPaymentCashAmount').value) || 0;
        const paymentTransfer = document.getElementById('editPaymentTransfer').checked;
        const paymentTransferAmount = parseFloat(document.getElementById('editPaymentTransferAmount').value) || 0;

        const payment = {
            cash: paymentCash,
            cashAmount: paymentCashAmount,
            transfer: paymentTransfer,
            transferAmount: paymentTransferAmount
        };

        const data = {
            invoiceNumber: invoiceNumber,
            date: date,
            customer: {
                name: customerName,
                address: customerAddress,
                taxId: customerTaxId,
                phone: customerPhone
            },
            branchType: this.editingInvoice.branchType || 'hq',
            branchNumber: this.editingInvoice.branchNumber || '',
            payment: payment,
            items: items,
            subtotal: subtotal,
            vat: vat,
            total: total,
            includeCopy: false
        };

        Invoice.print(data);
    },

    /**
     * อัปเดต items จาก form inputs ปัจจุบัน
     */
    updateEditItemsFromForm() {
        const container = document.getElementById('editInvoiceItemsContainer');
        if (!container) return;

        const rows = container.querySelectorAll('.edit-item-row');
        const items = [];

        rows.forEach((row, index) => {
            const inputs = row.querySelectorAll('input');
            if (inputs.length >= 3) {
                const description = inputs[0].value || '';
                const quantity = parseFloat(inputs[1].value) || 1;
                const price = parseFloat(inputs[2].value) || 0;

                items.push({
                    id: index + 1,
                    description: description,
                    quantity: quantity,
                    price: price
                });
            }
        });

        this.editingInvoice.items = items;
        this.calculateEditInvoiceTotals();
    },

    /**
     * ดูรายละเอียดใบกำกับภาษี
     */
    viewInvoice(invoiceNumber) {
        const invoices = Storage.getInvoices();
        const invoice = invoices.find(inv => inv.invoiceNumber === invoiceNumber);

        if (!invoice) {
            this.showToast('ไม่พบใบกำกับภาษี', 'error');
            return;
        }

        this.currentViewInvoice = invoice;

        const data = {
            invoiceNumber: invoice.invoiceNumber,
            date: invoice.date,
            customer: {
                name: invoice.customerName,
                address: invoice.customerAddress || '',
                taxId: invoice.customerTaxId || '',
                phone: invoice.customerPhone || ''
            },
            branchType: invoice.branchType || 'hq',
            branchNumber: invoice.branchNumber || '',
            payment: invoice.payment || {
                cash: false,
                cashAmount: 0,
                transfer: false,
                transferAmount: 0
            },
            items: invoice.items || [],
            subtotal: invoice.subtotal,
            vat: invoice.vat,
            total: invoice.total
        };

        document.getElementById('invoiceDetailPreview').innerHTML = Invoice.generatePreviewHTML(data);
        this.openModal('invoiceDetailModal');
    },

    /**
     * พิมพ์ซ้ำจากประวัติ (รูปแบบเหมือนหน้าหลัก)
     */
    reprintFromHistory(invoiceNumber) {
        const invoices = Storage.getInvoices();
        const invoice = invoices.find(inv => inv.invoiceNumber === invoiceNumber);

        if (!invoice) {
            this.showToast('ไม่พบใบกำกับภาษี', 'error');
            return;
        }

        // Get includeCopy option from main checkbox or default to false
        const includeCopyEl = document.getElementById('historyIncludeCopy');
        const includeCopy = includeCopyEl ? includeCopyEl.checked : false;

        const data = {
            invoiceNumber: invoice.invoiceNumber,
            date: invoice.date,
            customer: {
                name: invoice.customerName,
                address: invoice.customerAddress || '',
                taxId: invoice.customerTaxId || '',
                phone: invoice.customerPhone || ''
            },
            branchType: invoice.branchType || 'hq',
            branchNumber: invoice.branchNumber || '',
            payment: invoice.payment || {
                cash: true,
                cashAmount: invoice.total,
                transfer: false,
                transferAmount: 0
            },
            items: invoice.items || [{
                id: 1,
                description: 'รายการสินค้า/บริการ',
                quantity: 1,
                price: invoice.total / 1.07
            }],
            subtotal: invoice.subtotal,
            vat: invoice.vat,
            total: invoice.total,
            includeCopy: includeCopy
        };

        Invoice.print(data);
    },

    /**
     * พิมพ์ซ้ำจาก Modal
     */
    reprintInvoice() {
        if (this.currentViewInvoice) {
            this.reprintFromHistory(this.currentViewInvoice.invoiceNumber);
        }
    },

    /**
     * เปิด Modal
     */
    openModal(modalId) {
        document.getElementById(modalId).classList.add('active');
    },

    /**
     * ปิด Modal
     */
    closeModal(modalId) {
        document.getElementById(modalId).classList.remove('active');
    },

    /**
     * แสดง Toast notification
     */
    showToast(message, type = 'info') {
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.textContent = message;
        document.body.appendChild(toast);

        setTimeout(() => {
            toast.remove();
        }, 3000);
    },

    /**
     * แสดง Loading Modal
     */
    showLoading(message = 'รอสักครู่...') {
        const modal = document.getElementById('loadingModal');
        const msgEl = document.getElementById('loadingMessage');
        if (modal && msgEl) {
            msgEl.textContent = message;
            modal.classList.add('active');
        }
    },

    /**
     * ซ่อน Loading Modal
     */
    hideLoading() {
        const modal = document.getElementById('loadingModal');
        if (modal) {
            modal.classList.remove('active');
        }
    },

    /**
     * เปิด Email Modal
     */
    openEmailModal() {
        // ตรวจสอบว่ามีรายการสินค้าหรือไม่
        if (this.items.length === 0 || this.items.every(item => item.price === 0)) {
            this.showToast('กรุณาเพิ่มรายการสินค้าก่อนส่ง email', 'error');
            return;
        }

        // Reset ข้อมูลชั่วคราว (เพื่อให้ใช้ข้อมูลจากฟอร์มแทน)
        this.emailInvoiceData = null;
        this.emailTempSignature = null;

        const totals = Invoice.calculateTotal(this.items);
        const invoiceNumber = document.getElementById('invoiceNumber').value;
        const customerName = document.getElementById('customerName').value;

        // แสดงข้อมูลใน modal
        document.getElementById('emailInvoiceNumber').textContent = invoiceNumber;
        document.getElementById('emailCustomerName').textContent = customerName || '-';
        document.getElementById('emailTotal').textContent = Invoice.formatCurrency(totals.total) + ' บาท';

        // ดึง email ลูกค้าจากข้อมูลที่เลือก (ถ้ามี)
        let customerEmail = '';
        if (this.selectedCustomer && this.selectedCustomer.email) {
            customerEmail = this.selectedCustomer.email;
        }
        document.getElementById('recipientEmail').value = customerEmail;

        // Reset signature preview - แสดงลายเซ็นจาก storage ถ้ามี
        const savedSignature = Storage.getSignature();
        if (savedSignature) {
            document.getElementById('emailSignaturePreview').src = savedSignature;
            document.getElementById('emailSignaturePreview').style.display = 'block';
            document.getElementById('emailSignaturePlaceholder').style.display = 'none';
        } else {
            document.getElementById('emailSignaturePreview').style.display = 'none';
            document.getElementById('emailSignaturePlaceholder').style.display = 'block';
        }

        this.openModal('emailModal');
    },

    /**
     * ส่งใบกำกับภาษีทาง Email
     */
    async sendInvoiceEmail() {
        const recipientEmail = document.getElementById('recipientEmail').value.trim();

        // ตรวจสอบ email
        if (!recipientEmail) {
            this.showToast('กรุณาระบุ email ลูกค้า', 'error');
            return;
        }

        // ตรวจสอบรูปแบบ email
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(recipientEmail)) {
            this.showToast('รูปแบบ email ไม่ถูกต้อง', 'error');
            return;
        }

        const settings = Storage.getSettings();
        if (!settings.scriptUrl) {
            this.showToast('กรุณาตั้งค่า Google Apps Script URL ก่อน', 'error');
            return;
        }

        // ปิด modal และแสดง loading
        this.closeModal('emailModal');
        this.showLoading('กำลังส่ง email...');

        try {
            const totals = Invoice.calculateTotal(this.items);
            const company = Storage.getCompany();
            const logo = Storage.getLogo();
            const signature = Storage.getSignature();

            // Get branch type
            const branchType = document.querySelector('input[name="branchType"]:checked')?.value || 'hq';
            const branchNumber = document.getElementById('branchNumber').value || '';

            // Get payment info
            const paymentCash = document.getElementById('paymentCash').checked;
            const paymentCashAmount = parseFloat(document.getElementById('paymentCashAmount').value) || 0;
            const paymentTransfer = document.getElementById('paymentTransfer').checked;
            const paymentTransferAmount = parseFloat(document.getElementById('paymentTransferAmount').value) || 0;

            const invoiceData = {
                invoiceNumber: document.getElementById('invoiceNumber').value,
                date: document.getElementById('invoiceDate').value,
                customer: {
                    name: document.getElementById('customerName').value,
                    address: document.getElementById('customerAddress').value,
                    taxId: document.getElementById('customerTaxId').value,
                    phone: document.getElementById('customerPhone').value
                },
                branchType: branchType,
                branchNumber: branchNumber,
                payment: {
                    cash: paymentCash,
                    cashAmount: paymentCashAmount,
                    transfer: paymentTransfer,
                    transferAmount: paymentTransferAmount
                },
                items: this.items,
                subtotal: totals.subtotal,
                vat: totals.vat,
                total: totals.total
            };

            // สร้าง HTML สำหรับ PDF
            const invoiceHtml = Invoice.generatePDFInlineContent(invoiceData, company, logo, signature);

            // ส่งข้อมูลไป Apps Script
            const emailPayload = {
                customerEmail: recipientEmail,
                customerName: invoiceData.customer.name,
                invoiceNumber: invoiceData.invoiceNumber,
                invoiceHtml: invoiceHtml,
                companyName: company.name || 'ร้านระเบียงบัว ศรีสมาน',
                total: totals.total
            };

            const result = await SheetsAPI.sendInvoiceEmail(emailPayload);

            this.hideLoading();

            if (result.success) {
                this.showToast(`ส่ง email ไปยัง ${recipientEmail} สำเร็จ!`, 'success');

                // แสดงข้อความเพิ่มเติมถ้ามี
                if (result.note) {
                    setTimeout(() => {
                        this.showToast(result.note, 'info');
                    }, 2000);
                }
            } else {
                this.showToast(result.error || 'ไม่สามารถส่ง email ได้', 'error');
            }

        } catch (error) {
            this.hideLoading();
            console.error('Error sending email:', error);
            this.showToast('ไม่สามารถส่ง email ได้: ' + error.message, 'error');
        }
    },

    // ตัวแปรเก็บข้อมูลใบกำกับสำหรับส่ง email จาก history
    emailInvoiceData: null,
    emailTempSignature: null,

    /**
     * เปิด Email Modal จากหน้าประวัติใบกำกับ
     */
    openEmailModalFromHistory(invoiceNumber) {
        // ดึงข้อมูลใบกำกับจาก localStorage
        const invoices = Storage.getInvoices();
        const invoice = invoices.find(inv => inv.invoiceNumber === invoiceNumber);

        if (!invoice) {
            this.showToast('ไม่พบใบกำกับภาษีนี้', 'error');
            return;
        }

        // เก็บข้อมูลใบกำกับไว้ใช้ตอนส่ง
        this.emailInvoiceData = invoice;
        this.emailTempSignature = null;

        // แสดงข้อมูลใน modal
        document.getElementById('emailInvoiceNumber').textContent = invoice.invoiceNumber;
        document.getElementById('emailCustomerName').textContent = invoice.customerName || '-';
        document.getElementById('emailTotal').textContent = Invoice.formatCurrency(invoice.total) + ' บาท';

        // ดึง email ลูกค้าจากข้อมูลลูกค้า (ถ้ามี)
        let customerEmail = '';
        const customer = this.customers.find(c => c.name === invoice.customerName);
        if (customer && customer.email) {
            customerEmail = customer.email;
        }
        document.getElementById('recipientEmail').value = customerEmail;

        // Reset signature preview
        const savedSignature = Storage.getSignature();
        if (savedSignature) {
            document.getElementById('emailSignaturePreview').src = savedSignature;
            document.getElementById('emailSignaturePreview').style.display = 'block';
            document.getElementById('emailSignaturePlaceholder').style.display = 'none';
        } else {
            document.getElementById('emailSignaturePreview').style.display = 'none';
            document.getElementById('emailSignaturePlaceholder').style.display = 'block';
        }

        this.openModal('emailModal');
    },

    /**
     * Handle signature upload for email
     */
    handleEmailSignatureUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (event) => {
            const base64 = event.target.result;
            this.emailTempSignature = base64;

            document.getElementById('emailSignaturePreview').src = base64;
            document.getElementById('emailSignaturePreview').style.display = 'block';
            document.getElementById('emailSignaturePlaceholder').style.display = 'none';

            this.showToast('เลือกลายเซ็นสำหรับ email แล้ว', 'success');
        };
        reader.readAsDataURL(file);
    },

    /**
     * ส่งใบกำกับภาษีทาง Email (ใช้ได้ทั้งจากหน้าสร้างและหน้าประวัติ)
     */
    async sendInvoiceEmailFromHistory() {
        const recipientEmail = document.getElementById('recipientEmail').value.trim();

        // ตรวจสอบ email
        if (!recipientEmail) {
            this.showToast('กรุณาระบุ email ลูกค้า', 'error');
            return;
        }

        // ตรวจสอบรูปแบบ email
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(recipientEmail)) {
            this.showToast('รูปแบบ email ไม่ถูกต้อง', 'error');
            return;
        }

        const settings = Storage.getSettings();
        if (!settings.scriptUrl) {
            this.showToast('กรุณาตั้งค่า Google Apps Script URL ก่อน', 'error');
            return;
        }

        // ปิด modal และแสดง loading
        this.closeModal('emailModal');
        this.showLoading('กำลังส่ง email...');

        try {
            const company = Storage.getCompany();
            const logo = Storage.getLogo();
            // ใช้ลายเซ็นที่เลือกใหม่ หรือจาก storage
            const signature = this.emailTempSignature || Storage.getSignature();

            let invoiceData;
            let totals;

            // ตรวจสอบว่าส่งจากหน้าไหน
            if (this.emailInvoiceData) {
                // ส่งจากหน้าประวัติ - ใช้ข้อมูลที่เก็บไว้
                const inv = this.emailInvoiceData;
                invoiceData = {
                    invoiceNumber: inv.invoiceNumber,
                    date: inv.date,
                    customer: {
                        name: inv.customerName,
                        address: inv.customerAddress || '',
                        taxId: inv.customerTaxId || '',
                        phone: inv.customerPhone || ''
                    },
                    branchType: inv.branchType || 'hq',
                    branchNumber: inv.branchNumber || '',
                    payment: inv.payment || { cash: true, cashAmount: inv.total, transfer: false, transferAmount: 0 },
                    items: typeof inv.items === 'string' ? JSON.parse(inv.items) : (inv.items || []),
                    subtotal: inv.subtotal || 0,
                    vat: inv.vat || 0,
                    total: inv.total || 0
                };
                totals = { total: inv.total };
            } else {
                // ส่งจากหน้าสร้าง - ใช้ข้อมูลจากฟอร์ม
                totals = Invoice.calculateTotal(this.items);
                const branchType = document.querySelector('input[name="branchType"]:checked')?.value || 'hq';
                const branchNumber = document.getElementById('branchNumber').value || '';
                const paymentCash = document.getElementById('paymentCash').checked;
                const paymentCashAmount = parseFloat(document.getElementById('paymentCashAmount').value) || 0;
                const paymentTransfer = document.getElementById('paymentTransfer').checked;
                const paymentTransferAmount = parseFloat(document.getElementById('paymentTransferAmount').value) || 0;

                invoiceData = {
                    invoiceNumber: document.getElementById('invoiceNumber').value,
                    date: document.getElementById('invoiceDate').value,
                    customer: {
                        name: document.getElementById('customerName').value,
                        address: document.getElementById('customerAddress').value,
                        taxId: document.getElementById('customerTaxId').value,
                        phone: document.getElementById('customerPhone').value
                    },
                    branchType: branchType,
                    branchNumber: branchNumber,
                    payment: {
                        cash: paymentCash,
                        cashAmount: paymentCashAmount,
                        transfer: paymentTransfer,
                        transferAmount: paymentTransferAmount
                    },
                    items: this.items,
                    subtotal: totals.subtotal,
                    vat: totals.vat,
                    total: totals.total
                };
            }

            // สร้าง HTML สำหรับ PDF
            const invoiceHtml = Invoice.generatePDFInlineContent(invoiceData, company, logo, signature);

            // ส่งข้อมูลไป Apps Script
            const emailPayload = {
                customerEmail: recipientEmail,
                customerName: invoiceData.customer.name,
                invoiceNumber: invoiceData.invoiceNumber,
                invoiceHtml: invoiceHtml,
                companyName: company.name || 'ร้านระเบียงบัว ศรีสมาน',
                total: invoiceData.total
            };

            const result = await SheetsAPI.sendInvoiceEmail(emailPayload);

            this.hideLoading();

            // ล้างข้อมูลชั่วคราว
            this.emailInvoiceData = null;
            this.emailTempSignature = null;

            if (result.success) {
                this.showToast(`ส่ง email ไปยัง ${recipientEmail} สำเร็จ!`, 'success');

                if (result.note) {
                    setTimeout(() => {
                        this.showToast(result.note, 'info');
                    }, 2000);
                }
            } else {
                this.showToast(result.error || 'ไม่สามารถส่ง email ได้', 'error');
            }

        } catch (error) {
            this.hideLoading();
            this.emailInvoiceData = null;
            this.emailTempSignature = null;
            console.error('Error sending email:', error);
            this.showToast('ไม่สามารถส่ง email ได้: ' + error.message, 'error');
        }
    }
};

// Initialize app when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    App.init();
});

// Export for use
window.App = App;
