document.addEventListener('DOMContentLoaded', () => {

    // --- Global Variables ---
    let debts = [];
    let currentPage = 1;
    const itemsPerPage = 9; // แสดงผล 9 รายการต่อหน้า

    // --- DOM Elements ---
    const views = {
        dashboard: document.getElementById('dashboard-view'),
        addDebt: document.getElementById('add-debt-view'),
        detail: document.getElementById('detail-view'),
    };
    
    const navLinks = {
        brand: document.getElementById('nav-brand'),
        dashboard: document.getElementById('nav-dashboard'),
        addDebt: document.getElementById('nav-add-debt'),
    };

    const activeDebtList = document.getElementById('active-debt-list');
    const completedDebtList = document.getElementById('completed-debt-list');
    const paginationControls = document.getElementById('pagination-controls');
    const prevPageBtn = document.getElementById('prev-page-btn');
    const nextPageBtn = document.getElementById('next-page-btn');
    const pageIndicator = document.getElementById('page-indicator');
    const prevPageLi = document.getElementById('prev-page-li');
    const nextPageLi = document.getElementById('next-page-li');

    const dashboardSummary = document.getElementById('dashboard-summary');
    
    const debtForm = document.getElementById('debt-form');
    const formTitle = document.getElementById('form-title');
    const debtIdInput = document.getElementById('debt-id');
    const itemNameInput = document.getElementById('item-name');
    const creditorNameInput = document.getElementById('creditor-name');
    const principalInput = document.getElementById('principal');
    const interestRateInput = document.getElementById('interest-rate');
    const totalMonthsInput = document.getElementById('total-months');
    const paymentDayInput = document.getElementById('payment-day');
    const enableCustomPaymentCheck = document.getElementById('enable-custom-payment');
    const customPaymentGroup = document.getElementById('custom-payment-group');
    const customPaymentInput = document.getElementById('custom-payment');

    const detailTitle = document.getElementById('detail-title');
    const detailItemName = document.getElementById('detail-item-name');
    const detailCreditor = document.getElementById('detail-creditor');
    const detailSummary = document.getElementById('detail-summary');
    const paymentList = document.getElementById('payment-list');
    const backToDashboardBtn = document.getElementById('back-to-dashboard-btn');
    
    const exportExcelBtn = document.getElementById('export-excel-btn');
    const importExcelBtn = document.getElementById('import-excel-btn');
    const importExcelInput = document.getElementById('import-excel-input');

    const realtimeClockElement = document.getElementById('realtime-clock');


    // --- Utility Functions ---

    /**
     * Shows a specific view and hides others (SPA behavior)
     */
    const showView = (viewId) => {
        Object.values(views).forEach(view => view.style.display = 'none');
        views[viewId].style.display = 'block';

        // Update active nav link
        Object.values(navLinks).forEach(link => link.classList.remove('active'));
        if (viewId === 'dashboard') navLinks.dashboard.classList.add('active');
        if (viewId === 'addDebt') navLinks.addDebt.classList.add('active');
    };

    /**
     * Loads debts from localStorage
     */
    const loadDebts = () => {
        debts = JSON.parse(localStorage.getItem('debts')) || [];
    };

    /**
     * Saves debts to localStorage
     */
    const saveDebts = () => {
        localStorage.setItem('debts', JSON.stringify(debts));
    };

    /**
     * Calculates monthly payment for an amortized loan
     */
    const calculateAmortization = (P, i_percent, n) => {
        if (i_percent === 0) {
            return P / n;
        }
        const i_monthly = (i_percent / 100) / 12;
        const M = P * (i_monthly * Math.pow(1 + i_monthly, n)) / (Math.pow(1 + i_monthly, n) - 1);
        return M;
    };

    /**
     * Calculates total remaining balance for a debt
     */
    const getRemainingBalance = (debt) => {
        return debt.monthlyPayments
            .filter(p => p.status === 'unpaid')
            .reduce((sum, p) => sum + p.amount, 0);
    };

    /**
     * Formats number to currency (2 decimal places)
     */
    const formatCurrency = (num) => {
        return num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    };

    /**
     * Calculates days remaining until the next paymentDay
     */
    const getDaysRemaining = (paymentDay) => {
        const today = new Date();
        today.setHours(0, 0, 0, 0); 
        
        const currentYear = today.getFullYear();
        const currentMonth = today.getMonth();

        if (today.getDate() === paymentDay) {
            return 0; // ถึงกำหนดวันนี้
        }
        
        let nextPaymentDate = new Date(currentYear, currentMonth, paymentDay);
        
        if (nextPaymentDate.getTime() < today.getTime()) {
            nextPaymentDate = new Date(currentYear, currentMonth + 1, paymentDay);
        }
        
        const diffTime = nextPaymentDate.getTime() - today.getTime();
        return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    };

    /**
     * Updates the real-time clock element
     */
    const updateClock = () => {
        if (!realtimeClockElement) return;

        const now = new Date();
        
        const options = {
            year: 'numeric',
            month: 'long', 
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
            hour12: false 
        };
        
        const formattedTime = now.toLocaleString('th-TH', options);
        
        realtimeClockElement.textContent = formattedTime;
    };

    /**
     * Generates the HTML string for a single debt card
     * @param {object} debt - The debt object
     * @returns {string} HTML string
     */
    function generateDebtCardHtml(debt) {
        const remainingBalance = getRemainingBalance(debt);
        const totalDebtAmount = debt.monthlyPayments.reduce((sum, p) => sum + p.amount, 0);
        const paidAmount = totalDebtAmount - remainingBalance;
        
        const paidMonths = debt.monthlyPayments.filter(p => p.status === 'paid').length;
        const totalMonths = debt.totalMonths;
        const progressPercent = (totalDebtAmount > 0) ? (paidAmount / totalDebtAmount) * 100 : 0;

        let dueDateHtml = '';
        
        if (remainingBalance <= 0 && totalDebtAmount > 0) {
            // จ่ายครบหมดแล้ว
            dueDateHtml = `
                <p class="card-text mb-2">
                    <strong><i class="bi bi-check-circle-fill text-success"></i> ชำระครบแล้ว</strong>
                </p>`;
        } else {
            // ยังไม่ครบ, คำนวณวันคงเหลือ
            const daysRemaining = getDaysRemaining(debt.paymentDay);
            let colorClass = 'text-success'; // > 10 วัน
            let iconClass = 'bi bi-calendar-check';
            let statusText = `เหลืออีก ${daysRemaining} วัน`;

            if (daysRemaining === 0) {
                colorClass = 'text-danger fw-bold'; // วันนี้
                iconClass = 'bi bi-exclamation-triangle-fill';
                statusText = 'ถึงกำหนดชำระวันนี้!';
            } else if (daysRemaining <= 5) {
                colorClass = 'text-danger'; // <= 5 วัน
                iconClass = 'bi bi-calendar-x';
            } else if (daysRemaining <= 10) {
                colorClass = 'text-warning'; // <= 10 วัน
                iconClass = 'bi bi-calendar-event';
            }

            dueDateHtml = `
                <p class="card-text mb-2">
                    <strong><i class="${iconClass} ${colorClass}"></i> ${statusText}</strong>
                </p>`;
        }

        return `
            <div class="col-md-6 col-lg-4 mb-4">
                <div class="card shadow-sm h-100">
                    <div class="card-body">
                        <h5 class="card-title">${debt.itemName}</h5>
                        <h6 class="card-subtitle mb-2 text-muted">${debt.creditor}</h6>
                        <p class="card-text mb-1">
                            <strong>ยอดคงเหลือ:</strong> <span class="text-danger h5">${formatCurrency(remainingBalance)}</span> บาท
                        </p>
                        <p class="card-text">
                            <small>จ่ายแล้ว: ${formatCurrency(paidAmount)} / ${formatCurrency(totalDebtAmount)}</small>
                        </p>
                        
                        <p class="card-text mb-2">
                            <strong>สถานะ:</strong> <span class="text-success fw-bold">${paidMonths} / ${totalMonths}</span> เดือน
                        </p>

                        ${dueDateHtml}
                        
                        <div class="progress mb-3" style="height: 10px;">
                            <div class="progress-bar bg-success" role="progressbar" style="width: ${progressPercent}%" aria-valuenow="${progressPercent}" aria-valuemin="0" aria-valuemax="100"></div>
                        </div>
                        <button class="btn btn-primary btn-sm view-details-btn" data-id="${debt.id}"><i class="bi bi-search"></i> ดูรายละเอียด</button>
                        <button class="btn btn-outline-secondary btn-sm edit-debt-btn" data-id="${debt.id}"><i class="bi bi-pencil"></i> แก้ไข</button>
                        <button class="btn btn-outline-danger btn-sm delete-debt-btn" data-id="${debt.id}"><i class="bi bi-trash"></i> ลบ</button>
                    </div>
                </div>
            </div>
        `;
    }


    // --- Core Logic Functions ---

    /**
     * Renders the main dashboard view
     */
    const renderDashboard = () => {
        activeDebtList.innerHTML = '';
        completedDebtList.innerHTML = '';
        paginationControls.style.display = 'none'; // ซ่อนไว้ก่อน

        if (debts.length === 0) {
            activeDebtList.innerHTML = `<div class="col-12"><div class="alert alert-info">ยังไม่มีรายการหนี้... คลิก "เพิ่มรายการหนี้" เพื่อเริ่มต้น</div></div>`;
            completedDebtList.innerHTML = `<div class="col-12"><div class="alert alert-secondary">ยังไม่มีหนี้ที่ชำระครบ</div></div>`;
            dashboardSummary.innerHTML = '';
            return;
        }

        // 1. แยกหนี้เป็น Active และ Completed
        const activeDebts = [];
        const completedDebts = [];
        let totalPaid = 0; // ยอดจ่ายแล้วทั้งหมด
        let totalRemaining = 0; // ยอดคงเหลือ (เฉพาะหนี้ที่ Active)

        debts.forEach(debt => {
            const remaining = getRemainingBalance(debt);
            const totalDebtAmount = debt.monthlyPayments.reduce((sum, p) => sum + p.amount, 0);
            const paid = totalDebtAmount - remaining;
            
            totalPaid += paid; 

            // ตรวจสอบว่าหนี้มีมูลค่า (totalDebtAmount > 0)
            if (totalDebtAmount > 0) {
                if (remaining > 0) {
                    activeDebts.push(debt);
                    totalRemaining += remaining; // ยอดคงเหลือ คิดเฉพาะหนี้ที่ Active
                } else {
                    completedDebts.push(debt);
                }
            }
        });

        // 2. แสดงผล Summary Cards
        dashboardSummary.innerHTML = `
            <div class="col-md-4 mb-3">
                <div class="card text-white bg-danger shadow">
                    <div class="card-body">
                        <h5 class="card-title"><i class="bi bi-cash-stack"></i> ยอดหนี้คงเหลือ</h5>
                        <p class="card-text h3">${formatCurrency(totalRemaining)}</p>
                    </div>
                </div>
            </div>
            <div class="col-md-4 mb-3">
                <div class="card text-white bg-success shadow">
                    <div class="card-body">
                        <h5 class="card-title"><i class="bi bi-cash-coin"></i> ยอดที่จ่ายไปแล้ว</h5>
                        <p class="card-text h3">${formatCurrency(totalPaid)}</p>
                    </div>
                </div>
            </div>
            <div class="col-md-4 mb-3">
                <div class="card text-dark bg-warning shadow">
                    <div class="card-body">
                        <h5 class="card-title"><i class="bi bi-list-ol"></i> หนี้ที่ยังต้องชำระ</h5>
                        <p class="card-text h3">${activeDebts.length}</p>
                    </div>
                </div>
            </div>
        `;

        // 3. ตรรกะการแบ่งหน้า (Pagination)
        const totalPages = Math.ceil(activeDebts.length / itemsPerPage);
        
        if (currentPage > totalPages) {
            currentPage = totalPages;
        }
        if (currentPage < 1) {
            currentPage = 1;
        }
        
        const startIndex = (currentPage - 1) * itemsPerPage;
        const endIndex = currentPage * itemsPerPage;
        const paginatedActiveDebts = activeDebts.slice(startIndex, endIndex);

        // 4. แสดงผลหนี้ที่ Active (ที่แบ่งหน้าแล้ว)
        if (activeDebts.length === 0) {
             activeDebtList.innerHTML = `<div class="col-12"><div class="alert alert-info">ไม่มีหนี้ที่ยังต้องชำระ</div></div>`;
        } else {
            paginatedActiveDebts.forEach(debt => {
                activeDebtList.innerHTML += generateDebtCardHtml(debt);
            });
        }

        // 5. แสดงผลหนี้ที่จ่ายครบแล้ว (ทั้งหมด)
        if (completedDebts.length === 0) {
            completedDebtList.innerHTML = `<div class="col-12"><div class="alert alert-secondary">ยังไม่มีหนี้ที่ชำระครบ</div></div>`;
        } else {
            completedDebts.forEach(debt => {
                completedDebtList.innerHTML += generateDebtCardHtml(debt); 
            });
        }

        // 6. อัปเดตปุ่ม Pagination
        if (totalPages > 1) {
            paginationControls.style.display = 'flex'; 
            pageIndicator.textContent = `Page ${currentPage} / ${totalPages}`;
            
            prevPageLi.classList.toggle('disabled', currentPage === 1);
            nextPageLi.classList.toggle('disabled', currentPage === totalPages);
            
            prevPageBtn.style.pointerEvents = (currentPage === 1) ? 'none' : 'auto';
            nextPageBtn.style.pointerEvents = (currentPage === totalPages) ? 'none' : 'auto';
        }

        addDashboardListeners();
    };

    /**
     * Renders the detail view for a specific debt
     */
    const renderDetailView = (debtId) => {
        const debt = debts.find(d => d.id === debtId);
        if (!debt) return;

        detailTitle.textContent = `รายละเอียด: ${debt.itemName}`;
        detailItemName.textContent = debt.itemName;
        detailCreditor.textContent = `เจ้าหนี้: ${debt.creditor} (จ่ายทุกวันที่ ${debt.paymentDay})`;
        
        const remainingBalance = getRemainingBalance(debt);
        const totalDebtAmount = debt.monthlyPayments.reduce((sum, p) => sum + p.amount, 0);
        const paidAmount = totalDebtAmount - remainingBalance;
        const paidMonths = debt.monthlyPayments.filter(p => p.status === 'paid').length;

        detailSummary.innerHTML = `
            <div class="col-md-4"><strong>ยอดคงเหลือ:</strong> <span class="text-danger h5">${formatCurrency(remainingBalance)}</span></div>
            <div class="col-md-4"><strong>จ่ายแล้ว:</strong> <span class="text-success h5">${formatCurrency(paidAmount)}</span></div>
            <div class="col-md-4"><strong>สถานะ:</strong> ${paidMonths} / ${debt.totalMonths} เดือน</div>
        `;

        paymentList.innerHTML = '';
        debt.monthlyPayments.forEach(p => {
            const isPaid = p.status === 'paid';
            const statusClass = isPaid ? 'payment-paid' : 'payment-unpaid';
            const statusText = isPaid ? 'จ่ายแล้ว' : 'ยังไม่จ่าย';
            const btnClass = isPaid ? 'btn-outline-warning' : 'btn-outline-success';
            const btnIcon = isPaid ? 'bi-x-circle' : 'bi-check-circle';
            const btnText = isPaid ? 'ยกเลิก' : 'จ่ายแล้ว';

            const itemHtml = `
                <div class="list-group-item d-flex justify-content-between align-items-center payment-item ${statusClass}">
                    <div>
                        <h6 class="mb-1 payment-month">เดือนที่ ${p.month}</h6>
                        <p class="mb-0">ยอดชำระ: <strong>${formatCurrency(p.amount)}</strong> บาท</p>
                        <small class="payment-status-text">สถานะ: ${statusText}</small>
                    </div>
                    <button class="btn ${btnClass} toggle-payment-btn" data-debt-id="${debt.id}" data-month="${p.month}">
                        <i class="bi ${btnIcon}"></i> ${btnText}
                    </button>
                </div>
            `;
            paymentList.innerHTML += itemHtml;
        });

        document.querySelectorAll('.toggle-payment-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const debtId = btn.dataset.debtId;
                const month = parseInt(btn.dataset.month);
                togglePaymentStatus(debtId, month);
            });
        });

        showView('detail');
    };

    /**
     * Clears the form fields for a new entry
     */
    const clearForm = () => {
        debtForm.reset();
        debtIdInput.value = '';
        formTitle.textContent = 'เพิ่มรายการหนี้ใหม่';
        interestRateInput.disabled = false;
        customPaymentGroup.style.display = 'none';
        customPaymentInput.value = '';
        enableCustomPaymentCheck.checked = false;
    };

    /**
     * Populates the form for editing an existing debt
     */
    const populateFormForEdit = (debtId) => {
        const debt = debts.find(d => d.id === debtId);
        if (!debt) return;

        clearForm();
        formTitle.textContent = 'แก้ไขรายการหนี้';
        debtIdInput.value = debt.id;
        itemNameInput.value = debt.itemName;
        creditorNameInput.value = debt.creditor;
        principalInput.value = debt.principal;
        interestRateInput.value = debt.interestRate;
        totalMonthsInput.value = debt.totalMonths;
        paymentDayInput.value = debt.paymentDay;
        
        if (debt.customPayment) {
            enableCustomPaymentCheck.checked = true;
            customPaymentGroup.style.display = 'block';
            customPaymentInput.value = debt.customPayment;
            interestRateInput.disabled = true;
        }

        showView('addDebt');
    };

    // --- Event Handlers ---

    /**
     * Handles navigation link clicks
     */
    const handleNavigation = (e) => {
        e.preventDefault();
        const targetId = e.currentTarget.id;

        if (targetId === 'nav-dashboard' || targetId === 'nav-brand') {
            currentPage = 1; 
            renderDashboard();
            showView('dashboard');
        } else if (targetId === 'nav-add-debt') {
            clearForm();
            showView('addDebt');
        }
    };

    /**
     * Handles the debt form submission (Add or Edit)
     */
    const handleFormSubmit = (e) => {
        e.preventDefault();

        const id = debtIdInput.value || crypto.randomUUID();
        const principal = parseFloat(principalInput.value);
        let interestRate = parseFloat(interestRateInput.value);
        const totalMonths = parseInt(totalMonthsInput.value);
        const customPayment = parseFloat(customPaymentInput.value);
        const isCustomPayment = enableCustomPaymentCheck.checked && customPayment > 0;

        let monthlyAmount;
        let finalInterestRate = interestRate;
        let finalCustomPayment = null;

        if (principal <= 0 || totalMonths <= 0) {
            Swal.fire('ข้อมูลไม่ถูกต้อง', 'ยอดเงินต้น และ จำนวนเดือน ต้องมากกว่า 0', 'error');
            return;
        }

        if (isCustomPayment) {
            if (customPayment <= 0) {
                Swal.fire('ข้อมูลไม่ถูกต้อง', 'จำนวนเงินที่จ่ายต่อเดือน ต้องมากกว่า 0', 'error');
                return;
            }
            monthlyAmount = customPayment;
            finalInterestRate = 0;
            finalCustomPayment = customPayment;
        } else {
            monthlyAmount = calculateAmortization(principal, interestRate, totalMonths);
        }

        const monthlyPayments = [];
        for (let i = 1; i <= totalMonths; i++) {
            monthlyPayments.push({
                month: i,
                amount: monthlyAmount,
                status: 'unpaid'
            });
        }

        const newDebt = {
            id: id,
            itemName: itemNameInput.value,
            creditor: creditorNameInput.value,
            principal: principal,
            interestRate: finalInterestRate,
            totalMonths: totalMonths,
            paymentDay: parseInt(paymentDayInput.value) || 1, 
            customPayment: finalCustomPayment,
            monthlyPayments: monthlyPayments
        };

        const existingIndex = debts.findIndex(d => d.id === id);
        
        if (existingIndex > -1) {
            const oldDebt = debts[existingIndex];
            if (
                oldDebt.principal !== newDebt.principal ||
                oldDebt.interestRate !== newDebt.interestRate ||
                oldDebt.totalMonths !== newDebt.totalMonths ||
                oldDebt.customPayment !== newDebt.customPayment
            ) {
                // Core data changed, replace payments
                debts[existingIndex] = newDebt;
            } else {
                // Only metadata changed, keep existing payment statuses
                newDebt.monthlyPayments = oldDebt.monthlyPayments; 
                debts[existingIndex] = newDebt;
            }
        } else {
            // Adding new debt
            debts.push(newDebt);
        }

        saveDebts();
        Swal.fire({
            title: 'สำเร็จ!',
            text: 'บันทึกข้อมูลหนี้เรียบร้อยแล้ว',
            icon: 'success',
            timer: 1500,
            showConfirmButton: false
        });

        renderDashboard();
        showView('dashboard');
    };

    /**
     * Toggles the "Custom Payment" input group
     */
    const toggleCustomPayment = () => {
        if (enableCustomPaymentCheck.checked) {
            customPaymentGroup.style.display = 'block';
            interestRateInput.disabled = true;
            interestRateInput.value = 0; 
        } else {
            customPaymentGroup.style.display = 'none';
            interestRateInput.disabled = false;
        }
    };

    /**
     * Adds listeners to buttons on the dashboard cards
     */
    const addDashboardListeners = () => {
        // This queries the whole document, so it works
        // for buttons in both activeDebtList and completedDebtList.
        document.querySelectorAll('.view-details-btn').forEach(btn => {
            btn.addEventListener('click', () => renderDetailView(btn.dataset.id));
        });

        document.querySelectorAll('.edit-debt-btn').forEach(btn => {
            btn.addEventListener('click', () => populateFormForEdit(btn.dataset.id));
        });

        document.querySelectorAll('.delete-debt-btn').forEach(btn => {
            btn.addEventListener('click', () => deleteDebt(btn.dataset.id));
        });
    };

    /**
     * Deletes a debt item after confirmation
     */
    const deleteDebt = (debtId) => {
        Swal.fire({
            title: 'คุณแน่ใจหรือไม่?',
            text: "คุณจะไม่สามารถกู้คืนข้อมูลนี้ได้!",
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#d33',
            cancelButtonColor: '#3085d6',
            confirmButtonText: 'ใช่, ลบเลย!',
            cancelButtonText: 'ยกเลิก'
        }).then((result) => {
            if (result.isConfirmed) {
                debts = debts.filter(d => d.id !== debtId);
                saveDebts();
                renderDashboard(); // Re-render
                Swal.fire(
                    'ลบแล้ว!',
                    'รายการหนี้ของคุณถูกลบแล้ว',
                    'success'
                );
            }
        });
    };

    /**
     * Toggles the payment status (paid/unpaid) for a month
     */
    const togglePaymentStatus = (debtId, month) => {
        const debt = debts.find(d => d.id === debtId);
        if (!debt) return;

        const payment = debt.monthlyPayments.find(p => p.month === month);
        if (!payment) return;

        payment.status = (payment.status === 'unpaid') ? 'paid' : 'unpaid';
        saveDebts();
        renderDetailView(debtId); // Re-render the detail view
    };


    /**
     * Exports all debt data to an Excel file
     */
    const handleExport = () => {
        if (debts.length === 0) {
            Swal.fire('ไม่มีข้อมูล', 'ไม่มีข้อมูลหนี้สำหรับ Export', 'info');
            return;
        }

        const summaryData = [
            ["ID", "ชื่อรายการ", "เจ้าหนี้", "เงินต้น", "ดอกเบี้ย (%)", "จำนวนเดือน", "ยอดคงเหลือ", "จ่ายวันที่", "ยอดจ่ายกำหนดเอง"]
        ];
        
        debts.forEach(debt => {
            summaryData.push([
                debt.id,
                debt.itemName,
                debt.creditor,
                debt.principal,
                debt.interestRate,
                debt.totalMonths,
                getRemainingBalance(debt), 
                debt.paymentDay,
                debt.customPayment || null
            ]);
        });

        const paymentsData = [
            ["Debt ID", "ชื่อรายการ", "เดือนที่", "ยอดชำระ", "สถานะ"]
        ];

        debts.forEach(debt => {
            debt.monthlyPayments.forEach(p => {
                paymentsData.push([
                    debt.id,
                    debt.itemName,
                    p.month,
                    p.amount,
                    p.status
                ]);
            });
        });

        const wb = XLSX.utils.book_new();
        const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
        const wsPayments = XLSX.utils.aoa_to_sheet(paymentsData);

        XLSX.utils.book_append_sheet(wb, wsSummary, "สรุปรายการหนี้");
        XLSX.utils.book_append_sheet(wb, wsPayments, "รายละเอียดการชำระ");

        XLSX.writeFile(wb, "debt_records.xlsx");

        Swal.fire('Export สำเร็จ!', 'ข้อมูลถูกบันทึกเป็นไฟล์ debt_records.xlsx', 'success');
    };

    /**
     * Triggers the hidden file input
     */
    const triggerImport = () => {
        importExcelInput.click();
    };

    /**
     * Handles the file import process
     */
    const handleFileImport = (e) => {
        const file = e.target.files[0];
        if (!file) return;

        Swal.fire({
            title: 'นำเข้าข้อมูล',
            text: 'คุณต้องการ "แทนที่" ข้อมูลทั้งหมด หรือ "รวม" (เพิ่ม/อัปเดต) ข้อมูล?',
            icon: 'question',
            showCancelButton: true,
            confirmButtonText: '<i class="bi bi-plus-lg"></i> รวมข้อมูล',
            cancelButtonText: '<i class="bi bi-trash3"></i> แทนที่ทั้งหมด',
            reverseButtons: true
        }).then((result) => {
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    const summarySheet = workbook.Sheets["สรุปรายการหนี้"];
                    const paymentsSheet = workbook.Sheets["รายละเอียดการชำระ"];

                    if (!summarySheet || !paymentsSheet) {
                        throw new Error("ไฟล์ไม่ถูกต้อง ขาด Sheet 'สรุปรายการหนี้' หรือ 'รายละเอียดการชำระ'");
                    }

                    const summaryData = XLSX.utils.sheet_to_json(summarySheet);
                    const paymentsData = XLSX.utils.sheet_to_json(paymentsSheet);

                    const paymentStatusMap = new Map();
                    paymentsData.forEach(p => {
                        const debtId = p["Debt ID"];
                        if (!paymentStatusMap.has(debtId)) {
                            paymentStatusMap.set(debtId, new Map());
                        }
                        paymentStatusMap.get(debtId).set(p["เดือนที่"], p["สถานะ"]);
                    });

                    const importedDebts = [];
                    summaryData.forEach(item => {
                        const debtId = item["ID"];
                        const principal = parseFloat(item["เงินต้น"]);
                        const interestRate = parseFloat(item["ดอกเบี้ย (%)"]);
                        const totalMonths = parseInt(item["จำนวนเดือน"]);
                        const customPayment = parseFloat(item["ยอดจ่ายกำหนดเอง"]) || null;
                        
                        const monthlyAmount = customPayment ? customPayment : calculateAmortization(principal, interestRate, totalMonths);
                        
                        const debtPaymentsMap = paymentStatusMap.get(debtId) || new Map();
                        const monthlyPayments = [];

                        for (let i = 1; i <= totalMonths; i++) {
                            monthlyPayments.push({
                                month: i,
                                amount: monthlyAmount,
                                status: debtPaymentsMap.get(i) || 'unpaid'
                            });
                        }

                        const newDebt = {
                            id: debtId,
                            itemName: item["ชื่อรายการ"],
                            creditor: item["เจ้าหนี้"],
                            principal: principal,
                            interestRate: interestRate,
                            totalMonths: totalMonths,
                            paymentDay: parseInt(item["จ่ายวันที่"]),
                            customPayment: customPayment,
                            monthlyPayments: monthlyPayments
                        };
                        importedDebts.push(newDebt);
                    });

                    if (result.isConfirmed) {
                        // Merge
                        importedDebts.forEach(newDebt => {
                            const existingIndex = debts.findIndex(d => d.id === newDebt.id);
                            if (existingIndex > -1) {
                                debts[existingIndex] = newDebt; 
                            } else {
                                debts.push(newDebt); 
                            }
                        });
                    } else if (result.dismiss === Swal.DismissReason.cancel) {
                        // Replace
                        debts = importedDebts;
                    }

                    saveDebts();
                    currentPage = 1; 
                    renderDashboard();
                    Swal.fire('นำเข้าสำเร็จ!', 'ข้อมูลของคุณถูกนำเข้าเรียบร้อยแล้ว', 'success');

                } catch (error) {
                    Swal.fire('เกิดข้อผิดพลาด!', error.message, 'error');
                } finally {
                    importExcelInput.value = '';
                }
            };
            reader.readAsArrayBuffer(file);
        });
    };


    // --- Initialization ---
    
    // Setup Navigation Listeners
    navLinks.brand.addEventListener('click', handleNavigation);
    navLinks.dashboard.addEventListener('click', handleNavigation);
    navLinks.addDebt.addEventListener('click', handleNavigation);

    // Setup Form Listeners
    debtForm.addEventListener('submit', handleFormSubmit);
    enableCustomPaymentCheck.addEventListener('change', toggleCustomPayment);
    
    // Setup Other Listeners
    backToDashboardBtn.addEventListener('click', (e) => {
        e.preventDefault();
        renderDashboard(); 
        showView('dashboard');
    });

    exportExcelBtn.addEventListener('click', handleExport);
    importExcelBtn.addEventListener('click', triggerImport);
    importExcelInput.addEventListener('change', handleFileImport);

    // Pagination Listeners
    prevPageBtn.addEventListener('click', (e) => {
        e.preventDefault();
        if (currentPage > 1) {
            currentPage--;
            renderDashboard();
        }
    });

    nextPageBtn.addEventListener('click', (e) => {
        e.preventDefault();
        currentPage++;
        renderDashboard();
    });
    
    // Initialize and update the clock
    updateClock(); // เรียกครั้งแรกทันที
    setInterval(updateClock, 1000); // อัปเดตทุก 1 วินาที


    // Initial Load
    loadDebts();
    renderDashboard();
    showView('dashboard');

});