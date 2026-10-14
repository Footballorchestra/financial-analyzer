// balances-manager.js — Управление страницей остатков
class BalancesManager {
    constructor(storage) {
        this.storage = storage;
        this.editingAccount = null;
        this.editingDeposits = []; // временный массив депозитов для модального окна
        this.init();
    }

    init() {
        console.log('Initializing BalancesManager...');
        this.setupEventListeners();
    }

    setupEventListeners() {
        document.getElementById('balancesTableBody').addEventListener('click', (e) => {
            const cell = e.target.closest('.editable');
            if (cell && cell.dataset.account) {
                this.openEditModal(cell.dataset.account, cell.dataset.field);
            }
        });

        const searchInput = document.getElementById('balancesSearchInput');
        const clearBtn = document.getElementById('balancesClearSearchBtn');
        if (searchInput) searchInput.addEventListener('input', () => this.filterTable(searchInput.value));
        if (clearBtn) clearBtn.addEventListener('click', () => { searchInput.value = ''; this.filterTable(''); });
    }

    filterTable(query) {
        const tbody = document.getElementById('balancesTableBody');
        const rows = tbody.querySelectorAll('tr:not(.empty-row)');
        const term = query.toLowerCase().trim();
        let visibleCount = 0;
        rows.forEach(row => {
            row.style.display = (!term || row.textContent.toLowerCase().includes(term)) ? '' : 'none';
            if (row.style.display !== 'none') visibleCount++;
        });
        document.getElementById('totalAccountsCount').textContent = visibleCount;
    }

    // Режим «только проценты по длинным депозитам» — галочка на странице остатков
    // (включена по умолчанию).
    // Длинный (срочный) депозит — депозит с endDate: вносится через форму и
    // подсвечивается в таблице. Обычный депозит — без endDate: загружается из Excel.
    // Когда галочка включена, тело длинных депозитов не показывается и не прибавляется
    // к остаткам (учитываются только их проценты), а обычные депозиты — целиком (тело + проценты).
    isInterestOnlyMode() {
        const checkbox = document.getElementById('depositsInterestOnly');
        return checkbox ? checkbox.checked : true; // по умолчанию — включено
    }

    updateTable() {
        const accounts = this.storage.getAccounts();
        const accountNumbers = Object.keys(accounts);
        const tbody = document.getElementById('balancesTableBody');
        if (accountNumbers.length === 0) {
            tbody.innerHTML = `<tr class="empty-row"><td colspan="9">Нет счетов. Добавьте счета в разделе «Настройки» → «Счета».</td></tr>`;
            return;
        }

        const interestOnly = this.isInterestOnlyMode();
        let html = '', totalBalance = 0, totalDeposit = 0, totalInterest = 0;
        accountNumbers.forEach(accountNumber => {
            const account = accounts[accountNumber];
            const deposits = this.storage.getDepositForAccount(accountNumber);
            const latestStatement = this.getLatestStatementForAccount(accountNumber);
            const balance = latestStatement?.balance || account?.balance || 0;
            const rawDate = new Date().toISOString().split('T')[0];
            const statementDate = rawDate.split('-').reverse().join('.');

            const result = this.calculateInterestAndRealBalance(accountNumber, balance, deposits, interestOnly);
            const { realBalance, unlockedAmount, unlockedInterest, regularAmount, isFrozen } = result;

            // В режиме «только проценты по длинным депозитам» тело длинных (срочных)
            // депозитов не показываем — остаётся тело обычных депозитов (Excel).
            // unlockedAmount остаётся «сырым» — нужен для подсказки о заморозке.
            const depositBodyShown = interestOnly ? regularAmount : unlockedAmount;
            // Начисленные проценты в обоих режимах — по обычным и «вернувшимся» депозитам;
            // проценты ещё замороженных (не вернувшихся) длинных депозитов не начисляются.
            const interestShown = unlockedInterest;

            const balanceFormatted = this.storage.formatNumber(balance);
            const depositFormatted = depositBodyShown > 0 ? this.storage.formatNumber(depositBodyShown) : '';
            const depositTitle = interestOnly
                ? ' title="Режим «Только проценты»: тело длинных (срочных) депозитов не показывается и не учитывается в остатках"'
                : '';
            const interestFormatted = interestShown > 0 ? this.storage.formatNumber(interestShown) : '';
            const realBalanceFormatted = this.storage.formatNumber(realBalance);

            const calcDate = document.getElementById('calculationDate').value;
            const unlockedDeposits = deposits.filter(d => {
                if (!d.amount || d.amount <= 0) return false;
                if (!d.endDate) return true;
                return calcDate >= d.endDate;
            });
            const rateDisplay = unlockedDeposits.length > 0 ? unlockedDeposits.map(d => d.rate || 0).join('% / ') + '%' : '';

            totalBalance += balance;
            totalDeposit += depositBodyShown;
            totalInterest += interestShown;

            const frozenClass = isFrozen ? ' class="deposit-frozen"' : '';
            let frozenTitle = '';
            if (isFrozen) {
                const totalAll = deposits.reduce((s, d) => s + (d.amount || 0), 0);
                const frozenSum = totalAll - unlockedAmount;
                const frozenDep = deposits.find(d => d.endDate && d.amount > 0);
                if (frozenDep) frozenTitle = ` title="Заморожен срочный депозит на ${this.storage.formatNumber(frozenSum)} ₽ до ${frozenDep.endDate}"`;
            }

            html += `<tr${frozenClass}${frozenTitle}><td>${account.company || ''}</td><td>${account.bank || ''}</td><td>${accountNumber}</td><td class="number-cell">${balanceFormatted}</td><td class="number-cell editable" data-account="${accountNumber}" data-field="amount"${depositTitle}>${depositFormatted}</td><td class="number-cell editable" data-account="${accountNumber}" data-field="rate">${rateDisplay}</td><td class="number-cell">${interestFormatted}</td><td class="number-cell">${realBalanceFormatted}</td><td>${statementDate}</td></tr>`;
        });

        tbody.innerHTML = html;
        this.updateSummary(totalBalance, totalDeposit, totalInterest, accountNumbers.length);
    }

    calculateInterestAndRealBalance(accountNumber, balance, deposits, interestOnly = false) {
        const result = {
            interest: 0,          // начисленные проценты по всем действующим депозитам
            realBalance: balance,
            isFrozen: false,
            unlockedAmount: 0,    // тело «вернувшихся»/бессрочных депозитов (нужно для подсказки о заморозке)
            unlockedInterest: 0,  // проценты по «вернувшихся»/бессрочным депозитам
            regularAmount: 0      // тело обычных депозитов (без endDate — загружаются из Excel)
        };
        if (!Array.isArray(deposits) || deposits.length === 0) return result;

        const calculationDate = document.getElementById('calculationDate').value;
        if (!calculationDate) return result;

        const mibAccounts = ['40702810700990012381', '40702810100990012143'];
        const isMib = mibAccounts.includes(accountNumber);

        let totalInterest = 0, unlockedBody = 0, unlockedInt = 0, regularBody = 0;
        let hasFrozen = false, anyDepositActive = false;

        deposits.forEach(dep => {
            const amount = dep.amount || 0, rate = dep.rate || 0, startDate = dep.startDate || '', endDate = dep.endDate || null;
            if (amount <= 0) return;

            // Обычный депозит (без endDate) приходит из Excel-файла — его тело учитываем
            // всегда, даже если дату начала не удалось определить и проценты не считаются.
            // Длинный/срочный (с endDate) вносится через форму и подсвечивается в таблице.
            if (!endDate) {
                regularBody += amount;
            }

            // Определяем заморозку до проверки days — депозит заморожен,
            // даже если расчётная дата ещё не достигла даты начала.
            // endDate — первый день, когда депозит уже НЕ действует
            if (endDate && calculationDate < endDate) {
                hasFrozen = true;
            }

            if (!startDate) return;

            const days = this.storage.getDaysBetween(startDate, calculationDate);
            if (days <= 0) return;

            anyDepositActive = true;
            const interest = this.storage.calculateInterest(amount, rate, days);
            totalInterest += interest;

            // endDate — первый день, когда депозит уже НЕ действует
            if (!endDate || calculationDate >= endDate) {
                unlockedBody += amount;
                unlockedInt += interest;
            }
        });

        result.isFrozen = hasFrozen;
        // Тело обычных депозитов (Excel) учитывается и когда проценты посчитать нельзя
        result.regularAmount = regularBody;
        if (!anyDepositActive) return result;

        result.interest = totalInterest;
        result.unlockedAmount = unlockedBody;
        result.unlockedInterest = unlockedInt;
        // Реальный остаток:
        //  • счета МИБ (40702810700990012381, 40702810100990012143): особенность банка —
        //    при размещении денег на депозит фактического списания не происходит, деньги
        //    остаются на счёте, и к моменту «возврата» депозита они там уже есть. Поэтому
        //    тело депозита (и обычного из Excel, и «вернувшегося» срочного) НЕ прибавляем —
        //    иначе остаток задваивается. Учитываем только начисленные проценты.
        //    Правило действует всегда и имеет приоритет над режимом «только проценты»;
        //  • галочка «только проценты по длинным депозитам» включена (по умолчанию):
        //    длинные (срочные) депозиты не участвуют в остатке ни телом, ни процентами,
        //    пока депозит не вернулся (endDate >= расчётной даты). Обычные депозиты (Excel)
        //    учитываются целиком — тело + проценты;
        //  • галочка выключена: прежнее поведение — тело «вернувшихся»/бессрочных депозитов
        //    и их проценты.
        if (isMib) {
            result.realBalance = balance + unlockedInt;
        } else if (interestOnly) {
            result.realBalance = balance + regularBody + unlockedInt;
        } else {
            result.realBalance = balance + unlockedBody + unlockedInt;
        }

        return result;
    }

    getLatestStatementForAccount(accountNumber) {
        const statements = this.storage.getStatements().filter(s => s.account === accountNumber);
        if (statements.length === 0) return null;
        statements.sort((a, b) => this.storage.parseDate(b.date) - this.storage.parseDate(a.date));
        return statements[0];
    }

    updateSummary(totalBalance, totalDeposit, totalInterest, accountCount) {
        document.getElementById('totalAccountsCount').textContent = accountCount;
        document.getElementById('totalBalanceAmount').textContent = this.storage.formatCurrency(totalBalance);
        document.getElementById('totalInterestsAmount').textContent = this.storage.formatCurrency(totalInterest);
    }

    openEditModal(accountNumber, field) {
        const accounts = this.storage.getAccounts();
        const account = accounts[accountNumber];
        const deposits = this.storage.getDepositForAccount(accountNumber);
        this.editingAccount = accountNumber;
        this.editingDeposits = deposits.map(d => ({ ...d }));
        document.getElementById('editCompany').value = account?.company || '';
        document.getElementById('editAccount').value = accountNumber;
        this.renderDepositList();
        document.getElementById('depositModal').classList.add('active');
    }

    renderDepositList() {
        const container = document.getElementById('depositListContainer');
        if (!container) return;
        if (this.editingDeposits.length === 0) {
            container.innerHTML = '<p style="color: var(--text-tertiary); font-size: 13px; text-align: center; padding: 12px;">Нет депозитов. Нажмите «Добавить депозит».</p>';
            return;
        }
        let html = '';
        this.editingDeposits.forEach((dep, index) => {
            const isFrozen = dep.endDate ? ' (срочный)' : ' (обычный)';
            html += `<div class="deposit-item" style="display: flex; gap: 8px; align-items: flex-end; margin-bottom: 12px; padding: 12px; background: var(--bg-tertiary); border-radius: var(--radius-sm);"><div style="flex: 1; display: grid; grid-template-columns: 1fr 1fr; gap: 8px;"><div><label style="font-size: 10px; color: var(--text-secondary); display: block;">Сумма</label><input type="number" class="dep-amount" data-index="${index}" value="${dep.amount || ''}" step="0.01" style="width: 100%; padding: 6px 8px; border: 1px solid var(--border-light); border-radius: 4px; font-size: 12px;"></div><div><label style="font-size: 10px; color: var(--text-secondary); display: block;">Ставка, %</label><input type="number" class="dep-rate" data-index="${index}" value="${dep.rate || ''}" step="0.01" style="width: 100%; padding: 6px 8px; border: 1px solid var(--border-light); border-radius: 4px; font-size: 12px;"></div><div><label style="font-size: 10px; color: var(--text-secondary); display: block;">Дата начала</label><input type="date" class="dep-start" data-index="${index}" value="${dep.startDate || ''}" style="width: 100%; padding: 6px 8px; border: 1px solid var(--border-light); border-radius: 4px; font-size: 12px;"></div><div><label style="font-size: 10px; color: var(--text-secondary); display: block;">Дата окончания${isFrozen}</label><input type="date" class="dep-end" data-index="${index}" value="${dep.endDate || ''}" style="width: 100%; padding: 6px 8px; border: 1px solid var(--border-light); border-radius: 4px; font-size: 12px;"></div></div><button class="btn btn-sm btn-danger remove-deposit-btn" data-index="${index}" title="Удалить депозит" style="flex-shrink: 0; height: 30px;">✕</button></div>`;
        });
        container.innerHTML = html;
        container.querySelectorAll('.remove-deposit-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                this.editingDeposits.splice(parseInt(btn.dataset.index), 1);
                this.renderDepositList();
            });
        });
    }

    async saveDeposit() {
        if (!this.editingAccount) return;
        const container = document.getElementById('depositListContainer');
        const updatedDeposits = [];
        this.editingDeposits.forEach((dep, index) => {
            const amountEl = container.querySelector(`.dep-amount[data-index="${index}"]`);
            if (!amountEl) return;
            const amount = parseFloat(amountEl.value) || 0;
            const rate = parseFloat(container.querySelector(`.dep-rate[data-index="${index}"]`)?.value) || 0;
            const startDate = container.querySelector(`.dep-start[data-index="${index}"]`)?.value || '';
            const endDate = container.querySelector(`.dep-end[data-index="${index}"]`)?.value || null;
            if (amount > 0 && rate > 0) updatedDeposits.push({ amount, rate, startDate, endDate: endDate || null });
        });
        this.storage.setDepositForAccount(this.editingAccount, updatedDeposits);
        // Синхронизируем с БД, явно передавая счёт: если в модальном окне удалили все
        // депозиты счёта, он всё равно должен попасть в payload (с пустым списком),
        // иначе запись в term_deposits останется и депозит вернётся после перезагрузки.
        await this.storage.syncDepositsToServer([this.editingAccount]);
        document.getElementById('depositModal').classList.remove('active');
        this.editingAccount = null;
        this.editingDeposits = [];
        this.updateTable();
        window.app.showNotification('Данные по депозитам сохранены', 'success');
    }

    addDeposit() {
        this.editingDeposits.push({ amount: 0, rate: 0, startDate: '', endDate: null });
        this.renderDepositList();
    }

    calculateInterests() {
        if (!document.getElementById('calculationDate').value) { alert('Пожалуйста, укажите дату расчета'); return; }
        this.updateTable();
        window.app.showNotification('Проценты рассчитаны', 'success');
    }

    // Сегодняшняя дата в формате YYYY-MM-DD (локальная)
    todayIso() {
        const d = new Date();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${d.getFullYear()}-${m}-${day}`;
    }

    // Следующий рабочий день после указанной даты (пропускаем субботу и воскресенье)
    getNextWorkingDayIso(dateStr) {
        const base = dateStr || this.todayIso();
        const [y, m, d] = base.split('-').map(Number);
        const date = new Date(y, m - 1, d);
        do {
            date.setDate(date.getDate() + 1);
        } while (date.getDay() === 0 || date.getDay() === 6); // 0 — воскресенье, 6 — суббота
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');
        return `${date.getFullYear()}-${mm}-${dd}`;
    }

    // Форматирование ISO-даты в ДД.ММ.ГГГГ для уведомлений
    formatDateRu(dateStr) {
        if (!dateStr) return '';
        const [y, m, d] = dateStr.split('-');
        return `${d}.${m}.${y}`;
    }

    // Продление всех срочных (длинных) депозитов: от текущего дня до следующего рабочего дня
    async extendTermDepositsToNextWorkingDay() {
        try {
            const calcInput = document.getElementById('calculationDate');
            const startDate = (calcInput && calcInput.value) ? calcInput.value : this.todayIso();
            const endDate = this.getNextWorkingDayIso(startDate);

            const accounts = Object.keys(this.storage.getDepositData());
            let accountsChanged = 0;
            let depositsChanged = 0;

            accounts.forEach(account => {
                const deposits = this.storage.getDepositForAccount(account);
                let changed = false;
                const updated = deposits.map(dep => {
                    if (!dep.endDate) return dep; // обычные (без даты окончания) не трогаем
                    changed = true;
                    depositsChanged++;
                    return { ...dep, startDate, endDate };
                });
                if (changed) {
                    this.storage.setDepositForAccount(account, updated);
                    accountsChanged++;
                }
            });

            if (depositsChanged === 0) {
                window.app.showNotification('Срочные (длинные) депозиты не найдены', 'info');
                return;
            }

            await this.storage.syncDepositsToServer();
            this.updateTable();
            window.app.showNotification(
                `Продлено срочных депозитов: ${depositsChanged} (на ${accountsChanged} счетах) — период ${this.formatDateRu(startDate)}–${this.formatDateRu(endDate)}`,
                'success'
            );
        } catch (error) {
            console.error('Ошибка продления депозитов на следующий рабочий день:', error);
            window.app.showNotification('Ошибка продления депозитов', 'error');
        }
    }

    // Мержит депозит из Excel с существующими депозитами.
    // Срочные (с endDate) сохраняются, обычный (без endDate) ЗАМЕНЯЕТСЯ данными из файла:
    // файл — снимок текущих депозитов, поэтому повторная загрузка того же файла
    // (в т.ч. после перезагрузки страницы, когда обычные депозиты восстанавливаются
    // из localStorage) не должна суммировать суммы и удваивать остаток.
    mergeExcelDeposit(account, excelDeposit) {
        const existing = this.storage.getDepositForAccount(account);
        const termOnly = existing.filter(d => d.endDate); // сохраняем срочные
        const regularExisting = existing.find(d => !d.endDate); // уже известный обычный депозит

        termOnly.push({
            amount: excelDeposit.amount || 0,
            rate: excelDeposit.rate || 0,
            // если в файле не оказалось даты начала — оставляем ранее известную
            startDate: excelDeposit.startDate || (regularExisting ? regularExisting.startDate : '') || '',
            endDate: null
        });

        this.storage.setDepositForAccount(account, termOnly);
    }

    async loadDepositData(file) {
        if (!file) return;
        try {
            if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
                await this.loadExcelDepositData(file);
            } else {
                const text = await this.readFile(file);
                const depositData = this.parseDepositData(text);
                Object.entries(depositData).forEach(([account, data]) => this.mergeExcelDeposit(account, data));
                window.app.showNotification(`Загружены данные по ${Object.keys(depositData).length} депозитам`, 'success');
            }
            this.updateTable();
        } catch (error) {
            console.error('Error loading deposit data:', error);
            window.app.showNotification('Ошибка загрузки данных по депозитам', 'error');
        }
    }

    async loadExcelDepositData(file) {
        try {
            const workbook = await this.readExcelFile(file);
            const depositData = this.parseDepositDataFromExcel(workbook);
            Object.entries(depositData).forEach(([account, data]) => this.mergeExcelDeposit(account, data));
            window.app.showNotification(`Загружены данные по ${Object.keys(depositData).length} депозитам из Excel`, 'success');
        } catch (error) {
            console.error('Ошибка обработки Excel файла:', error);
            throw new Error('Ошибка обработки Excel файла: ' + error.message);
        }
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try { resolve(XLSX.read(new Uint8Array(e.target.result), { type: 'array' })); }
                catch (error) { reject(new Error('Ошибка чтения Excel файла: ' + error.message)); }
            };
            reader.onerror = () => reject(new Error('Ошибка чтения файла'));
            reader.readAsArrayBuffer(file);
        });
    }

    parseDepositDataFromExcel(workbook) {
        const depositData = {};
        let sheetName = workbook.SheetNames.find(n => n.toLowerCase().includes('свод')) || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        let accountCol = 1, amountCol = 2, rateCol = 3, startDateCol = 5;
        if (jsonData.length > 0) {
            const h = jsonData[0];
            for (let i = 0; i < h.length; i++) {
                const v = String(h[i] || '').toLowerCase();
                if (v.includes('счет') || v.includes('номер')) accountCol = i;
                if (v.includes('сумма') && v.includes('депозит')) amountCol = i;
                if (v.includes('ставка')) rateCol = i;
                if (v.includes('дата') && v.includes('начал')) startDateCol = i;
            }
        }
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length <= Math.max(accountCol, amountCol, rateCol, startDateCol)) continue;
            const accountMatch = String(row[accountCol] || '').trim().match(/\d{20}/);
            if (!accountMatch) continue;
            const account = accountMatch[0];
            let amount = 0;
            const rawA = row[amountCol];
            if (rawA !== undefined && rawA !== null && rawA !== '') amount = typeof rawA === 'number' ? rawA : parseFloat(String(rawA).replace(/[^\d.-]/g, '').replace(',', '.')) || 0;
            let rate = 0;
            const rawR = row[rateCol];
            if (rawR !== undefined && rawR !== null && rawR !== '') rate = typeof rawR === 'number' ? rawR : parseFloat(String(rawR).replace(/[^\d.-]/g, '').replace(',', '.')) || 0;
            let startDate = '';
            const rawD = row[startDateCol];
            if (rawD !== undefined && rawD !== null && rawD !== '') {
                if (typeof rawD === 'number') startDate = new Date((rawD - 25569) * 86400 * 1000).toISOString().split('T')[0];
                else if (typeof rawD === 'string') {
                    const s = rawD.trim();
                    if (s.match(/^\d{4}-\d{2}-\d{2}$/)) startDate = s;
                    else if (s.match(/^\d{2}\.\d{2}\.\d{4}$/)) { const p = s.split('.'); startDate = `${p[2]}-${p[1]}-${p[0]}`; }
                    else { const d = new Date(s); if (!isNaN(d.getTime())) startDate = d.toISOString().split('T')[0]; }
                }
            }
            if (account && amount > 0 && rate > 0) {
                if (depositData[account]) {
                    // Суммируем несколько строк Excel для одного счёта (длинный + короткий депозит)
                    const existing = depositData[account];
                    const totalAmount = existing.amount + amount;
                    const weightedRate = (existing.amount * existing.rate + amount * rate) / totalAmount;
                    depositData[account] = {
                        amount: totalAmount,
                        rate: weightedRate,
                        startDate: existing.startDate || startDate || new Date().toISOString().split('T')[0],
                        endDate: null
                    };
                } else {
                    depositData[account] = { amount, rate, startDate: startDate || new Date().toISOString().split('T')[0], endDate: null };
                }
            }
        }
        return depositData;
    }

    readFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = () => reject(new Error('Ошибка чтения файла'));
            reader.readAsText(file, 'UTF-8');
        });
    }

    parseDepositData(text) {
        const depositData = {};
        text.split('\n').forEach(line => {
            const t = line.trim();
            if (!t) return;
            const parts = t.split(/[\t,;]/).map(p => p.trim());
            if (parts.length >= 3) {
                const account = parts[0].replace(/\s/g, '');
                const amount = parseFloat(parts[1].replace(',', '.')) || 0;
                const rate = parseFloat(parts[2].replace(',', '.')) || 0;
                if (account && amount > 0 && rate > 0) {
                    if (depositData[account]) {
                        // Суммируем несколько строк для одного счёта
                        const existing = depositData[account];
                        const totalAmount = existing.amount + amount;
                        const weightedRate = (existing.amount * existing.rate + amount * rate) / totalAmount;
                        depositData[account] = { amount: totalAmount, rate: weightedRate, startDate: existing.startDate || parts[3] || '', endDate: null };
                    } else {
                        depositData[account] = { amount, rate, startDate: parts[3] || '', endDate: null };
                    }
                }
            }
        });
        return depositData;
    }

    exportToExcel() {
        const accounts = this.storage.getAccounts();
        const accountNumbers = Object.keys(accounts);
        if (accountNumbers.length === 0) { alert('Нет данных для экспорта'); return; }

        let sorted = [...accountNumbers];
        try {
            const parser = new BankStatementParser();
            const order = Object.keys(parser.loadAccountMapping());
            sorted.sort((a, b) => { const ia = order.indexOf(a), ib = order.indexOf(b); if (ia !== -1 && ib !== -1) return ia - ib; if (ia !== -1) return -1; if (ib !== -1) return 1; return a.localeCompare(b); });
        } catch (e) {}

        const interestOnly = this.isInterestOnlyMode();
        const data = [['Компания', 'Банк', 'Счёт', 'Остаток по выписке', 'Вернувшийся депозит', 'Начисленные проценты', 'Реальный остаток', 'Дата']];
        let totalBalance = 0, totalDeposit = 0, totalInterest = 0, totalRealBalance = 0;

        sorted.forEach(accountNumber => {
            const account = accounts[accountNumber];
            const deposits = this.storage.getDepositForAccount(accountNumber);
            const latestStatement = this.getLatestStatementForAccount(accountNumber);
            const balance = latestStatement?.balance || account?.balance || 0;
            const rawDate = new Date().toISOString().split('T')[0];
            const statementDate = rawDate.split('-').reverse().join('.');
            const { realBalance, unlockedAmount, unlockedInterest, regularAmount } = this.calculateInterestAndRealBalance(accountNumber, balance, deposits, interestOnly);
            // В режиме «только проценты по длинным депозитам» в выгрузку не попадает тело длинных депозитов
            const depositBodyShown = interestOnly ? regularAmount : unlockedAmount;
            const interestShown = unlockedInterest;
            data.push([account.company || '', account.bank || '', accountNumber, balance, depositBodyShown, interestShown, realBalance, statementDate]);
            totalBalance += balance; totalDeposit += depositBodyShown; totalInterest += interestShown; totalRealBalance += realBalance;
        });

        data.push(['ИТОГО', '', '', totalBalance, totalDeposit, totalInterest, totalRealBalance, '']);
        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Остатки');
        ws['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 20 }, { wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 20 }, { wch: 12 }];
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        [3, 4, 5, 6].forEach(col => { for (let r = 1; r <= range.e.r; r++) { const c = XLSX.utils.encode_cell({ r, c: col }); if (ws[c]) { ws[c].t = 'n'; ws[c].z = '#,##0.00'; } } });
        for (let col = 0; col <= 6; col++) { const c = XLSX.utils.encode_cell({ r: range.e.r, c: col }); if (ws[c]) ws[c].s = { font: { bold: true } }; }
        XLSX.writeFile(wb, `Остатки_${new Date().toISOString().slice(0, 10)}.xlsx`, { cellStyles: true });
        window.app.showNotification(`Экспортировано ${sorted.length} счетов`, 'success');
    }
}