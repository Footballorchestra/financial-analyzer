// receipts-manager.js - Управление страницей поступлений
// Все правила загружаются из storage (БД), без хардкода
class ReceiptsManager {
    constructor(storage) {
        this.storage = storage;
        this.searchTerm = '';
        this.init();
    }

    init() {
        console.log('Initializing ReceiptsManager...');
    }

    updateTable() {
        const transactions = this.getFilteredTransactions();
        this.renderTable(transactions);
        this.updateSummary(transactions);
    }

    getFilteredTransactions() {
        let transactions = this.storage.getIncomingTransactions();
        console.log('Всего входящих транзакций:', transactions.length);

        // Применяем правила исключения из БД (настраиваются в разделе "Настройки")
        for (const rule of this.storage.getExclusionRules()) {
            transactions = transactions.filter(t => {
                // API возвращает поле "type" (purpose | counterparty)
                let targetValue = '';
                if (rule.type === 'purpose') {
                    targetValue = (t.purpose || '').toLowerCase();
                } else if (rule.type === 'counterparty') {
                    targetValue = this.getDisplayCounterCompany(t).toLowerCase();
                }
                const pattern = rule.pattern.toLowerCase();
                if (rule.is_regex) {
                    try { return !new RegExp(pattern, 'i').test(targetValue); } catch(e) { return true; }
                }
                return !targetValue.includes(pattern);
            });
        }

        console.log('После исключений осталось:', transactions.length);

        // Применяем поиск
        if (this.searchTerm) {
            const term = this.searchTerm.toLowerCase();
            transactions = transactions.filter(t => {
                const payer = (t.payer || '').toLowerCase();
                const purpose = (t.purpose || '').toLowerCase();
                return payer.includes(term) || purpose.includes(term);
            });
        }

        return transactions;
    }

    renderTable(transactions) {
        const tbody = document.getElementById('receiptsTableBody');

        if (!transactions || transactions.length === 0) {
            tbody.innerHTML = `
                <tr class="empty-row">
                    <td colspan="8">
                        ${this.storage.getTransactions().length === 0
                            ? 'Загрузите выписки на странице "Загрузка выписок"'
                            : 'Нет данных по текущим фильтрам'}
                    </td>
                </tr>
            `;
            return;
        }

        let html = '';
        transactions.forEach(transaction => {
            let payerINN = transaction.payerINN || '';
            if (!payerINN && this.storage.getINNData()[transaction.payer]) {
                payerINN = this.storage.getINNData()[transaction.payer];
            }

            const amountFormatted = this.storage.formatNumber(transaction.amount);

            html += `
                <tr>
                    <td>${this.getDisplayCounterCompany(transaction)}</td>
                    <td>${payerINN}</td>
                    <td>${transaction.ourCompany || ''}</td>
                    <td>${transaction.ourAccount || transaction.recipientAccount || ''}</td>
                    <td class="number-cell">${amountFormatted}</td>
                    <td>${transaction.ourBank || ''}</td>
                    <td>${transaction.purpose || ''}</td>
                    <td>${transaction.date || ''}</td>
                </tr>
            `;
        });

        tbody.innerHTML = html;
    }

    updateSummary(transactions) {
        const totalAmount = transactions.reduce((sum, t) => sum + (t.amount || 0), 0);
        document.getElementById('totalReceiptsCount').textContent = transactions.length;
        document.getElementById('totalReceiptsAmount').textContent =
            this.storage.formatCurrency(totalAmount);
    }

    getDisplayCounterCompany(transaction) {
        const purpose = (transaction.purpose || '').toLowerCase();
        const payer = (transaction.counterCompany || transaction.payer || '').toLowerCase();

        // Применяем правила категоризации из БД
        for (const rule of this.storage.getCategorizationRules()) {
            const fieldValue = rule.field === 'purpose' ? purpose :
                               rule.field === 'counterparty' ? payer :
                               (transaction.payerINN || '').toLowerCase();
            const pattern = rule.pattern.toLowerCase();
            if (fieldValue.includes(pattern)) {
                return rule.display_name;
            }
        }

        // Применяем синонимы компаний из БД (парсер уже нормализовал, но на всякий случай)
        for (const alias of this.storage.getCompanyAliases()) {
            const pattern = alias.pattern || '';
            const canonical = alias.canonical || '';
            if (!pattern || !canonical) continue;
            if (alias.match_type === 'exact') {
                if (payer === pattern.toLowerCase() || (transaction.payer || '').toLowerCase() === pattern.toLowerCase()) return canonical;
            } else if (alias.match_type === 'regex') {
                try { if (new RegExp(pattern, 'i').test(payer)) return canonical; } catch(e) {}
            } else {
                if (payer.includes(pattern.toLowerCase())) return canonical;
            }
        }

        // Универсальное правило: госномер = продажа ТС
        const plateRegex = /[авекмнорстух]\d{3}[авекмнорстух]{2}\d{2,3}/i;
        if (plateRegex.test(purpose)) return 'Продажа ТС';

        return transaction.payer || transaction.counterCompany || '';
    }

    async loadINNData(file) {
        if (!file) return;
        try {
            const workbook = await this.readExcelFile(file);
            const innData = this.parseINNDataFromExcel(workbook);
            this.storage.setINNData(innData);
            this.updateTransactionsWithINNData(innData);
            this.updateTable();
            window.app.showNotification(`Загружено ${Object.keys(innData).length} ИНН`, 'success');
        } catch (error) {
            console.error('Error loading INN data:', error);
            window.app.showNotification('Ошибка загрузки ИНН', 'error');
        }
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    resolve(XLSX.read(data, { type: 'array' }));
                } catch (error) {
                    reject(new Error('Ошибка чтения Excel: ' + error.message));
                }
            };
            reader.onerror = () => reject(new Error('Ошибка чтения файла'));
            reader.readAsArrayBuffer(file);
        });
    }

    parseINNDataFromExcel(workbook) {
        const innData = {};
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        let innCol = -1, nameCol = -1;
        if (rows.length > 0) {
            rows[0].forEach((cell, i) => {
                const c = String(cell).toLowerCase();
                if (c.includes('инн')) innCol = i;
                if (c.includes('наименование') || c.includes('название')) nameCol = i;
            });
        }
        if (innCol === -1) innCol = 0;
        if (nameCol === -1) nameCol = 1;

        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length <= Math.max(innCol, nameCol)) continue;
            let inn = String(row[innCol] || '').trim().replace(/\D/g, '');
            let name = String(row[nameCol] || '').trim();
            if (inn.length >= 10 && inn.length <= 12 && name) {
                innData[inn] = name;
            }
        }
        return innData;
    }

    updateTransactionsWithINNData(innData) {
        const allTransactions = this.storage.getTransactions();
        let updated = false;
        const updatedTransactions = allTransactions.map(transaction => {
            if (transaction.direction === 'incoming' && transaction.payerINN) {
                const correctName = innData[transaction.payerINN];
                if (correctName) {
                    transaction.payer = correctName;
                    transaction.counterCompany = correctName;
                    updated = true;
                }
            }
            return transaction;
        });
        if (updated) {
            this.storage.setTransactions(updatedTransactions);
        }
    }

    // Дата в формате YYYY-MM-DD по локальному времени (без UTC-сдвига)
    toIsoDate(date) {
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${date.getFullYear()}-${m}-${d}`;
    }

    // Предыдущий рабочий день (пропускаем субботу и воскресенье).
    // Поступления в файле относятся к предыдущему рабочему дню, а не к текущему.
    getPreviousWorkingDayIso() {
        const date = new Date();
        do {
            date.setDate(date.getDate() - 1);
        } while (date.getDay() === 0 || date.getDay() === 6); // 0 — воскресенье, 6 — суббота
        return this.toIsoDate(date);
    }

    exportToExcel() {
        const transactions = this.getFilteredTransactions();
        if (transactions.length === 0) {
            alert('Нет данных для экспорта');
            return;
        }

        const data = [
            ['Заказчик', 'ИНН заказчика', 'Юридическое лицо', 'Счет получателя',
             'Сумма', 'Банк получателя', 'Назначение платежа', 'Дата', 'Файл источника']
        ];

        let totalAmount = 0;
        transactions.forEach(t => {
            const bank = t.ourBank || '';
            data.push([
                this.getDisplayCounterCompany(t),
                t.payerINN || '',
                t.ourCompany || '',
                t.ourAccount || t.recipientAccount || '',
                t.amount,
                bank,
                t.purpose || '',
                t.date || '',
                t.sourceFile || ''
            ]);
            totalAmount += t.amount;
        });

        data.push(['ИТОГО', '', '', '', totalAmount, '', '', '', '']);

        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Поступления');

        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let row = 1; row <= range.e.r; row++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: 4 });
            if (!ws[cellAddress]) continue;
            ws[cellAddress].t = 'n';
            ws[cellAddress].z = '#,##0.00';
        }

        const lastRow = range.e.r;
        for (let col = 0; col <= 8; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: lastRow, c: col });
            if (!ws[cellAddress]) continue;
            ws[cellAddress].s = { font: { bold: true } };
        }

        const maxWidth = data.reduce((max, row) => Math.max(max, row.length), 0);
        const colWidths = [];
        for (let i = 0; i < maxWidth; i++) {
            let maxLength = 0;
            data.forEach(row => {
                const cellValue = row[i] || '';
                const length = String(cellValue).length;
                if (length > maxLength) maxLength = length;
            });
            colWidths.push({ wch: Math.min(maxLength + 2, 50) });
        }
        ws['!cols'] = colWidths;

        // Имя файла: по умолчанию — предыдущий рабочий день (поступления относятся к нему),
        // при включённом чекбоксе «Имя файла текущей датой» — текущая дата
        const useCurrentDate = document.getElementById('receiptsUseCurrentDate');
        const fileDate = (useCurrentDate && useCurrentDate.checked)
            ? this.toIsoDate(new Date())
            : this.getPreviousWorkingDayIso();

        XLSX.writeFile(wb, `Поступления_${fileDate}.xlsx`, { cellStyles: true });
    }

    clearSearch() {
        this.searchTerm = '';
        document.getElementById('searchReceipts').value = '';
        this.updateTable();
    }

    searchTransactions(term) {
        this.searchTerm = term;
        this.updateTable();
    }
}