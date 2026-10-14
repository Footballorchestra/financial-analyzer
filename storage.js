// storage.js - Управление хранением данных (упрощенная версия для теста)
class StorageManager {
    constructor() {
        this.files = [];
        this.statements = [];
        this.transactions = [];
        this.accounts = {};
        this.innData = {};
        this.depositData = {}; // Будет загружено асинхронно в initDeposits()
        this.exclusionRules = [];
        this.categorizationRules = [];
        this.companyAliases = [];
    }

    // Асинхронная инициализация депозитов (вызывается из App.init())
    async initDeposits() {
        this.depositData = await this.loadDepositsFromServer();
        // Срочные (длинные) депозиты приходят из БД, обычные (без endDate, из Excel) —
        // их в БД нет (end_date NOT NULL), поэтому возвращаем их из localStorage.
        this.restoreRegularDeposits();
        await this.cleanExpiredDeposits();
    }

    // Управление файлами
    addFiles(newFiles) {
        const existingNames = this.files.map(f => f.name);
        const uniqueFiles = newFiles.filter(f => !existingNames.includes(f.name));
        this.files = [...this.files, ...uniqueFiles];
    }

    getFiles() { return [...this.files]; }
    clearFiles() { 
        this.files = [];
        this.statements = [];
        this.transactions = [];
        this.accounts = {};
    }

    // Управление выписками
    setStatements(statements) { this.statements = statements; }
    getStatements() { return [...this.statements]; }

    // Управление транзакциями
    setTransactions(transactions) { this.transactions = transactions; }
    getTransactions() { return [...this.transactions]; }
    getIncomingTransactions() { return this.transactions.filter(t => t.direction === 'incoming'); }
    getOutgoingTransactions() { return this.transactions.filter(t => t.direction === 'outgoing'); }

    // Управление счетами
    setAccounts(accounts) { this.accounts = accounts; }
    getAccounts() { return { ...this.accounts }; }
    updateAccount(account, data) {
        // Merge: не затираем уже известные значения (например, остаток из выписки)
        // пустыми значениями из нового прохода парсера.
        const merged = { ...(this.accounts[account] || {}) };
        Object.entries(data || {}).forEach(([key, value]) => {
            if (value === undefined || value === null || value === '') return;
            merged[key] = value;
        });
        this.accounts[account] = merged;
    }

    // Управление ИНН
    setINNData(data) { this.innData = data; }
    getINNData() { return { ...this.innData }; }
    getCompanyByINN(inn) { return this.innData[inn]; }

    // Управление депозитами
    setDepositData(data) { this.depositData = data; }
    getDepositData() { return { ...this.depositData }; }
    getDepositForAccount(account) {
        const raw = this.depositData[account];
        if (!raw) return [];
        // Обратная совместимость: старый формат {amount,rate,startDate} → массив
        if (!Array.isArray(raw)) return [raw];
        return raw;
    }
    setDepositForAccount(account, data) {
        // data может быть массивом или объектом (обратная совместимость)
        this.depositData[account] = Array.isArray(data) ? data : [data];
        this.saveRegularDepositsToLocal();
    }

    // ── Обычные депозиты (без endDate, загружаются из Excel) ────────────────
    // В БД они не сохраняются: у term_deposits поле end_date NOT NULL, а API
    // синхронизации отправляет только срочные (getDepositForAccount().filter(d => d.endDate)).
    // Поэтому держим их в localStorage, иначе после перезагрузки страницы они
    // пропадают из остатков (тело и проценты обычного депозита перестают учитываться).
    saveRegularDepositsToLocal() {
        try {
            const regular = {};
            Object.entries(this.depositData).forEach(([account, deposits]) => {
                const list = (Array.isArray(deposits) ? deposits : [deposits])
                    .filter(d => d && !d.endDate && (d.amount || 0) > 0);
                if (list.length > 0) regular[account] = list;
            });
            localStorage.setItem('regularDeposits', JSON.stringify(regular));
        } catch (e) {
            console.warn('Не удалось сохранить обычные депозиты в localStorage:', e);
        }
    }

    loadRegularDepositsFromLocal() {
        try {
            const stored = localStorage.getItem('regularDeposits');
            return stored ? JSON.parse(stored) : {};
        } catch (e) {
            console.warn('Не удалось прочитать обычные депозиты из localStorage:', e);
            return {};
        }
    }

    // Возвращает обычные депозиты в состояние после загрузки срочных из БД
    restoreRegularDeposits() {
        const regular = this.loadRegularDepositsFromLocal();
        let restored = 0;
        Object.entries(regular).forEach(([account, list]) => {
            if (!Array.isArray(list) || list.length === 0) return;
            // срочные из БД сохраняем, обычные подставляем из localStorage
            const longTerm = this.getDepositForAccount(account).filter(d => d && d.endDate);
            this.depositData[account] = [...longTerm, ...list];
            restored += list.length;
        });
        if (restored > 0) {
            console.log(`Восстановлены обычные депозиты из localStorage: ${restored} записей на ${Object.keys(regular).length} счетах`);
        }
    }

    // Сохранение срочных депозитов в БД через API (с fallback на localStorage)
    // accountsToSync — счета, которых уже нет в локальном depositData,
    // но состояние которых всё равно надо записать в БД (например, счета,
    // полностью очищенные автоочисткой истекших депозитов).
    async syncDepositsToServer(accountsToSync = []) {
        // Собираем срочные депозиты (с датой окончания) по всем счетам локального состояния.
        // ВАЖНО: счёт попадает в payload даже с пустым массивом. Бэкенд
        // (db.sync_term_deposits) удаляет записи только по счетам из payload, поэтому
        // без этого удаление последнего срочного депозита не сохранилось бы в БД,
        // и депозит «воскресал» бы при следующем открытии программы.
        const accounts = new Set([...Object.keys(this.depositData), ...accountsToSync.filter(Boolean)]);
        const termDeposits = {};
        accounts.forEach(account => {
            // getDepositForAccount нормализует старый формат {amount, rate, startDate} в массив
            termDeposits[account] = this.getDepositForAccount(account).filter(d => d && d.endDate);
        });

        const accountCount = Object.keys(termDeposits).length;
        const totalDeps = Object.values(termDeposits).reduce((s, arr) => s + arr.length, 0);
        console.log(`syncDepositsToServer: ${totalDeps} срочных депозитов на ${accountCount} счетах`, termDeposits);

        // Fallback: всегда сохраняем в localStorage
        try {
            localStorage.setItem('termDeposits', JSON.stringify(termDeposits));
            console.log('Срочные депозиты сохранены в localStorage');
        } catch (e) {
            console.warn('Не удалось сохранить в localStorage:', e);
        }

        // Основной путь: синхронизация с БД (merge, не полная замена)
        try {
            const resp = await fetch('/api/term-deposits/sync', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ deposits: termDeposits })
            });
            const result = await resp.json();
            if (result.success) {
                console.log(`Срочные депозиты синхронизированы с БД: ${result.synced} записей`);
            } else {
                console.warn('Синхронизация с БД не удалась, данные в localStorage');
            }
        } catch (e) {
            console.warn('БД недоступна — срочные депозиты только в localStorage:', e.message);
        }
    }

    // Загрузка срочных депозитов из БД (с fallback на localStorage)
    async loadDepositsFromServer() {
        // Сначала пробуем БД
        try {
            const resp = await fetch('/api/term-deposits');
            const result = await resp.json();
            if (result.success && result.data) {
                const fromDb = result.data;
                const count = Object.values(fromDb).reduce((s, arr) => s + (Array.isArray(arr) ? arr.length : 0), 0);
                if (count > 0) {
                    console.log(`Загружены срочные депозиты из БД: ${count} записей на ${Object.keys(fromDb).length} счетах`);
                    return fromDb;
                }
            }
        } catch (e) {
            console.warn('Не удалось загрузить депозиты из БД:', e.message);
        }

        // Fallback: localStorage
        try {
            const stored = localStorage.getItem('termDeposits');
            if (stored) {
                const parsed = JSON.parse(stored);
                const count = Object.values(parsed).reduce((s, arr) => s + (Array.isArray(arr) ? arr.length : 0), 0);
                if (count > 0) {
                    console.log(`Загружены срочные депозиты из localStorage: ${count} записей`);
                    return parsed;
                }
            }
        } catch (e) {
            console.warn('Не удалось загрузить из localStorage:', e);
        }

        return {};
    }

    // Автоочистка истекших срочных депозитов
    // endDate — первый день, когда депозит НЕ действует
    // Удаляем только на следующий день после endDate (когда today > endDate)
    async cleanExpiredDeposits() {
        const today = new Date().toISOString().split('T')[0];
        let cleaned = 0;
        const clearedAccounts = []; // счета, у которых не осталось ни одного депозита

        Object.entries(this.depositData).forEach(([account, deposits]) => {
            const list = Array.isArray(deposits) ? deposits : [deposits];
            // endDate — последний день действия, today > endDate означает что депозит истёк
            const active = list.filter(d => !d.endDate || d.endDate >= today);
            if (active.length !== list.length) {
                cleaned += list.length - active.length;
                if (active.length > 0) {
                    this.depositData[account] = active;
                } else {
                    delete this.depositData[account];
                    clearedAccounts.push(account);
                }
            }
        });

        if (cleaned > 0) {
            console.log(`Автоочистка: удалено ${cleaned} истекших депозитов`);
            // Очищенные счета передаём явно: из depositData они уже удалены,
            // но их записи в БД тоже надо удалить.
            await this.syncDepositsToServer(clearedAccounts);
            // держим localStorage с обычными депозитами в актуальном состоянии
            this.saveRegularDepositsToLocal();
        }
    }
    getExclusionRules() { return [...this.exclusionRules]; }

    // Управление правилами категоризации
    setCategorizationRules(rules) { this.categorizationRules = rules; }
    getCategorizationRules() { return [...this.categorizationRules]; }

    // Управление синонимами компаний
    setCompanyAliases(aliases) { this.companyAliases = aliases; }
    getCompanyAliases() { return [...this.companyAliases]; }

    // Загрузка конфигурации с API
    async loadConfig() {
        try {
            const resp = await fetch('/api/config');
            const cfg = await resp.json();
            if (cfg.success) {
                const mapping = cfg.data.account_mapping || {};
                // ВАЖНО: не перезаписываем счета целиком, а сливаем маппинг с уже известными
                // данными (из выписок: balance, date). Иначе поздний ответ /api/config
                // затирает остатки, загруженные при обработке выписок, и в разделе
                // «Остатки» пропадают данные по счетам.
                const merged = { ...this.accounts };
                Object.entries(mapping).forEach(([account, data]) => {
                    merged[account] = { ...(merged[account] || {}), ...(data || {}) };
                });
                this.accounts = merged;
                this.exclusionRules = cfg.data.exclusion_rules || [];
                this.categorizationRules = cfg.data.categorization_rules || [];
                this.companyAliases = cfg.data.company_aliases || [];
                return true;
            }
        } catch (e) { console.warn('loadConfig failed:', e); }
        return false;
    }

    // Вспомогательные методы
    formatCurrency(amount) {
        return new Intl.NumberFormat('ru-RU', {
            style: 'currency',
            currency: 'RUB',
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(amount);
    }

    formatNumber(num) {
        return new Intl.NumberFormat('ru-RU', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(num);
    }

    parseDate(dateStr) {
        console.log('Парсинг даты:', dateStr);
        if (!dateStr) return new Date();
        
        // Формат DD.MM.YYYY
        const parts = dateStr.split('.');
        if (parts.length === 3) {
            const date = new Date(parts[2], parts[1] - 1, parts[0]);
            console.log('Парсинг DD.MM.YYYY результат:', date);
            return date;
        }
        
        // Формат YYYY-MM-DD
        const isoMatch = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (isoMatch) {
            const date = new Date(isoMatch[1], isoMatch[2] - 1, isoMatch[3]);
            console.log('Парсинг YYYY-MM-DD результат:', date);
            return date;
        }
        
        // Формат MM/DD/YYYY или другие
        const date = new Date(dateStr);
        console.log('Парсинг через new Date результат:', date);
        
        // Проверяем, что дата валидна
        if (isNaN(date.getTime())) {
            console.log('Неверный формат даты, возвращаем текущую дату');
            return new Date();
        }
        
        return date;
    }

    getDaysBetween(startDate, endDate) {
        console.log('Расчет дней между:', { startDate, endDate });
        const start = this.parseDate(startDate);
        const end = this.parseDate(endDate);
        console.log('Даты после парсинга:', { start, end });
        const diffTime = Math.abs(end - start);
        const days = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        console.log('Результат дней:', days);
        return days;
    }

    calculateInterest(amount, rate, days) {
        console.log('Расчет процентов:', { amount, rate, days });
        if (!amount || !rate || !days) return 0;
        
        // Проверяем типы данных
        amount = parseFloat(amount);
        rate = parseFloat(rate);
        days = parseInt(days);
        
        console.log('После преобразования:', { amount, rate, days });
        
        const result = (amount * rate * days) / (100 * 365);
        console.log('Результат расчета:', result);
        
        return result;
    }
}
