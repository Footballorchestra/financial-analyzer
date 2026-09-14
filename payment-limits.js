// payment-limits.js — Модуль «Контроль лимитов платежей поставщикам»
// Сверка ведётся по колонкам файла «Тип оплаты» (категория лимита) и «Детальный лимит» (подстатья).
// Период определяется по дате загрузки/подтверждения. Суммы — в рублях «1 234 567,89».
// Лимиты (названия и значения) настраиваются в интерфейсе и хранятся в БД.

class PaymentLimitsManager {
    constructor(storage) {
        this.storage = storage;
        this.previewRows = [];       // строки текущей (неподтверждённой) загрузки
        this.confirmedRows = [];     // подтверждённые платежи месяца из БД
        this.currentPeriod = 1;      // 1..4 — рабочий период (выбранный или по дате)
        this.selectedPeriod = null;  // null — авто (по дате загрузки), 1..4 — выбран вручную
        this.currentMonth = null;    // 'YYYY-MM-01'
        this.fileName = null;
        this.confirmedLoaded = false;
        this.config = null;          // конфигурация лимитов (с сервера или defaults)
        this.configLoaded = false;
        this.onlyCurrentPeriod = this.loadOnlyCurrentPeriodSetting(); // режим «только текущий период»
        this.previewFilter = { category: null, detail: null };        // фильтр списка заявок
    }

    static PERIODS = [1, 2, 3, 4];
    static PERIOD_DAYS = { 1: '1–7', 2: '8–15', 3: '16–22', 4: '23–31' };

    // Значения по умолчанию (используются, если настройки недоступны). На сервере — те же defaults.
    static DEFAULT_CONFIG = {
        globalLimitP12: 31000000,
        categories: [
            { id: 'cat_aks', name: 'Аренда, Коммуналка, Связь',
              limits: { 1: 6500000, 2: 1000000, 3: 7000000, 4: 0 },
              details: [
                  { id: 'det_internet', name: 'Связь и интернет', limits: { 1: 3500000, 2: 0, 3: 0, 4: 0 } },
                  { id: 'det_utility', name: 'Коммунальные услуги', limits: { 1: 1500000, 2: 1000000, 3: 0, 4: 0 } },
                  { id: 'det_rent_flat', name: 'Аренда квартир', limits: { 1: 1500000, 2: 0, 3: 0, 4: 0 } },
                  { id: 'det_rent_office', name: 'Аренда баз и офисов', limits: { 1: 0, 2: 0, 3: 7000000, 4: 0 } }
              ] },
            { id: 'cat_life', name: 'Жизнеобеспечение', limits: { 1: 6500000, 2: 5000000, 3: 7000000, 4: 7000000 }, details: [] },
            { id: 'cat_fleet', name: 'Автопарк', limits: { 1: 4500000, 2: 2500000, 3: 20000000, 4: 14000000 }, details: [] },
            { id: 'cat_other', name: 'Прочая операционка', limits: { 1: 1500000, 2: 1500000, 3: 10000000, 4: 10500000 }, details: [] },
            { id: 'cat_overdue', name: 'Графики просрочки', limits: { 1: 500000, 2: 500000, 3: 5000000, 4: 19000000 }, details: [] }
        ]
    };

    // Название периода для интерфейса: «Период 1»..«Период 4»
    periodLabel(p) {
        return `Период ${p}`;
    }

    // ===== ПЕРИОД ПО ДАТЕ ЗАГРУЗКИ =====
    definePeriod(date) {
        const d = date || new Date();
        const day = d.getDate();
        if (day <= 7) return 1;
        if (day <= 15) return 2;
        if (day <= 22) return 3;
        return 4;
    }

    // Рабочий период загрузки: выбранный вручную или авто по дате
    getLoadPeriod() {
        return this.selectedPeriod || this.definePeriod(new Date());
    }

    // Установка периода загрузки. period: ''/null — авто по дате, 1..4 — вручную
    setLoadPeriod(period) {
        const p = parseInt(period, 10);
        this.selectedPeriod = (p >= 1 && p <= 4) ? p : null;
        this.currentPeriod = this.getLoadPeriod();
        return this.currentPeriod;
    }

    monthPrefix(date) {
        const d = date || new Date();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        return `${d.getFullYear()}-${m}-01`;
    }

    monthHuman(date) {
        const d = date || new Date();
        const months = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
                        'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
        return `${months[d.getMonth()]} ${d.getFullYear()}`;
    }

    // ===== ЗАГРУЗКА ФАЙЛА (ПРЕВЬЮ) =====
    async loadExcelFile(file) {
        if (!file) return { success: false, message: 'Файл не выбран' };
        try {
            const workbook = await this.readExcelFile(file);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
            if (rows.length < 2) {
                return { success: false, message: 'В файле нет данных' };
            }

            const headers = rows[0].map(h => String(h == null ? '' : h).trim());
            const columns = this.detectColumns(headers);
            if (columns.category === undefined || columns.amount === undefined) {
                return { success: false, message: 'Не найдены обязательные колонки: "Тип оплаты" и "Сумма"' };
            }

            const preview = [];
            rows.slice(1).forEach(row => {
                if (!row.some(cell => cell !== null && cell !== '')) return;
                const amount = this.parseAmount(row[columns.amount]);
                if (amount === 0) return;
                preview.push({
                    category: this.cleanString(row[columns.category]),
                    amount,
                    articleDds: columns.article !== undefined ? this.cleanString(row[columns.article]) : '',
                    recipient: this.cleanString(row[columns.recipient]),
                    detailLimit: columns.detailLimit !== undefined ? this.cleanString(row[columns.detailLimit]) : '',
                    requestNo: columns.request !== undefined ? this.cleanString(row[columns.request]) : '',
                    requestDate: columns.requestDate !== undefined ? this.formatExcelDate(row[columns.requestDate]) : null,
                    status: columns.status !== undefined ? this.cleanString(row[columns.status]) : '',
                    organization: columns.organization !== undefined ? this.cleanString(row[columns.organization]) : '',
                    subdivision: columns.subdivision !== undefined ? this.cleanString(row[columns.subdivision]) : '',
                    priority: columns.priority !== undefined ? this.cleanString(row[columns.priority]) : '',
                    purpose: columns.purpose !== undefined ? this.cleanString(row[columns.purpose]) : '',
                    paymentDate: columns.paymentDate !== undefined ? this.formatExcelDate(row[columns.paymentDate]) : null,
                    applicant: columns.applicant !== undefined ? this.cleanString(row[columns.applicant]) : '',
                    paymentType: columns.paymentType !== undefined ? this.cleanString(row[columns.paymentType]) : ''
                });
            });

            if (!preview.some(r => r.category)) {
                return { success: false, message: 'В файле не заполнена колонка "Тип оплаты"' };
            }

            const now = new Date();
            this.previewRows = preview;
            // Период загрузки: выбранный пользователем либо авто по дате загрузки
            this.currentPeriod = this.getLoadPeriod();
            this.currentMonth = this.monthPrefix(now);
            this.fileName = file.name;
            return { success: true, count: preview.length, message: `Прочитано строк: ${preview.length}` };
        } catch (error) {
            console.error('Ошибка загрузки файла лимитов:', error);
            return { success: false, message: 'Ошибка загрузки: ' + error.message };
        }
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                    resolve(workbook);
                } catch (error) {
                    reject(new Error('Ошибка чтения Excel: ' + error.message));
                }
            };
            reader.onerror = () => reject(new Error('Ошибка чтения файла'));
            reader.readAsArrayBuffer(file);
        });
    }

    // ===== ОПРЕДЕЛЕНИЕ КОЛОНОК =====
    detectColumns(headers) {
        const candidates = [];
        headers.forEach((header, index) => { candidates.push({ index, text: header.toLowerCase() }); });
        const find = (patterns) => {
            for (const c of candidates) for (const p of patterns) if (c.text === p.toLowerCase()) return c.index;
            let best = null;
            for (const c of candidates) for (const p of patterns) {
                if (c.text.includes(p.toLowerCase())) { if (!best || p.length > best.len) best = { index: c.index, len: p.length }; }
            }
            return best ? best.index : undefined;
        };
        return {
            category: find(['тип оплаты']),
            article: find(['статья движения денежных средств', 'статья ддс']),
            amount: find(['сумма']),
            recipient: find(['получатель']),
            detailLimit: find(['детальный лимит', 'детальный']),
            request: find(['номер заявки', '№ заявки', 'заявка']),
            requestDate: find(['дата заявки', 'дата подачи заявки']),
            status: find(['статус']),
            organization: find(['организация']),
            subdivision: find(['подразделение']),
            priority: find(['приоритет']),
            purpose: find(['назначение платежа', 'назначение']),
            paymentDate: find(['дата оплаты', 'дата платежа']),
            applicant: find(['заявитель']),
            paymentType: find(['тип оплаты'])
        };
    }

    // ===== КОНФИГУРАЦИЯ ЛИМИТОВ =====
    async loadConfig() {
        try {
            const response = await fetch('/api/payment-limits/config');
            const result = await response.json();
            if (result.success && result.config) {
                this.config = result.config;
                this.configLoaded = true;
                return { success: true };
            }
            throw new Error(result.error || 'Некорректный ответ сервера');
        } catch (error) {
            console.warn('Не удалось загрузить настройки лимитов, используются значения по умолчанию:', error);
            if (!this.config) this.config = JSON.parse(JSON.stringify(PaymentLimitsManager.DEFAULT_CONFIG));
            this.configLoaded = true;
            return { success: false, message: error.message };
        }
    }

    async saveConfig(config) {
        try {
            const response = await fetch('/api/payment-limits/config', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ config })
            });
            const result = await response.json();
            if (!response.ok || !result.success) {
                return { success: false, message: result.error || result.message || 'Ошибка сохранения настроек' };
            }
            this.config = result.config || config;
            this.configLoaded = true;
            return result;
        } catch (error) {
            console.error('Ошибка сохранения настроек лимитов:', error);
            return { success: false, message: error.message };
        }
    }

    getConfig() {
        if (!this.config) this.config = JSON.parse(JSON.stringify(PaymentLimitsManager.DEFAULT_CONFIG));
        return this.config;
    }

    getCategories() {
        const cats = this.getConfig().categories;
        return Array.isArray(cats) ? cats : [];
    }

    getCategory(categoryName) {
        return this.getCategories().find(c => c.name === categoryName) || null;
    }

    getLimits(categoryName, period) {
        const c = this.getCategory(categoryName);
        if (!c || !c.limits) return 0;
        return Number(c.limits[String(period)]) || 0;
    }

    getDetails(categoryName) {
        const c = this.getCategory(categoryName);
        return c && Array.isArray(c.details) ? c.details : [];
    }

    getGlobalLimitP12() {
        return Number(this.getConfig().globalLimitP12) || 0;
    }

    // ===== РЕЖИМ «ТОЛЬКО ТЕКУЩИЙ ПЕРИОД» (без переноса неиспользованного остатка) =====
    loadOnlyCurrentPeriodSetting() {
        try {
            return localStorage.getItem('paymentLimitsOnlyCurrentPeriod') === 'true';
        } catch (e) {
            return false;
        }
    }

    setOnlyCurrentPeriod(value) {
        this.onlyCurrentPeriod = !!value;
        try {
            localStorage.setItem('paymentLimitsOnlyCurrentPeriod', this.onlyCurrentPeriod ? 'true' : 'false');
        } catch (e) { /* localStorage недоступен — не критично */ }
    }

    // Расчёт строки контроля лимита (учитывает режим «только текущий период»)
    cumulativeRow(categoryName, conf, prev) {
        const cur = this.currentPeriod;
        let total = 0;
        let limit = 0;

        if (this.onlyCurrentPeriod) {
            // Доступен только лимит текущего периода; расход — за текущий период
            total = ((conf[categoryName] && conf[categoryName][cur]) || 0)
                + ((prev[categoryName] && prev[categoryName][cur]) || 0);
            limit = this.getLimits(categoryName, cur);
        } else {
            // Накопительно: подтверждено за периоды 1..cur + превью текущего периода
            PaymentLimitsManager.PERIODS.forEach(p => {
                if (p <= cur) {
                    total += (conf[categoryName] && conf[categoryName][p]) || 0;
                    limit += this.getLimits(categoryName, p);
                }
            });
            if (prev[categoryName] && prev[categoryName][cur]) {
                total += prev[categoryName][cur] || 0;
            }
        }
        return { total, limit };
    }

    // ===== НОРМАЛИЗАЦИЯ (по названиям из настроек) =====
    normalizeCategory(raw) {
        if (!raw) return '';
        const s = raw.toLowerCase().replace(/ё/g, 'е').trim();
        const cats = this.getCategories();
        // 1) точное совпадение
        for (const c of cats) {
            if (this.cleanString(c.name).toLowerCase().replace(/ё/g, 'е') === s) return c.name;
        }
        // 2) вхождение подстроки
        let best = null;
        for (const c of cats) {
            const n = this.cleanString(c.name).toLowerCase().replace(/ё/g, 'е');
            if (n && (s.includes(n) || n.includes(s))) {
                if (!best || n.length > best.len) best = { name: c.name, len: n.length };
            }
        }
        if (best) return best.name;
        // 3) эвристики для стандартных названий
        const heuristics = [
            { words: ['аренда, коммуналка, связь', 'аренда коммуналка связь'], idx: 0 },
            { words: ['жизнеобеспечение'], idx: 1 },
            { words: ['автопарк'], idx: 2 },
            { words: ['прочая операционка', 'прочие операционн'], idx: 3 },
            { words: ['графики просрочки', 'график просрочки'], idx: 4 }
        ];
        for (const h of heuristics) {
            if (h.words.some(w => s.includes(w)) && cats[h.idx]) return cats[h.idx].name;
        }
        if ((s.includes('коммуналк') || (s.includes('аренд') && s.includes('связ'))) && cats[0]) return cats[0].name;
        return this.cleanString(raw);
    }

    normalizeDetailLimit(raw) {
        if (!raw) return '';
        const s = raw.toLowerCase().trim();
        const allDetails = [];
        this.getCategories().forEach(c => (c.details || []).forEach(d => allDetails.push(d.name)));
        for (const name of allDetails) if (this.cleanString(name).toLowerCase() === s) return name;
        for (const name of allDetails) {
            const n = this.cleanString(name).toLowerCase();
            if (n && s.includes(n)) return name;
        }
        const heuristics = [
            { words: ['связь и интернет', 'интернет', 'связь'], name: 'Связь и интернет' },
            { words: ['коммунальн'], name: 'Коммунальные услуги' },
            { words: ['квартир'], name: 'Аренда квартир' },
            { words: ['офисов', 'баз и офисов', 'офис'], name: 'Аренда баз и офисов' }
        ];
        for (const h of heuristics) if (h.words.some(w => s.includes(w))) return h.name;
        return this.cleanString(raw);
    }

    rowCategory(row) {
        return this.normalizeCategory(row.category || row.article || '');
    }

    // ===== АГРЕГАЦИИ (рубли) =====
    aggregatePreview(rows, targetPeriod) {
        const result = {};
        rows.forEach(row => {
            const cat = this.rowCategory(row);
            if (!cat) return;
            if (!result[cat]) result[cat] = { 1: 0, 2: 0, 3: 0, 4: 0 };
            const p = targetPeriod || this.currentPeriod;
            result[cat][p] = (result[cat][p] || 0) + (Number(row.amount) || 0);
        });
        return result;
    }

    confirmedByPeriod() {
        const result = {};
        this.confirmedRows.forEach(r => {
            const cat = this.normalizeCategory(r.article);
            if (!cat) return;
            if (!result[cat]) result[cat] = { 1: 0, 2: 0, 3: 0, 4: 0 };
            const p = Number(r.period) || this.currentPeriod;
            result[cat][p] = (result[cat][p] || 0) + (Number(r.amount) || 0);
        });
        return result;
    }

    // Детальные суммы по подстатьям указанной категории (confirmed или preview)
    detailByPeriod(rows, categoryName, targetPeriod) {
        const result = {};
        rows.forEach(row => {
            const cat = this.rowCategory(row);
            if (cat !== categoryName) return;
            const sub = this.normalizeDetailLimit(row.detail_limit !== undefined ? row.detail_limit : row.detailLimit);
            if (!sub) return;
            if (!result[sub]) result[sub] = { 1: 0, 2: 0, 3: 0, 4: 0 };
            const p = targetPeriod || Number(row.period) || this.currentPeriod;
            result[sub][p] = (result[sub][p] || 0) + (Number(row.amount) || 0);
        });
        return result;
    }

    // ===== ЗАГРУЗКА ПОДТВЕРЖДЁННЫХ ДАННЫХ ИЗ БД =====
    async loadConfirmed(month) {
        try {
            const url = month ? `/api/payment-limits/confirmed?month=${month}` : '/api/payment-limits/confirmed';
            const response = await fetch(url);
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            const result = await response.json();
            if (result.success) {
                this.confirmedRows = result.rows || [];
                this.currentMonth = result.month || this.currentMonth || this.monthPrefix(new Date());
                // Период не сбрасываем: он выбранный вручную либо авто по дате
                this.currentPeriod = this.getLoadPeriod();
                this.confirmedRows.forEach(r => { r.amount = Number(r.amount) || 0; });
                this.confirmedLoaded = true;
                return { success: true, rows: this.confirmedRows };
            }
            return { success: false, message: result.error || 'Не удалось загрузить подтверждённые платежи' };
        } catch (error) {
            console.warn('Не удалось загрузить подтверждённые платежи:', error);
            return { success: false, message: error.message };
        }
    }

    // ===== ПОДТВЕРЖДЕНИЕ =====
    async confirmSave() {
        if (!this.previewRows.length) return { success: false, message: 'Нет данных для подтверждения' };
        try {
            const rows = this.previewRows.map(r => ({
                article: r.category || r.articleDds || '',
                detail_limit: r.detailLimit || '',
                amount: r.amount,
                recipient: r.recipient || '',
                request_no: r.requestNo || '',
                request_date: r.requestDate || null,
                status: r.status || '',
                organization: r.organization || '',
                subdivision: r.subdivision || '',
                priority: r.priority || '',
                purpose: r.purpose || '',
                payment_date: r.paymentDate || null,
                applicant: r.applicant || '',
                payment_type: r.paymentType || ''
            }));
            const response = await fetch('/api/payment-limits/confirm', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                // period — выбранный пользователем период загрузки (иначе сервер возьмёт по дате)
                body: JSON.stringify({ rows, period: this.getLoadPeriod() })
            });
            const result = await response.json();
            if (!response.ok || !result.success) {
                return { success: false, message: result.error || result.message || 'Ошибка подтверждения' };
            }
            return result;
        } catch (error) {
            console.error('Ошибка подтверждения:', error);
            return { success: false, message: error.message };
        }
    }

    cancelPreview() {
        this.previewRows = [];
        this.fileName = null;
        this.previewFilter = { category: null, detail: null };
    }

    // ===== ФИЛЬТР СПИСКА ЗАЯВОК (по категории / подстатье) =====
    setPreviewFilter(category, detail) {
        this.previewFilter = { category: category || null, detail: detail || null };
    }

    hasPreviewFilter() {
        return !!(this.previewFilter.category || this.previewFilter.detail);
    }

    filteredPreviewRows() {
        const f = this.previewFilter;
        return this.previewRows.filter(row => {
            if (f.category && this.rowCategory(row) !== f.category) return false;
            if (f.detail && this.normalizeDetailLimit(row.detailLimit) !== f.detail) return false;
            return true;
        });
    }

    // Обновление только блока списка заявок (без полной перерисовки страницы)
    refreshPreviewList() {
        const el = document.getElementById('plPreviewListBlock');
        if (!el) return;
        el.outerHTML = this.previewRowsHtml();
    }

    // ===== ФОРМАТИРОВАНИЕ =====
    cleanString(value) {
        if (!value) return '';
        return String(value).trim();
    }

    parseAmount(value) {
        if (!value && value !== 0) return 0;
        if (typeof value === 'number') return value;
        const cleaned = String(value).replace(/[\s\u00A0]/g, '').replace(',', '.');
        const num = parseFloat(cleaned);
        return isNaN(num) ? 0 : num;
    }

    formatExcelDate(value) {
        if (value === null || value === undefined || value === '') return null;
        if (value instanceof Date) return value.toISOString().slice(0, 10);
        if (typeof value === 'number') {
            const d = new Date(Math.round((value - 25569) * 86400 * 1000));
            return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
        }
        const s = String(value).trim();
        if (!s) return null;
        let m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
        if (m) { let y = m[3]; if (y.length === 2) y = '20' + y; return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
        m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
        return s.slice(0, 10);
    }

    // «1 234 567,89»
    fmtMoney(value) {
        const v = Number(value) || 0;
        const sign = v < 0 ? '-' : '';
        const s = Math.abs(v).toFixed(2).replace('.', ',');
        const parts = s.split(',');
        return sign + parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ',' + parts[1];
    }

    fmtPct(value) {
        const v = Number(value) || 0;
        if (!isFinite(v)) return '∞';
        return v.toLocaleString('ru-RU', { maximumFractionDigits: 1 });
    }

    escapeHtml(str) {
        if (str === null || str === undefined) return '';
        return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
    }

    barWidth(pct) {
        const v = Number(pct) || 0;
        if (!isFinite(v)) return 100;
        return Math.min(Math.max(v, 0), 100);
    }

    barClass(percent) {
        if (percent >= 100) return 'limit-over';
        if (percent >= 90) return 'limit-warn';
        return 'limit-ok';
    }

    // ===== UI =====
    renderHeaderInfo() {
        const now = new Date();
        const elDate = document.getElementById('plHeaderDate');
        const elPeriod = document.getElementById('plHeaderPeriod');
        const elMonth = document.getElementById('plHeaderMonth');
        if (elDate) elDate.textContent = now.toLocaleDateString('ru-RU');
        if (elPeriod) {
            const p = this.getLoadPeriod();
            const mode = this.selectedPeriod ? 'выбрано вручную' : 'авто по дате';
            elPeriod.textContent = `${this.periodLabel(p)} (${PaymentLimitsManager.PERIOD_DAYS[p]} числа) · ${mode}`;
        }
        if (elMonth) elMonth.textContent = this.monthHuman(now);
    }

    periodThHtml(p) {
        const cur = this.currentPeriod;
        return `<th>${p === cur ? '<span class="pl-current-period">▶</span> ' : ''}${this.periodLabel(p)}
            <small>${PaymentLimitsManager.PERIOD_DAYS[p]} числа</small></th>`;
    }

    cellHtml(confirmedRub, previewRub, limitRub) {
        const total = (confirmedRub || 0) + (previewRub || 0);
        let pct;
        if (limitRub > 0) { pct = (total / limitRub) * 100; } else { pct = total > 0 ? Infinity : 0; }
        const cls = this.barClass(pct);
        const pctText = isFinite(pct) ? `${this.fmtPct(pct)}%` : '∞';
        return `
            <div class="pl-cell">
                <div class="pl-amounts"><b>${this.fmtMoney(total)}</b>
                    <span class="pl-of">/ ${limitRub > 0 ? this.fmtMoney(limitRub) : '—'}</span></div>
                <div class="limit-bar"><div class="limit-bar-fill ${cls}" style="width:${this.barWidth(pct)}%;"></div></div>
                <div class="pl-meta">БД ${this.fmtMoney(confirmedRub || 0)} · Превью ${this.fmtMoney(previewRub || 0)} · ${pctText}</div>
            </div>`;
    }

    cellHtmlTotal(totalRub, limitRub) {
        const total = Number(totalRub) || 0;
        let pct;
        if (limitRub > 0) { pct = (total / limitRub) * 100; } else { pct = total > 0 ? Infinity : 0; }
        const cls = this.barClass(pct);
        const pctText = isFinite(pct) ? `${this.fmtPct(pct)}%` : '∞';
        return `
            <div class="pl-cell">
                <div class="pl-amounts"><b>${this.fmtMoney(total)}</b>
                    <span class="pl-of">/ ${limitRub > 0 ? this.fmtMoney(limitRub) : '—'}</span></div>
                <div class="limit-bar"><div class="limit-bar-fill ${cls}" style="width:${this.barWidth(pct)}%;"></div></div>
                <div class="pl-meta">${pctText}</div>
            </div>`;
    }

    controlTableHtml() {
        const conf = this.confirmedByPeriod();
        const prev = this.aggregatePreview(this.previewRows, this.currentPeriod);
        const cur = this.currentPeriod;
        const headersHtml = PaymentLimitsManager.PERIODS.map(p => this.periodThHtml(p)).join('');

        let html = `
            <div class="pl-block pl-block-flat">
                <div class="pl-block-title"><i class="fas fa-table"></i> Контроль лимитов по периодам
                    <small>сверка по «Тип оплаты» из файла, суммы в рублях</small></div>
                <div class="data-table-wrapper pl-table-wrap">
                <table class="data-table pl-table">
                    <thead><tr>
                        <th rowspan="2">Тип оплаты (категория лимита)</th>
                        ${headersHtml}
                        <th rowspan="2">За месяц</th>
                    </tr></thead><tbody>`;

        const grandByPeriod = {};
        PaymentLimitsManager.PERIODS.forEach(p => { grandByPeriod[p] = { total: 0, limit: 0 }; });
        let grandMonth = 0;

        const rowsMap = {};
        this.getCategories().forEach(c => { rowsMap[c.name] = true; });
        [...Object.keys(conf), ...Object.keys(prev)].forEach(k => { if (k) rowsMap[k] = true; });

        Object.keys(rowsMap).forEach(cat => {
            const known = !!this.getCategory(cat);
            const activeCat = this.previewFilter.category === cat;
            let rowMonth = 0;
            html += `<tr class="${known ? '' : 'pl-unmatched'}${activeCat ? ' pl-filter-row' : ''}">
                <td class="pl-article-name"><button type="button" class="pl-cat-link" data-pl-category="${this.escapeHtml(cat)}" title="Показать заявки категории «${this.escapeHtml(cat)}» в списке ниже">${this.escapeHtml(cat)}</button>${known ? '' : '<br><small class="pl-warn-label">не сопоставлено с матрицей</small>'}</td>`;
            PaymentLimitsManager.PERIODS.forEach(p => {
                const confirmedRub = (conf[cat] && conf[cat][p]) || 0;
                const previewRub = p === cur && prev[cat] && prev[cat][p] ? (prev[cat][p] || 0) : 0;
                const limitRub = known ? this.getLimits(cat, p) : 0;
                const total = confirmedRub + previewRub;
                grandByPeriod[p].total += total;
                grandByPeriod[p].limit += limitRub;
                rowMonth += total;
                html += `<td class="pl-cell-td">${this.cellHtml(confirmedRub, previewRub, limitRub)}</td>`;
            });
            grandMonth += rowMonth;
            html += `<td class="number-cell pl-month-total">${this.fmtMoney(rowMonth)}</td></tr>`;
        });

        html += `<tr class="pl-grand-row"><td><strong>ИТОГО по всем категориям</strong></td>`;
        PaymentLimitsManager.PERIODS.forEach(p => {
            html += `<td class="pl-cell-td">${this.cellHtmlTotal(grandByPeriod[p].total, grandByPeriod[p].limit)}</td>`;
        });
        html += `<td class="number-cell pl-month-total">${this.fmtMoney(grandMonth)}</td></tr>`;
        html += `</tbody></table></div></div>`;
        return html;
    }

    // Детальный контроль по подстатьям (категории, у которых заданы detail-лимиты)
    detailTablesHtml() {
        const cur = this.currentPeriod;
        const blocks = [];
        this.getCategories().forEach(cat => {
            const details = (cat.details || []);
            if (!details.length) return;
            const conf = this.detailByPeriod(this.confirmedRows, cat.name, null);
            const prev = this.detailByPeriod(this.previewRows, cat.name, this.currentPeriod);
            const aksHasRows = this.previewRows.some(r => this.rowCategory(r) === cat.name)
                || this.confirmedRows.some(r => this.normalizeCategory(r.article) === cat.name);
            if (!Object.keys(conf).length && !Object.keys(prev).length && !aksHasRows) return;

            const grandByPeriod = {};
            PaymentLimitsManager.PERIODS.forEach(p => { grandByPeriod[p] = { total: 0, limit: 0 }; });
            let grandMonth = 0;
            const detailLimits = {};
            details.forEach(d => { detailLimits[d.name] = d.limits || {}; });

            const rowsMap = {};
            details.forEach(d => { rowsMap[d.name] = true; });
            [...Object.keys(conf), ...Object.keys(prev)].forEach(k => { if (k) rowsMap[k] = true; });

            let rowsHtml = '';
            Object.keys(rowsMap).forEach(sub => {
                const known = Object.prototype.hasOwnProperty.call(detailLimits, sub);
                const activeSub = this.previewFilter.category === cat.name && this.previewFilter.detail === sub;
                let rowMonth = 0;
                rowsHtml += `<tr class="${known ? '' : 'pl-unmatched'}${activeSub ? ' pl-filter-row' : ''}"><td class="pl-article-name"><button type="button" class="pl-cat-link" data-pl-category="${this.escapeHtml(cat.name)}" data-pl-detail="${this.escapeHtml(sub)}" title="Показать заявки подстатьи «${this.escapeHtml(sub)}» в списке ниже">${this.escapeHtml(sub)}</button></td>`;
                PaymentLimitsManager.PERIODS.forEach(p => {
                    const confirmedRub = (conf[sub] && conf[sub][p]) || 0;
                    const previewRub = p === cur && prev[sub] && prev[sub][p] ? (prev[sub][p] || 0) : 0;
                    const limitRub = known ? (Number(detailLimits[sub][String(p)]) || 0) : 0;
                    const total = confirmedRub + previewRub;
                    grandByPeriod[p].total += total;
                    grandByPeriod[p].limit += limitRub;
                    rowMonth += total;
                    rowsHtml += `<td class="pl-cell-td">${this.cellHtml(confirmedRub, previewRub, limitRub)}</td>`;
                });
                grandMonth += rowMonth;
                rowsHtml += `<td class="number-cell pl-month-total">${this.fmtMoney(rowMonth)}</td></tr>`;
            });

            let totalRow = `<tr class="pl-grand-row"><td><strong>ИТОГО по подстатьям</strong></td>`;
            PaymentLimitsManager.PERIODS.forEach(p => {
                totalRow += `<td class="pl-cell-td">${this.cellHtmlTotal(grandByPeriod[p].total, grandByPeriod[p].limit)}</td>`;
            });
            totalRow += `<td class="number-cell pl-month-total">${this.fmtMoney(grandMonth)}</td></tr>`;

            blocks.push(`
                <div class="pl-block pl-block-flat">
                    <div class="pl-block-title"><i class="fas fa-sliders-h"></i>
                        Детальный контроль: ${this.escapeHtml(cat.name)}
                        <small>по подстатьям из колонки «Детальный лимит»</small></div>
                    <div class="data-table-wrapper pl-table-wrap">
                    <table class="data-table pl-table">
                        <thead><tr><th rowspan="2">Подстатья</th>${PaymentLimitsManager.PERIODS.map(p => this.periodThHtml(p)).join('')}<th rowspan="2">За месяц</th></tr></thead>
                        <tbody>${rowsHtml}${totalRow}</tbody>
                    </table></div>
                </div>`);
        });
        return blocks.join('');
    }

    summaryLimitsHtml() {
        const conf = this.confirmedByPeriod();
        const prev = this.aggregatePreview(this.previewRows, this.currentPeriod);
        const cur = this.currentPeriod;

        const rowData = this.getCategories().map(cat => {
            const { total, limit } = this.cumulativeRow(cat.name, conf, prev);
            return { name: cat.name, total, limit };
        });
        const grand = rowData.reduce((a, r) => { a.total += r.total; a.limit += r.limit; return a; }, { total: 0, limit: 0 });

        const rowHtml = (r) => {
            const pct = r.limit > 0 ? (r.total / r.limit) * 100 : (r.total > 0 ? Infinity : 0);
            const cls = this.barClass(pct);
            return `<tr><td>${this.escapeHtml(r.name)}</td>
                <td class="number-cell">${this.fmtMoney(r.total)}</td>
                <td class="number-cell">${r.limit > 0 ? this.fmtMoney(r.limit) : '—'}</td>
                <td class="number-cell">${r.limit > 0 ? this.fmtMoney(Math.max(r.limit - r.total, 0)) : '—'}</td>
                <td class="pl-bar-cell"><div class="limit-bar"><div class="limit-bar-fill ${cls}" style="width:${this.barWidth(pct)}%;"></div></div><span class="pl-meta">${isFinite(pct) ? this.fmtPct(pct) + '%' : '∞'}</span></td></tr>`;
        };

        let totalP1 = 0, totalP2 = 0;
        this.getCategories().forEach(cat => {
            const c = conf[cat.name] || {};
            totalP1 += c[1] || 0;
            totalP2 += c[2] || 0;
            if (cur === 1 && prev[cat.name] && prev[cat.name][1]) totalP1 += prev[cat.name][1] || 0;
            if (cur === 2 && prev[cat.name] && prev[cat.name][2]) totalP2 += prev[cat.name][2] || 0;
        });
        const sum12 = totalP1 + totalP2;
        const limit12 = this.getGlobalLimitP12();
        const pct12 = limit12 > 0 ? (sum12 / limit12) * 100 : (sum12 > 0 ? Infinity : 0);
        const cls12 = this.barClass(pct12);
        const grandPct = grand.limit > 0 ? (grand.total / grand.limit) * 100 : (grand.total > 0 ? Infinity : 0);
        const grandCls = this.barClass(grandPct);

        return `
            <div class="pl-block pl-block-flat">
                <div class="pl-block-title"><i class="fas fa-chart-line"></i> ${this.onlyCurrentPeriod ? 'Лимит текущего периода' : 'Накопительный итог и лимиты'}
                    <small>${this.onlyCurrentPeriod
                        ? `режим «только текущий период» (${this.periodLabel(cur)}, без переноса неиспользованного остатка)`
                        : `подтверждено за периоды 1–${cur} + превью текущей загрузки`}</small></div>
                <div class="data-table-wrapper pl-table-wrap">
                <table class="data-table pl-table-narrow">
                    <thead><tr><th>Тип оплаты</th><th>Итого, ₽</th><th>Доступный лимит, ₽</th><th>Остаток, ₽</th><th style="min-width:180px;">Исполнение</th></tr></thead>
                    <tbody>${rowData.map(rowHtml).join('')}
                    <tr class="pl-grand-row"><td><strong>ИТОГО</strong></td>
                        <td class="number-cell"><strong>${this.fmtMoney(grand.total)}</strong></td>
                        <td class="number-cell"><strong>${this.fmtMoney(grand.limit)}</strong></td>
                        <td class="number-cell"><strong>${this.fmtMoney(Math.max(grand.limit - grand.total, 0))}</strong></td>
                        <td class="pl-bar-cell"><div class="limit-bar"><div class="limit-bar-fill ${grandCls}" style="width:${this.barWidth(grandPct)}%;"></div></div><span class="pl-meta">${isFinite(grandPct) ? this.fmtPct(grandPct) + '%' : '∞'}</span></td></tr>
                    <tr class="pl-global-row-tr"><td colspan="4"><strong>Глобальный лимит (${this.periodLabel(1)} + ${this.periodLabel(2)}, первые 15 дней):</strong><br>
                        <span class="pl-meta">${this.periodLabel(1)}: ${this.fmtMoney(totalP1)} · ${this.periodLabel(2)}: ${this.fmtMoney(totalP2)} · всего ${this.fmtMoney(sum12)} из ${this.fmtMoney(limit12)}</span></td>
                        <td class="pl-bar-cell"><div class="limit-bar"><div class="limit-bar-fill ${cls12}" style="width:${this.barWidth(pct12)}%;"></div></div>
                        <span class="pl-meta ${cls12 === 'limit-ok' ? '' : 'pl-pct-' + cls12}">${isFinite(pct12) ? this.fmtPct(pct12) + '%' : '∞'} (${sum12 > limit12 ? 'превышение' : 'в пределах'})</span></td></tr>
                    </tbody>
                </table></div>
            </div>`;
    }

    previewRowsHtml() {
        if (!this.previewRows.length) return '';
        const rows = this.filteredPreviewRows();
        const filter = this.previewFilter;
        const hasFilter = this.hasPreviewFilter();
        const filterLabel = [filter.category, filter.detail].filter(Boolean).join(' / ');

        let rowsHtml = rows.map((row, idx) => `
            <tr class="${this.rowCategory(row) ? '' : 'row-needs-review'}">
                <td>${idx + 1}</td>
                <td class="number-cell">${this.fmtMoney(row.amount)}</td>
                <td>${this.escapeHtml(row.recipient)}</td>
                <td>${this.escapeHtml(row.articleDds)}</td>
                <td><strong>${this.escapeHtml(row.category)}</strong></td>
                <td>${this.escapeHtml(row.detailLimit)}</td>
                <td>${this.escapeHtml(row.requestNo)}</td>
                <td>${this.escapeHtml(row.requestDate)}</td>
                <td>${this.escapeHtml(row.status)}</td>
                <td>${this.escapeHtml(row.paymentDate)}</td>
                <td>${this.escapeHtml(row.subdivision)}</td>
                <td title="${this.escapeHtml(row.purpose)}">${this.escapeHtml(row.purpose).slice(0, 70)}</td>
            </tr>`).join('');

        if (!rows.length) {
            rowsHtml = `<tr><td colspan="12" class="pl-empty-filter">Нет заявок по фильтру «${this.escapeHtml(filterLabel)}»</td></tr>`;
        }

        return `
            <div class="pl-block pl-block-flat" id="plPreviewListBlock">
                <div class="pl-block-title"><i class="fas fa-list-alt"></i> Список загруженных заявок (превью)
                    <small>${this.fileName ? 'файл: ' + this.escapeHtml(this.fileName) : ''} · загрузка в ${this.periodLabel(this.currentPeriod)}</small>
                    ${hasFilter ? `<span class="pl-filter-chip">Фильтр: ${this.escapeHtml(filterLabel)} (${rows.length} из ${this.previewRows.length})
                        <button type="button" class="pl-filter-reset" data-pl-filter-reset="1" title="Сбросить фильтр">✕</button></span>` : ''}
                </div>
                <div class="data-table-wrapper pl-table-wrap">
                <table class="data-table pl-requests-table">
                    <thead><tr><th>№</th><th>Сумма, ₽</th><th>Получатель</th><th>Статья движения ДС</th><th>Тип оплаты</th><th>Детальный лимит</th><th>Заявка</th><th>Дата заявки</th><th>Статус</th><th>Дата оплаты</th><th>Подразделение</th><th>Назначение платежа</th></tr></thead>
                    <tbody>${rowsHtml}</tbody>
                </table></div>
            </div>`;
    }

    renderControlArea() {
        const area = document.getElementById('plControlArea');
        if (!area) return;
        if (!this.confirmedLoaded) {
            area.innerHTML = '<div class="empty-state"><div class="empty-state-icon"><i class="fas fa-spinner fa-spin"></i></div>' +
                '<h3 class="empty-state-title">Загрузка данных…</h3></div>';
            return;
        }
        const parts = [this.controlTableHtml(), this.summaryLimitsHtml(), this.detailTablesHtml(), this.previewRowsHtml()];
        area.innerHTML = parts.filter(Boolean).join('');
    }

    // ===== РЕДАКТОР НАСТРОЕК ЛИМИТОВ =====
    getDefaultConfig() {
        return JSON.parse(JSON.stringify(PaymentLimitsManager.DEFAULT_CONFIG));
    }

    parseEditorNumber(value) {
        if (value === null || value === undefined) return 0;
        const cleaned = String(value).replace(/[\s\u00A0]/g, '').replace(',', '.');
        const num = parseFloat(cleaned);
        if (isNaN(num) || num < 0) return 0;
        return num;
    }

    renderEditorHtml() {
        const container = document.getElementById('plLimitsEditor');
        if (!container) return;
        const cfg = this.getConfig();
        const periods = PaymentLimitsManager.PERIODS;

        let html = `<p class="pl-editor-hint">Изменяйте названия лимитов и суммы по периодам (в рублях). Изменения применятся после нажатия «Сохранить».</p>`;

        // Категории
        html += `<div class="pl-editor-section"><h4><i class="fas fa-layer-group"></i> Лимиты по категориям («Тип оплаты»)</h4>`;
        html += `<div class="data-table-wrapper"><table class="data-table pl-editor-table"><thead><tr><th>Название лимита</th>${periods.map(p => `<th>${this.periodLabel(p)}</th>`).join('')}</tr></thead><tbody>`;
        cfg.categories.forEach(cat => {
            html += `<tr><td><input type="text" class="pl-editor-input" data-kind="cat-name" data-id="${cat.id}" value="${this.escapeHtml(cat.name)}"></td>`;
            periods.forEach(p => {
                const val = (cat.limits && cat.limits[String(p)]) || 0;
                html += `<td><input type="number" min="0" step="0.01" class="pl-editor-input pl-editor-number" data-kind="cat-limit" data-id="${cat.id}" data-period="${p}" value="${val}"></td>`;
            });
            html += `</tr>`;
        });
        html += `</tbody></table></div></div>`;

        // Детальные лимиты
        const catsWithDetails = cfg.categories.filter(c => (c.details || []).length);
        catsWithDetails.forEach(cat => {
            html += `<div class="pl-editor-section"><h4><i class="fas fa-sliders-h"></i> Детальные лимиты: ${this.escapeHtml(cat.name)}</h4>`;
            html += `<div class="data-table-wrapper"><table class="data-table pl-editor-table"><thead><tr><th>Подстатья</th>${periods.map(p => `<th>${this.periodLabel(p)}</th>`).join('')}</tr></thead><tbody>`;
            (cat.details || []).forEach(det => {
                html += `<tr><td><input type="text" class="pl-editor-input" data-kind="det-name" data-id="${det.id}" value="${this.escapeHtml(det.name)}"></td>`;
                periods.forEach(p => {
                    const val = (det.limits && det.limits[String(p)]) || 0;
                    html += `<td><input type="number" min="0" step="0.01" class="pl-editor-input pl-editor-number" data-kind="det-limit" data-id="${det.id}" data-period="${p}" value="${val}"></td>`;
                });
                html += `</tr>`;
            });
            html += `</tbody></table></div></div>`;
        });

        // Глобальный лимит
        html += `<div class="pl-editor-section"><h4><i class="fas fa-globe"></i> Глобальный лимит</h4>
            <div class="pl-editor-global">
                <label>${this.periodLabel(1)} + ${this.periodLabel(2)} (первые 15 дней), руб.</label>
                <input type="number" min="0" step="0.01" class="pl-editor-input pl-editor-number" data-kind="global" value="${Number(cfg.globalLimitP12) || 0}">
            </div></div>`;

        container.innerHTML = html;
    }

    collectEditorConfig() {
        const cfg = JSON.parse(JSON.stringify(this.getConfig()));
        const root = document.getElementById('plLimitsEditor');
        if (!root) return cfg;

        cfg.categories.forEach(cat => {
            const nameInput = root.querySelector(`input[data-kind="cat-name"][data-id="${cat.id}"]`);
            if (nameInput && nameInput.value.trim()) cat.name = nameInput.value.trim();
            cat.limits = cat.limits || {};
            PaymentLimitsManager.PERIODS.forEach(p => {
                const inp = root.querySelector(`input[data-kind="cat-limit"][data-id="${cat.id}"][data-period="${p}"]`);
                if (inp) cat.limits[String(p)] = this.parseEditorNumber(inp.value);
            });
            (cat.details || []).forEach(det => {
                const dn = root.querySelector(`input[data-kind="det-name"][data-id="${det.id}"]`);
                if (dn && dn.value.trim()) det.name = dn.value.trim();
                det.limits = det.limits || {};
                PaymentLimitsManager.PERIODS.forEach(p => {
                    const inp = root.querySelector(`input[data-kind="det-limit"][data-id="${det.id}"][data-period="${p}"]`);
                    if (inp) det.limits[String(p)] = this.parseEditorNumber(inp.value);
                });
            });
        });

        const globalInput = root.querySelector('input[data-kind="global"]');
        if (globalInput) cfg.globalLimitP12 = this.parseEditorNumber(globalInput.value);

        return cfg;
    }

    resetEditorToDefaults() {
        this.config = this.getDefaultConfig();
        this.renderEditorHtml();
    }
}

// Класс PaymentLimitsManager объявлен глобально (как и остальные *-manager.js),
// app.js подключается после payment-limits.js в index.html.
