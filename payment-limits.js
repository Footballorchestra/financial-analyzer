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
        this.sourceFile = null;      // исходный загруженный файл (нужен для экспорта «как в источнике»)
        this.sourceSheetName = null; // имя листа исходного файла
        this.registry = null;        // сохранённый в БД реестр месяца (черновик/финал + файл)
        this.registries = [];        // все сохранённые реестры месяца
        this.localDirty = false;     // были ли изменения отметок после загрузки/сохранения
        // true, когда превью — это копия сохранённого в БД варианта (тогда её нельзя складывать с БД)
        this.previewIsSavedCopy = false;
        this.confirmedLoaded = false;
        this.config = null;          // конфигурация лимитов (с сервера или defaults)
        this.configLoaded = false;
        this.onlyCurrentPeriod = this.loadOnlyCurrentPeriodSetting(); // режим «только текущий период»
        this.activeTab = this.loadActiveTabSetting();  // активный блок на экране (вкладка)
        this.explanationGroup = this.loadExplanationGroupSetting(); // категория в блоке «Пояснения»
        this.library = new Map();    // справочник контрагентов: наименование → пояснение
        this.libraryLoaded = false;
        this.previewFilter = { category: null, detail: null, group: null, subdivision: null, recipient: null }; // фильтр списка заявок
    }

    static PERIODS = [1, 2, 3, 4];
    static PERIOD_DAYS = { 1: '1–7', 2: '8–15', 3: '16–22', 4: '23–31' };

    // Блоки, которые вызываются кнопками горизонтального меню.
    // Список заявок (блок 5) в меню не входит — он показывается всегда.
    static TABS = ['control', 'summary', 'detail', 'groups', 'explanations'];

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

    // Категории блока «Разбивка на категории» — значения колонки «Группа» (правый столбец файла).
    // Порядок задаётся здесь и не зависит от порядка строк в файле.
    static GROUP_CATEGORIES = [
        'Филиалы',
        'УК',
        'Питание, проживание, связь',
        'Оплаты в рамках 100 т.р.',
        'Непереносимые оплаты',
        'Графики',
        'Согласовано в оплату'
    ];

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
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
            if (rows.length < 2) {
                return { success: false, message: 'В файле нет данных' };
            }
            // Нужно для точного сопоставления строк файла и списка заявок (экспорт «как в источнике»)
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');

            const headers = rows[0].map(h => String(h == null ? '' : h).trim());
            const columns = this.detectColumns(headers);
            if (columns.category === undefined || columns.amount === undefined) {
                return { success: false, message: 'Не найдены обязательные колонки: "Тип оплаты" и "Сумма"' };
            }
            // Колонка «Группа» — самый правый столбец файла: если заголовок не найден, берём последний
            if (columns.group === undefined && headers.length > 0) {
                columns.group = headers.length - 1;
            }

            const preview = [];
            rows.slice(1).forEach((row, rowIndex) => {
                if (!row.some(cell => cell !== null && cell !== '')) return;
                const amount = this.parseAmount(row[columns.amount]);
                if (amount === 0) return;
                preview.push({
                    excelRow: range.s.r + rowIndex + 2, // номер строки в исходном файле (1-based)
                    selected: true,                     // сразу после загрузки все заявки отмечены
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
                    paymentType: columns.paymentType !== undefined ? this.cleanString(row[columns.paymentType]) : '',
                    group: columns.group !== undefined ? this.cleanString(row[columns.group]) : ''
                });
            });

            if (!preview.some(r => r.category)) {
                return { success: false, message: 'В файле не заполнена колонка "Тип оплаты"' };
            }

            const now = new Date();
            this.sourceFile = file;               // исходный файл — для экспорта без пересборки
            this.sourceSheetName = sheetName;
            // Ручная загрузка — это новая партия периода (складывается с сохранёнными)
            this.markAsNewUpload();
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

    // Загрузка вручную — это НОВАЯ партия (загрузка) периода: при сохранении
    // она добавится к ранее сохранённым загрузкам, а не заменит их.
    markAsNewUpload() {
        this.registry = null;              // не продолжаем ранее сохранённую загрузку
        this.previewIsSavedCopy = false;   // превью складывается с сохранёнными загрузками
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
            group: find(['группа']),
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

    // Нормализация значения колонки «Группа» по фиксированному списку категорий блока «Разбивка».
    // Пустая строка означает «не сопоставлено» — такие заявки показываются строкой «Прочее».
    normalizeGroup(raw) {
        const s = this.cleanString(raw).toLowerCase().replace(/ё/g, 'е');
        if (!s) return '';
        const norm = (name) => name.toLowerCase().replace(/ё/g, 'е');
        // 1) точное совпадение
        for (const name of PaymentLimitsManager.GROUP_CATEGORIES) {
            if (norm(name) === s) return name;
        }
        // 2) вхождение подстроки
        for (const name of PaymentLimitsManager.GROUP_CATEGORIES) {
            const n = norm(name);
            if (s.includes(n) || n.includes(s)) return name;
        }
        return '';
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

    // Сохранённые строки месяца по категориям и периодам.
    // excludeRegistryId — исключить партию (загрузку), данные которой уже показаны в превью.
    confirmedByPeriod(excludeRegistryId = null) {
        const result = {};
        const exclude = excludeRegistryId != null ? Number(excludeRegistryId) : null;
        this.confirmedRows.forEach(r => {
            if (exclude !== null && Number(r.registry_id) === exclude) return;
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
    // В итогах блоков («БД») учитывается финал, а если финала нет — сохранённый черновик
    async loadConfirmed(month) {
        try {
            const url = month
                ? `/api/payment-limits/confirmed?month=${month}&stage=effective`
                : '/api/payment-limits/confirmed?stage=effective';
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

    // Подтверждение (совместимость): сохраняет финальный вариант вместе с файлом реестра
    async confirmSave() {
        const result = await this.saveStage('final');
        if (!result.success) return result;
        return {
            success: true,
            count: result.count,
            total_amount: result.total,
            message: result.message || `Сохранено ${result.count} платежей`
        };
    }

    // ===== ЭКСПОРТ В EXCEL: исходный файл без снятых заявок =====
    // Файл не пересобирается: сервер копирует исходный .xlsx и удаляет из него строки
    // снятых заявок, поэтому формат, столбцы, оформление и «Дата заявки» остаются исходными.
    async exportToExcel() {
        const hasRegistryFile = !!(this.registry && this.registry.has_file);
        if (!this.previewRows.length && !hasRegistryFile) {
            return { success: false, message: 'Нет данных для экспорта' };
        }
        if (!this.sourceFile && !hasRegistryFile) {
            return { success: false, message: 'Не найден исходный файл — загрузите реестр заново' };
        }

        const removedBySelection = this.previewRows
            .filter(row => !this.isRowSelected(row))
            .map(row => row.excelRow);
        const exportStage = this.savedStage();   // 'final' | 'draft' | null
        // Если открыт сохранённый вариант и отметки не менялись — выгружаем сохранённый набор стадии,
        // иначе — то, что отмечено на экране.
        const useCurrentSelection = !!this.previewRows.length && (this.localDirty || !exportStage);

        const payload = {
            sheetName: this.sourceSheetName || '',
            fileName: this.fileName || ''
        };
        if (this.registry) payload.registryId = this.registry.id;
        if (exportStage) payload.stage = exportStage;
        if (useCurrentSelection) payload.removeRows = removedBySelection;

        try {
            const formData = new FormData();
            if (this.sourceFile) {
                formData.append('file', this.sourceFile, this.sourceFile.name || 'registry.xlsx');
            }
            formData.append('data', JSON.stringify(payload));

            const response = await fetch('/api/payment-limits/export', { method: 'POST', body: formData });
            if (!response.ok) {
                let message = `Ошибка экспорта (HTTP ${response.status})`;
                try {
                    const err = await response.json();
                    message = err.error || err.message || message;
                } catch (e) { /* ответ не JSON */ }
                return { success: false, message };
            }

            const blob = await response.blob();
            const fileName = this.fileNameFromResponse(response)
                || `Заявки_лимиты_${this.todayRu()}.xlsx`;

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);

            const removed = Number(response.headers.get('X-Export-Removed') || removeRows.length) || 0;
            let warning = response.headers.get('X-Export-Warning') || '';
            if (warning) {
                try { warning = decodeURIComponent(warning); } catch (e) { /* ignore */ }
            }

            // Уведомление «можно выгружать» снимается после выгрузки финального файла
            const exportedFinalAt = response.headers.get('X-Export-Final-At');
            if (exportedFinalAt && this.registry) {
                this.registry.final_exported_at = exportedFinalAt;
            } else if (this.registry) {
                await this.loadRegistryState();
            }
            return {
                success: true,
                removed,
                warning,
                message: `Файл «${fileName}»: ${this.selectedCount()} заявок, снято ${removed}` + (warning ? `. ${warning}` : '')
            };
        } catch (error) {
            console.error('Ошибка экспорта реестра лимитов:', error);
            return { success: false, message: 'Ошибка экспорта: ' + error.message };
        }
    }

    // Имя файла из заголовка Content-Disposition (с учётом filename*=UTF-8'')
    fileNameFromResponse(response) {
        const header = response.headers.get('Content-Disposition') || '';
        const utf8 = header.match(/filename\*=UTF-8''([^;]+)/i);
        if (utf8) {
            try { return decodeURIComponent(utf8[1]); } catch (e) { /* ignore */ }
        }
        const simple = header.match(/filename="?([^";]+)"?/i);
        return simple ? simple[1] : null;
    }

    // Сегодняшняя дата в формате ДД.ММ.ГГГГ
    todayRu() {
        const d = new Date();
        const day = String(d.getDate()).padStart(2, '0');
        const month = String(d.getMonth() + 1).padStart(2, '0');
        return `${day}.${month}.${d.getFullYear()}`;
    }

    // Готов ли финальный вариант к выгрузке: сохранён финал и ещё не выгружен
    finalExportPending() {
        const reg = this.registry;
        if (!reg || !reg.final_saved_at) return false;
        if (!reg.final_exported_at) return true;
        return new Date(reg.final_exported_at).getTime() < new Date(reg.final_saved_at).getTime();
    }

    // ===== СОХРАНЕНИЕ СТАДИЙ (ЧЕРНОВИК / ФИНАЛ) И СОХРАНЁННЫЙ РЕЕСТР =====
    // Какая стадия сохранена: сначала финал, иначе черновик
    savedStage() {
        if (!this.registry) return null;
        if (this.registry.final_saved_at) return 'final';
        if (this.registry.draft_saved_at) return 'draft';
        return null;
    }

    // Применяет сохранённый набор снятых строк к загруженному превью
    applyRemovedRows(removedRows) {
        const removed = new Set((removedRows || []).map(v => String(v)));
        this.previewRows.forEach(row => {
            row.selected = !removed.has(String(row.excelRow));
        });
        this.localDirty = false;
        return this.selectedCount();
    }

    async saveStage(stage) {
        if (!this.previewRows.length) {
            return { success: false, message: 'Нет данных для сохранения' };
        }
        const hasRegistryFile = !!(this.registry && this.registry.has_file);
        if (!this.sourceFile && !hasRegistryFile) {
            return { success: false, message: 'Нужен исходный файл реестра: загрузите файл заново' };
        }

        const selected = this.selectedPreviewRows();
        const removedRows = this.previewRows
            .filter(row => !this.isRowSelected(row))
            .map(row => row.excelRow);

        try {
            const formData = new FormData();
            if (this.sourceFile) {
                formData.append('file', this.sourceFile, this.sourceFile.name || 'registry.xlsx');
            }
            formData.append('data', JSON.stringify({
                stage,
                period: this.getLoadPeriod(),
                rows: selected.map(row => ({
                    article: row.category || row.articleDds || '',
                    detail_limit: row.detailLimit || '',
                    amount: row.amount,
                    recipient: row.recipient || '',
                    request_no: row.requestNo || '',
                    request_date: row.requestDate || null,
                    status: row.status || '',
                    organization: row.organization || '',
                    subdivision: row.subdivision || '',
                    priority: row.priority || '',
                    purpose: row.purpose || '',
                    payment_date: row.paymentDate || null,
                    applicant: row.applicant || '',
                    payment_type: row.paymentType || ''
                })),
                removedRows,
                sheetName: this.sourceSheetName || '',
                fileName: this.fileName || '',
                rowCount: this.previewRows.length,
                registryId: this.registry ? this.registry.id : null
            }));

            const response = await fetch('/api/payment-limits/save', { method: 'POST', body: formData });
            const result = await response.json();
            if (!response.ok || !result.success) {
                return { success: false, message: result.error || result.message || 'Ошибка сохранения' };
            }

            if (!this.registry) this.registry = {};
            this.registry.id = result.registry_id || this.registry.id;
            this.registry.has_file = true;
            this.registry.period = result.period;
            this.registry.limit_month = result.month;
            this.registry.row_count = this.previewRows.length;
            if (stage === 'final') {
                this.registry.final_saved_at = result.saved_at;
                this.registry.final_removed_rows = removedRows;
            } else {
                this.registry.draft_saved_at = result.saved_at;
                this.registry.draft_removed_rows = removedRows;
            }
            this.localDirty = false;
            // Локальный файл больше не нужен — он сохранён в БД
            this.sourceFile = null;
            // После сохранения превью совпадает с сохранённым вариантом — двойного счёта быть не должно
            this.previewIsSavedCopy = true;

            return {
                success: true,
                stage: result.stage,
                count: result.count,
                removed: result.removed_count,
                total: result.total_amount,
                saved_at: result.saved_at,
                message: result.message
            };
        } catch (error) {
            console.error('Ошибка сохранения стадии лимитов:', error);
            return { success: false, message: 'Ошибка сохранения: ' + error.message };
        }
    }

    // Состояние сохранённого реестра месяца (когда сохранены черновик и финал)
    async loadRegistryState() {
        try {
            const month = this.currentMonth || this.monthPrefix(new Date());
            const response = await fetch(`/api/payment-limits/registry?month=${month}`);
            const result = await response.json();
            if (!response.ok || !result.success) {
                return { success: false, message: result.error || 'Не удалось получить состояние реестра' };
            }
            this.registries = result.registries || [];
            // Последняя сохранённая загрузка (партия) периода — её и восстанавливаем
            const forPeriod = this.registries.filter(r => Number(r.period) === Number(this.getLoadPeriod()));
            this.registry = forPeriod.length
                ? forPeriod[forPeriod.length - 1]
                : (this.registries[this.registries.length - 1] || null);
            return { success: true, registry: this.registry, registries: this.registries };
        } catch (error) {
            console.warn('Не удалось получить состояние реестра лимитов:', error);
            return { success: false, message: error.message };
        }
    }

    // Автозагрузка сохранённого варианта месяца (чтобы данные появились сразу на любом компьютере)
    async loadRegistryPreview() {
        if (this.previewRows.length) return { success: false, message: 'Превью уже загружено' };
        if (!this.registry || !this.registry.has_file) {
            return { success: false, message: 'Сохранённый реестр не найден' };
        }
        try {
            const response = await fetch(`/api/payment-limits/registry/file?registry_id=${this.registry.id}`);
            if (!response.ok) {
                return { success: false, message: `Не удалось получить файл реестра (HTTP ${response.status})` };
            }
            const blob = await response.blob();
            const fileName = this.fileNameFromResponse(response) || this.registry.file_name || 'registry.xlsx';
            const file = this.blobToFile(blob, fileName);

            // loadExcelFile считает любую загрузку новой партией, поэтому сохраняем
            // и восстанавливаем выбранный сохранённый реестр после чтения файла.
            const reg = this.registry;
            const loaded = await this.loadExcelFile(file);
            if (!loaded.success) return loaded;
            this.registry = reg;

            // Файл уже сохранён в БД — повторно загружать его не нужно
            this.sourceFile = null;
            this.fileName = fileName;
            if (this.registry.period) this.currentPeriod = Number(this.registry.period);
            if (this.registry.sheet_name) this.sourceSheetName = this.registry.sheet_name;

            const stage = this.savedStage();
            const removedRows = stage === 'final'
                ? this.registry.final_removed_rows
                : this.registry.draft_removed_rows;
            this.applyRemovedRows(removedRows);
            // Превью — копия сохранённого варианта: в итогах не складываем её с БД
            this.previewIsSavedCopy = true;

            return {
                success: true,
                stage,
                removed: (removedRows || []).length,
                count: loaded.count,
                total: this.selectedTotal(),
                message: `Загружен ${stage === 'final' ? 'финальный вариант' : 'черновик'}: `
                    + `${loaded.count} заявок, снято ${(removedRows || []).length}`
            };
        } catch (error) {
            console.error('Ошибка автозагрузки сохранённого реестра:', error);
            return { success: false, message: 'Ошибка автозагрузки реестра: ' + error.message };
        }
    }

    blobToFile(blob, fileName) {
        try {
            return new File([blob], fileName, {
                type: blob.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });
        } catch (e) {
            // Среда без конструктора File — FileReader умеет читать и Blob
            blob.name = fileName;
            return blob;
        }
    }

    // Текст состояния сохранённого варианта для шапки страницы
    registryStatusText() {
        if (!this.registry) return 'нет сохранённых данных';
        const fmt = (iso) => {
            if (!iso) return '';
            const d = new Date(iso);
            if (isNaN(d.getTime())) return String(iso);
            return d.toLocaleString('ru-RU', {
                day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit'
            });
        };
        const parts = [];
        if (this.registry.draft_saved_at) parts.push(`черновик сохранён ${fmt(this.registry.draft_saved_at)}`);
        if (this.registry.final_saved_at) parts.push(`финал сохранён ${fmt(this.registry.final_saved_at)}`);
        if (!parts.length) return 'нет сохранённых данных';
        const dirty = this.localDirty ? ' · есть несохранённые правки' : '';
        const count = this.periodRegistriesCount();
        const regs = count > 1 ? ` · загрузок в периоде: ${count}` : '';
        return `${parts.join(' · ')} · ${this.periodLabel(this.registry.period)}${regs}${dirty}`;
    }

    // Обнуление ранее загруженных данных месяца (черновик, финал, платежи и исходный файл)
    async resetSavedData() {
        try {
            // Обнуляем весь месяц целиком (все периоды и обе стадии)
            const payload = { month: this.currentMonth || this.monthPrefix(new Date()) };

            const response = await fetch('/api/payment-limits/reset', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const result = await response.json();
            if (!response.ok || !result.success) {
                return { success: false, message: result.error || result.message || 'Ошибка обнуления' };
            }

            // Полностью очищаем состояние страницы, чтобы начать с чистого листа
            this.previewRows = [];
            this.previewFilter = { category: null, detail: null, group: null, subdivision: null, recipient: null };
            this.registry = null;
            this.registries = [];
            this.confirmedRows = [];
            this.confirmedLoaded = false;
            this.previewIsSavedCopy = false;
            this.localDirty = false;
            this.fileName = null;
            this.sourceFile = null;
            this.sourceSheetName = null;

            return {
                success: true,
                payments: result.payments_deleted,
                registries: result.registries_deleted,
                message: result.message
            };
        } catch (error) {
            console.error('Ошибка обнуления данных лимитов:', error);
            return { success: false, message: 'Ошибка обнуления: ' + error.message };
        }
    }

    cancelPreview() {
        this.previewRows = [];
        this.fileName = null;
        this.sourceFile = null;
        this.sourceSheetName = null;
        this.previewIsSavedCopy = false;
        this.localDirty = false;
        this.previewFilter = { category: null, detail: null, group: null, subdivision: null, recipient: null };
    }

    // ===== ВЫБОР ЗАЯВОК ЧЕКБОКСАМИ =====
    // Все итоги во всех блоках считаются только по отмеченным заявкам.
    selectedPreviewRows() {
        return this.previewRows.filter(row => row.selected !== false);
    }

    selectedCount() {
        return this.selectedPreviewRows().length;
    }

    selectedTotal() {
        return this.selectedPreviewRows().reduce((sum, row) => sum + (Number(row.amount) || 0), 0);
    }

    setRowSelected(excelRow, selected) {
        const row = this.previewRows.find(r => String(r.excelRow) === String(excelRow));
        if (!row) return false;
        row.selected = !!selected;
        this.localDirty = true;
        return true;
    }

    setRowsSelected(rows, selected) {
        (rows || []).forEach(row => { row.selected = !!selected; });
        if ((rows || []).length) this.localDirty = true;
        return (rows || []).length;
    }

    setAllSelected(selected) {
        this.previewRows.forEach(row => { row.selected = !!selected; });
        this.localDirty = true;
    }

    // Снятая чекбоксом заявка не попадает ни в итоги, ни в экспорт
    isRowSelected(row) {
        return row.selected !== false;
    }

    // ===== ОБЩИЕ (ГРУППОВЫЕ) ЧЕКБОКСЫ ПО КАТЕГОРИЯМ =====
    rowsByCategory(categoryName) {
        return this.previewRows.filter(row => this.rowCategory(row) === categoryName);
    }

    rowsByDetail(categoryName, detailName) {
        return this.previewRows.filter(row =>
            this.rowCategory(row) === categoryName
            && this.normalizeDetailLimit(row.detailLimit) === detailName);
    }

    rowsByGroup(groupName) {
        return this.previewRows.filter(row => this.normalizeGroup(row.group) === groupName);
    }

    // Состояние группового чекбокса: отмечен / не отмечен / часть (indeterminate)
    toggleState(rows) {
        const list = rows || [];
        const selected = list.filter(row => this.isRowSelected(row)).length;
        return {
            total: list.length,
            selected,
            checked: list.length > 0 && selected === list.length,
            partial: selected > 0 && selected < list.length
        };
    }

    // HTML группового чекбокса. attrs — дополнительные data-атрибуты (например, data-pl-cat-toggle).
    groupToggleHtml(rows, attrs, title) {
        const state = this.toggleState(rows);
        if (!state.total) return '';
        return `<input type="checkbox" class="pl-toggle" ${attrs}`
            + `${state.checked ? ' checked' : ''}`
            + `${state.partial ? ' data-pl-partial="1"' : ''}`
            + ` title="${this.escapeHtml(title)}">`;
    }

    // ===== ФИЛЬТР СПИСКА ЗАЯВОК (категория / подстатья / группа / подразделение / контрагент) =====
    setPreviewFilter(category, detail, group, subdivision, recipient) {
        this.previewFilter = {
            category: category || null,
            detail: detail || null,
            group: group || null,
            subdivision: subdivision || null,
            recipient: recipient || null
        };
    }

    hasPreviewFilter() {
        const f = this.previewFilter;
        return !!(f.category || f.detail || f.group || f.subdivision || f.recipient);
    }

    filteredPreviewRows() {
        const f = this.previewFilter;
        return this.previewRows.filter(row => {
            if (f.category && this.rowCategory(row) !== f.category) return false;
            if (f.detail && this.normalizeDetailLimit(row.detailLimit) !== f.detail) return false;
            if (f.group && this.normalizeGroup(row.group) !== f.group) return false;
            if (f.subdivision && (row.subdivision || '').trim() !== f.subdivision) return false;
            if (f.recipient && (row.recipient || '').trim() !== f.recipient) return false;
            return true;
        });
    }

    // Обновление только блока списка заявок (без полной перерисовки страницы)
    refreshPreviewList() {
        const el = document.getElementById('plPreviewListBlock');
        if (!el) return;
        el.outerHTML = this.previewRowsHtml();
        this.applyToggleStates();
    }

    // Проставляем «частичное» состояние групповых чекбоксов (indeterminate)
    applyToggleStates() {
        if (typeof document === 'undefined') return;
        const area = document.getElementById('plControlArea');
        if (!area) return;
        area.querySelectorAll('input.pl-toggle[data-pl-partial="1"]').forEach(input => {
            input.indeterminate = true;
        });
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
    // ===== ГОРИЗОНТАЛЬНОЕ МЕНЮ БЛОКОВ (ВКЛАДКИ) =====
    loadActiveTabSetting() {
        try {
            const tab = localStorage.getItem('paymentLimitsActiveTab');
            return PaymentLimitsManager.TABS.includes(tab) ? tab : 'control';
        } catch (e) {
            return 'control';
        }
    }

    setActiveTab(tab) {
        const value = PaymentLimitsManager.TABS.includes(tab) ? tab : 'control';
        this.activeTab = value;
        try {
            localStorage.setItem('paymentLimitsActiveTab', value);
        } catch (e) { /* localStorage недоступен — не критично */ }
        return value;
    }

    // HTML активного блока (список заявок показывается отдельно и всегда)
    activeBlockHtml() {
        switch (this.activeTab) {
            case 'summary': return this.summaryLimitsHtml();
            case 'detail': return this.detailTablesHtml();
            case 'groups': return this.groupBreakdownHtml();
            case 'explanations': return this.explanationsHtml();
            default: return this.controlTableHtml();
        }
    }

    // Синхронизация активной вкладки меню (активной всегда ровно одна)
    applyTabState() {
        if (typeof document === 'undefined') return;
        const tabs = document.getElementById('plTabs');
        if (!tabs) return;
        // Основной механизм — радио-кнопки: браузер сам гарантирует единственную активную
        tabs.querySelectorAll('input[name="plTabs"]').forEach(radio => {
            radio.checked = radio.value === this.activeTab;
        });
        // Совместимость со старой разметкой на кнопках
        tabs.querySelectorAll('[data-pl-tab]').forEach(btn => {
            btn.classList.toggle('active', btn.getAttribute('data-pl-tab') === this.activeTab);
        });
    }

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

    // Режим «открыт сохранённый вариант месяца»: превью — копия того же набора, что уже лежит в БД.
    // В этом режиме итоги считаем один раз (по превью), иначе получается двойной счёт.
    isSavedCopyMode() {
        return this.previewIsSavedCopy && this.previewRows.length > 0;
    }

    // Текущая сохранённая загрузка (партия), данные которой показаны в превью
    savedRegistryId() {
        return this.registry ? Number(this.registry.id) : null;
    }

    // Сколько загрузок (партий) сохранено в текущем периоде
    periodRegistriesCount() {
        return (this.registries || [])
            .filter(r => Number(r.period) === Number(this.getLoadPeriod()))
            .length;
    }

    // Периоды из БД: в режиме открытого сохранённого варианта его партия исключается
    // (её данные показаны в превью), а остальные сохранённые загрузки периода суммируются
    dbByPeriod() {
        const exclude = this.isSavedCopyMode() ? this.savedRegistryId() : null;
        return this.confirmedByPeriod(exclude);
    }

    // Строки из БД для детальных блоков (без партии, открытой в превью)
    dbRows() {
        if (!this.isSavedCopyMode()) return this.confirmedRows;
        const exclude = this.savedRegistryId();
        if (exclude === null) return this.confirmedRows;
        return this.confirmedRows.filter(r => Number(r.registry_id) !== exclude);
    }

    // Подпись режима для заголовков блоков
    modeNote() {
        if (!this.isSavedCopyMode()) return '';
        const stage = this.savedStage() === 'final' ? 'финальный вариант' : 'черновик';
        const count = this.periodRegistriesCount();
        return ` · открыта сохранённая загрузка (${stage})`
            + (count > 1 ? `; в периоде сохранено загрузок: ${count}` : '');
    }

    // Подпись под суммой в ячейке
    cellMetaText(confirmedRub, previewRub, pctText) {
        if (this.isSavedCopyMode()) {
            const note = this.localDirty ? ' · есть несохранённые правки' : '';
            return `другие загрузки ${this.fmtMoney(confirmedRub || 0)} · эта загрузка `
                + `${this.fmtMoney(previewRub || 0)}${note} · ${pctText}`;
        }
        return `сохранено ${this.fmtMoney(confirmedRub || 0)} (все загрузки) · `
            + `добавка ${this.fmtMoney(previewRub || 0)} · ${pctText}`;
    }
    cellHtml(confirmedRub, previewRub, limitRub) {
        const total = (confirmedRub || 0) + (previewRub || 0);
        let pct;
        if (limitRub > 0) { pct = (total / limitRub) * 100; } else { pct = total > 0 ? Infinity : 0; }
        const cls = this.barClass(pct);
        const pctText = isFinite(pct) ? `${this.fmtPct(pct)}%` : '∞';
        return `
            <div class="pl-cell">
                <div class="pl-amounts"><span class="pl-amount">${this.fmtMoney(total)}</span>
                    <span class="pl-of">/ ${limitRub > 0 ? this.fmtMoney(limitRub) : '—'}</span></div>
                <div class="limit-bar"><div class="limit-bar-fill ${cls}" style="width:${this.barWidth(pct)}%;"></div></div>
                <div class="pl-meta">${this.cellMetaText(confirmedRub, previewRub, pctText)}</div>
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
                <div class="pl-amounts"><span class="pl-amount">${this.fmtMoney(total)}</span>
                    <span class="pl-of">/ ${limitRub > 0 ? this.fmtMoney(limitRub) : '—'}</span></div>
                <div class="limit-bar"><div class="limit-bar-fill ${cls}" style="width:${this.barWidth(pct)}%;"></div></div>
                <div class="pl-meta">${pctText}</div>
            </div>`;
    }

    controlTableHtml() {
        // В режиме открытого сохранённого варианта БД не суммируем: превью — его же копия
        const conf = this.dbByPeriod();
        // Итоги считаем только по отмеченным заявкам
        const prev = this.aggregatePreview(this.selectedPreviewRows(), this.currentPeriod);
        const cur = this.currentPeriod;
        const headersHtml = PaymentLimitsManager.PERIODS.map(p => this.periodThHtml(p)).join('');

        let html = `
            <div class="pl-block pl-block-flat">
                <div class="pl-block-title"><i class="fas fa-table"></i> Контроль лимитов по периодам
                    <small>сверка по «Тип оплаты» из файла, суммы в рублях${this.modeNote()}</small></div>
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
                <td class="pl-article-name">${this.groupToggleHtml(this.rowsByCategory(cat), `data-pl-cat-toggle="${this.escapeHtml(cat)}"`, `Отметить/снять все заявки категории «${cat}» (${this.rowsByCategory(cat).length})`)}<button type="button" class="pl-cat-link" data-pl-category="${this.escapeHtml(cat)}" title="Показать заявки категории «${this.escapeHtml(cat)}» в списке ниже">${this.escapeHtml(cat)}</button>${known ? '' : '<br><small class="pl-warn-label">не сопоставлено с матрицей</small>'}</td>`;
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

        html += `<tr class="pl-grand-row"><td>ИТОГО по всем категориям</td>`;
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
            const conf = this.detailByPeriod(this.dbRows(), cat.name, null);
            // Итоги считаем только по отмеченным заявкам
            const prev = this.detailByPeriod(this.selectedPreviewRows(), cat.name, this.currentPeriod);
            const aksHasRows = this.selectedPreviewRows().some(r => this.rowCategory(r) === cat.name)
                || this.dbRows().some(r => this.normalizeCategory(r.article) === cat.name);
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
                rowsHtml += `<tr class="${known ? '' : 'pl-unmatched'}${activeSub ? ' pl-filter-row' : ''}"><td class="pl-article-name">${this.groupToggleHtml(this.rowsByDetail(cat.name, sub), `data-pl-detail-toggle="${this.escapeHtml(sub)}" data-pl-cat="${this.escapeHtml(cat.name)}"`, `Отметить/снять все заявки подстатьи «${sub}» (${this.rowsByDetail(cat.name, sub).length})`)}<button type="button" class="pl-cat-link" data-pl-category="${this.escapeHtml(cat.name)}" data-pl-detail="${this.escapeHtml(sub)}" title="Показать заявки подстатьи «${this.escapeHtml(sub)}» в списке ниже">${this.escapeHtml(sub)}</button></td>`;
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

            let totalRow = `<tr class="pl-grand-row"><td>ИТОГО по подстатьям</td>`;
            PaymentLimitsManager.PERIODS.forEach(p => {
                totalRow += `<td class="pl-cell-td">${this.cellHtmlTotal(grandByPeriod[p].total, grandByPeriod[p].limit)}</td>`;
            });
            totalRow += `<td class="number-cell pl-month-total">${this.fmtMoney(grandMonth)}</td></tr>`;

            blocks.push(`
                <div class="pl-block pl-block-flat">
                    <div class="pl-block-title"><i class="fas fa-sliders-h"></i>
                        Детальный контроль: ${this.escapeHtml(cat.name)}
                        <small>по подстатьям из колонки «Детальный лимит»${this.modeNote()}</small></div>
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
        // В режиме открытого сохранённого варианта БД не суммируем: превью — его же копия
        const conf = this.dbByPeriod();
        // Итоги считаем только по отмеченным заявкам
        const prev = this.aggregatePreview(this.selectedPreviewRows(), this.currentPeriod);
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
                        : `подтверждено за периоды 1–${cur} + превью текущей загрузки`}${this.modeNote()}</small></div>
                <div class="data-table-wrapper pl-table-wrap">
                <table class="data-table pl-table-narrow">
                    <thead><tr><th>Тип оплаты</th><th>Итого, ₽</th><th>Доступный лимит, ₽</th><th>Остаток, ₽</th><th style="min-width:180px;">Исполнение</th></tr></thead>
                    <tbody>${rowData.map(rowHtml).join('')}
                    <tr class="pl-grand-row"><td>ИТОГО</td>
                        <td class="number-cell">${this.fmtMoney(grand.total)}</td>
                        <td class="number-cell">${this.fmtMoney(grand.limit)}</td>
                        <td class="number-cell">${this.fmtMoney(Math.max(grand.limit - grand.total, 0))}</td>
                        <td class="pl-bar-cell"><div class="limit-bar"><div class="limit-bar-fill ${grandCls}" style="width:${this.barWidth(grandPct)}%;"></div></div><span class="pl-meta">${isFinite(grandPct) ? this.fmtPct(grandPct) + '%' : '∞'}</span></td></tr>
                    <tr class="pl-global-row-tr"><td colspan="4"><span class="pl-global-label">Глобальный лимит (${this.periodLabel(1)} + ${this.periodLabel(2)}, первые 15 дней):</span><br>
                        <span class="pl-meta">${this.periodLabel(1)}: ${this.fmtMoney(totalP1)} · ${this.periodLabel(2)}: ${this.fmtMoney(totalP2)} · всего ${this.fmtMoney(sum12)} из ${this.fmtMoney(limit12)}</span></td>
                        <td class="pl-bar-cell"><div class="limit-bar"><div class="limit-bar-fill ${cls12}" style="width:${this.barWidth(pct12)}%;"></div></div>
                        <span class="pl-meta ${cls12 === 'limit-ok' ? '' : 'pl-pct-' + cls12}">${isFinite(pct12) ? this.fmtPct(pct12) + '%' : '∞'} (${sum12 > limit12 ? 'превышение' : 'в пределах'})</span></td></tr>
                    </tbody>
                </table></div>
            </div>`;
    }

    // Блок «Разбивка на категории» — по значениям колонки «Группа» (правый столбец файла)
    groupBreakdownHtml() {
        if (!this.previewRows.length) return '';
        const selected = this.selectedPreviewRows();

        const byGroup = new Map(PaymentLimitsManager.GROUP_CATEGORIES.map(name => [name, { total: 0, count: 0 }]));
        const other = { total: 0, count: 0 };
        selected.forEach(row => {
            const name = this.normalizeGroup(row.group);
            const bucket = (name && byGroup.has(name)) ? byGroup.get(name) : other;
            bucket.total += Number(row.amount) || 0;
            bucket.count += 1;
        });

        const grand = { total: other.total, count: other.count };
        byGroup.forEach(v => { grand.total += v.total; grand.count += v.count; });
        const totalAll = grand.total;

        const rowHtml = (name, data, clickable) => {
            const share = totalAll > 0 ? (data.total / totalAll) * 100 : 0;
            const label = clickable
                ? `<button type="button" class="pl-cat-link" data-pl-group="${this.escapeHtml(name)}" title="Показать заявки группы «${this.escapeHtml(name)}» в списке ниже">${this.escapeHtml(name)}</button>`
                : `<span class="pl-warn-label">${this.escapeHtml(name)}</span>`;
            const active = clickable && this.previewFilter.group === name;
            const toggle = clickable
                ? this.groupToggleHtml(this.rowsByGroup(name), `data-pl-group-toggle="${this.escapeHtml(name)}"`,
                    `Отметить/снять все заявки группы «${name}» (${this.rowsByGroup(name).length})`)
                : '';
            return `<tr class="${clickable ? '' : 'pl-unmatched'}${active ? ' pl-filter-row' : ''}">
                <td class="pl-article-name">${toggle}${label}</td>
                <td class="number-cell">${this.fmtMoney(data.total)}</td>
                <td class="number-cell">${data.count}</td>
                <td class="pl-bar-cell"><div class="limit-bar"><div class="limit-bar-fill limit-ok" style="width:${this.barWidth(share)}%;"></div></div>
                    <span class="pl-meta">${this.fmtPct(share)}%</span></td>
            </tr>`;
        };

        let rowsHtml = PaymentLimitsManager.GROUP_CATEGORIES.map(name => rowHtml(name, byGroup.get(name), true)).join('');
        if (other.count) rowsHtml += rowHtml('Прочее (не сопоставлено)', other, false);

        return `
            <div class="pl-block pl-block-flat">
                <div class="pl-block-title"><i class="fas fa-layer-group"></i> Разбивка на категории
                    <small>по колонке «Группа» из файла · учтены только отмеченные заявки${this.modeNote()}</small></div>
                <div class="data-table-wrapper pl-table-wrap">
                <table class="data-table pl-table-narrow">
                    <thead><tr><th>Категория</th><th>Итого, ₽</th><th>Заявок</th><th style="min-width:180px;">Доля от итога</th></tr></thead>
                    <tbody>${rowsHtml}
                    <tr class="pl-grand-row"><td>ИТОГО</td>
                        <td class="number-cell">${this.fmtMoney(totalAll)}</td>
                        <td class="number-cell">${grand.count}</td>
                        <td class="pl-bar-cell"><div class="limit-bar"><div class="limit-bar-fill limit-ok" style="width:100%;"></div></div><span class="pl-meta">100%</span></td></tr>
                    </tbody>
                </table></div>
            </div>`;
    }

    // ===== БЛОК «ПОЯСНЕНИЯ»: СВОДНАЯ ПО КОНТРАГЕНТАМ + ПОЯСНЕНИЯ ИЗ БИБЛИОТЕКИ =====

    loadExplanationGroupSetting() {
        try {
            const group = localStorage.getItem('paymentLimitsExplanationGroup');
            return PaymentLimitsManager.GROUP_CATEGORIES.includes(group)
                ? group : PaymentLimitsManager.GROUP_CATEGORIES[0];
        } catch (e) {
            return PaymentLimitsManager.GROUP_CATEGORIES[0];
        }
    }

    setExplanationGroup(group) {
        const value = PaymentLimitsManager.GROUP_CATEGORIES.includes(group)
            ? group : PaymentLimitsManager.GROUP_CATEGORIES[0];
        this.explanationGroup = value;
        try {
            localStorage.setItem('paymentLimitsExplanationGroup', value);
        } catch (e) { /* localStorage недоступен — не критично */ }
        return value;
    }

    // Наименование контрагента сравниваем так же, как в «Библиотеке»: регистр и пробелы не важны
    normalizeContractorName(name) {
        return String(name || '').toUpperCase().replace(/\s+/g, ' ').trim();
    }

    // Справочник контрагентов (наименование → пояснение) из «Библиотеки контрагентов»
    async loadContractorsLibrary(force) {
        if (this.libraryLoaded && !force) return { success: true, count: this.library.size };
        try {
            const response = await fetch('/api/contractors');
            const result = await response.json();
            const list = result.data || result.contractors || [];
            if (!result.success || !Array.isArray(list)) {
                return { success: false, message: result.error || 'Библиотека контрагентов недоступна' };
            }
            this.library = new Map();
            list.forEach(item => {
                if (!item || !item.name) return;
                this.library.set(this.normalizeContractorName(item.name), {
                    id: item.id,
                    name: item.name,
                    organization: item.organization || '',
                    explanation: String(item.explanation || '').trim()
                });
            });
            this.libraryLoaded = true;
            return { success: true, count: this.library.size };
        } catch (error) {
            console.warn('Не удалось загрузить библиотеку контрагентов:', error);
            return { success: false, message: error.message };
        }
    }

    // Пояснение по наименованию получателя
    explanationFor(contractorName) {
        const item = this.library.get(this.normalizeContractorName(contractorName));
        return item ? (item.explanation || '') : '';
    }

    // Сохранение пояснения: обновляем запись в библиотеке или создаём новую
    async saveExplanation(contractorName, explanation) {
        const name = String(contractorName || '').trim();
        if (!name) return { success: false, message: 'Не указан контрагент' };
        const value = String(explanation == null ? '' : explanation).trim();
        const key = this.normalizeContractorName(name);
        const known = this.library.get(key);

        try {
            let result;
            if (known && known.id) {
                const response = await fetch(`/api/contractors/${known.id}`, {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ explanation: value })
                });
                result = await response.json();
            } else {
                const response = await fetch('/api/contractors', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ name, organization: '', explanation: value })
                });
                result = await response.json();
                if (result && result.success && result.id) {
                    this.library.set(key, { id: result.id, name, organization: '', explanation: value });
                }
            }
            if (!result || !result.success) {
                return { success: false, message: (result && result.error) || 'Не удалось сохранить пояснение' };
            }

            const current = this.library.get(key);
            if (current) {
                current.explanation = value;
            } else {
                this.library.set(key, { id: result.id || null, name, organization: '', explanation: value });
            }
            return {
                success: true,
                message: value ? `Пояснение сохранено в библиотеку: ${name}` : `Пояснение очищено: ${name}`
            };
        } catch (error) {
            return { success: false, message: 'Ошибка сохранения пояснения: ' + error.message };
        }
    }

    // Размер поля пояснения подбираем по длине текста, чтобы пояснение было видно целиком
    explanationInputSize(text) {
        const length = String(text || '').length;
        return Math.min(Math.max(length + 2, 14), 45);
    }

    // Сводная «Контрагент × Подразделение» + «Итого» + «Пояснение»
    explanationPivot() {
        const rows = this.selectedPreviewRows()
            .filter(row => this.normalizeGroup(row.group) === this.explanationGroup);

        const byContractor = new Map();
        const subdivisions = new Set();
        rows.forEach(row => {
            const contractor = (row.recipient || '').trim() || '(не указан)';
            const subdivision = (row.subdivision || '').trim() || '(не указано)';
            const amount = Number(row.amount) || 0;
            if (!byContractor.has(contractor)) byContractor.set(contractor, new Map());
            const cells = byContractor.get(contractor);
            cells.set(subdivision, (cells.get(subdivision) || 0) + amount);
            subdivisions.add(subdivision);
        });

        const columns = [...subdivisions].sort((a, b) => a.localeCompare(b, 'ru'));
        const items = [...byContractor.entries()].map(([contractor, cells]) => {
            const bySub = {};
            let total = 0;
            columns.forEach(c => {
                const value = cells.get(c) || 0;
                bySub[c] = value;
                total += value;
            });
            return { contractor, bySub, total, explanation: this.explanationFor(contractor) };
        }).sort((a, b) => a.contractor.localeCompare(b.contractor, 'ru'));

        const totals = {};
        let grand = 0;
        columns.forEach(c => {
            const value = items.reduce((sum, item) => sum + item.bySub[c], 0);
            totals[c] = value;
            grand += value;
        });

        return { columns, items, totals, grand, rowsCount: rows.length };
    }

    // Блок «Пояснения»: кнопки категорий + сводная таблица с пояснениями из библиотеки.
    // Столбец «Контрагент» и столбец «Пояснение» закреплены, середина прокручивается по горизонтали.
    explanationsHtml() {
        const groups = PaymentLimitsManager.GROUP_CATEGORIES;
        const selectedRows = this.selectedPreviewRows();
        const filter = this.previewFilter;

        const catButtons = groups.map((name, idx) => {
            const count = selectedRows.filter(row => this.normalizeGroup(row.group) === name).length;
            const id = `plExpGroup${idx}`;
            return `<input type="radio" class="pl-tab-input" name="plExpGroup" id="${id}" `
                + `value="${this.escapeHtml(name)}"${name === this.explanationGroup ? ' checked' : ''}>`
                + `<label class="pl-tab-btn" for="${id}" title="Показать пояснения и отфильтровать список заявок по категории «${this.escapeHtml(name)}»">`
                + `${this.escapeHtml(name)}<span class="pl-exp-count">${count}</span></label>`;
        }).join('');

        const pivot = this.explanationPivot();
        const columnsCount = pivot.columns.length + 3; // Контрагент + подразделения + Итого + Пояснение

        let bodyHtml;
        if (!pivot.items.length) {
            bodyHtml = `<tr><td colspan="${columnsCount}" class="pl-empty-filter">`
                + `Нет отмеченных заявок категории «${this.escapeHtml(this.explanationGroup)}»</td></tr>`;
        } else {
            const numCell = (value) => Number(value)
                ? `<td class="pl-pivot-num">${this.fmtMoney(value)}</td>`
                : `<td class="pl-pivot-num pl-pivot-zero">—</td>`;

            bodyHtml = pivot.items.map(item => {
                const activeRecipient = filter.recipient === item.contractor;
                return `
                <tr>
                    <td class="pl-pivot-name${activeRecipient ? ' pl-pivot-active' : ''}">
                        <button type="button" class="pl-cat-link" data-pl-recipient="${this.escapeHtml(item.contractor)}"
                                title="${this.escapeHtml(item.contractor)} — нажмите, чтобы отфильтровать список заявок">${this.escapeHtml(item.contractor)}</button>
                    </td>
                    ${pivot.columns.map(c => numCell(item.bySub[c])).join('')}
                    <td class="pl-pivot-num pl-pivot-total">${this.fmtMoney(item.total)}</td>
                    <td class="pl-pivot-exp">
                        <input type="text" class="pl-exp-input${item.explanation ? '' : ' pl-exp-missing'}"
                               data-pl-expl="${this.escapeHtml(item.contractor)}"
                               value="${this.escapeHtml(item.explanation)}"
                               size="${this.explanationInputSize(item.explanation)}"
                               placeholder="впишите пояснение" autocomplete="off"
                               title="Пояснение сохраняется в «Библиотеку контрагентов»">
                    </td>
                </tr>`;
            }).join('');

            bodyHtml += `
                <tr class="pl-grand-row">
                    <td class="pl-pivot-name">ИТОГО</td>
                    ${pivot.columns.map(c => `<td class="pl-pivot-num">${this.fmtMoney(pivot.totals[c])}</td>`).join('')}
                    <td class="pl-pivot-num">${this.fmtMoney(pivot.grand)}</td>
                    <td class="pl-pivot-exp"></td>
                </tr>`;
        }

        const headHtml = `<tr><th class="pl-pivot-name-col">Контрагент</th>`
            + pivot.columns.map(c => {
                const active = filter.subdivision === c;
                return `<th class="pl-pivot-sub${active ? ' pl-pivot-active' : ''}">`
                    + `<button type="button" class="pl-cat-link" data-pl-subdivision="${this.escapeHtml(c)}" `
                    + `title="Показать заявки подразделения «${this.escapeHtml(c)}» в списке ниже">`
                    + `${this.escapeHtml(c)}</button></th>`;
            }).join('')
            + `<th class="pl-pivot-total-col">Итого</th><th class="pl-pivot-exp-col">Пояснение</th></tr>`;

        return `
            <div class="pl-block pl-block-flat">
                <div class="pl-block-title"><i class="fas fa-comment-dots"></i> Пояснения
                    <small>сводная по контрагентам: ${pivot.items.length} контрагентов, ${pivot.columns.length} подразделений
                        · учтены только отмеченные заявки · пояснения из «Библиотеки»
                        · названия подразделений и контрагентов — кнопки-фильтры${this.modeNote()}</small></div>
                <div class="pl-tabs pl-exp-cats">${catButtons}</div>
                <div class="pl-pivot-wrap">
                    <table class="data-table pl-pivot">
                        <thead>${headHtml}</thead>
                        <tbody>${bodyHtml}</tbody>
                    </table>
                </div>
            </div>`;
    }

    previewRowsHtml() {
        if (!this.previewRows.length) return '';
        const rows = this.filteredPreviewRows();
        const filter = this.previewFilter;
        const hasFilter = this.hasPreviewFilter();
        const filterLabel = [filter.group, filter.category, filter.detail, filter.subdivision, filter.recipient]
            .filter(Boolean).join(' / ');
        const selectedCount = this.selectedCount();
        const selectedSum = this.selectedTotal();

        let rowsHtml = rows.map((row, idx) => {
            const selected = this.isRowSelected(row);
            return `
            <tr class="${this.rowCategory(row) ? '' : 'row-needs-review'}${selected ? '' : ' pl-row-unchecked'}">
                <td class="pl-check-cell"><input type="checkbox" data-pl-row="${row.excelRow}"${selected ? ' checked' : ''}
                        title="${selected ? 'Заявка учитывается в итогах и экспорте' : 'Заявка снята: не попадает в итоги и экспорт'}"></td>
                <td>${idx + 1}</td>
                <td class="number-cell">${this.fmtMoney(row.amount)}</td>
                <td>${this.escapeHtml(row.recipient)}</td>
                <td>${this.escapeHtml(row.articleDds)}</td>
                <td>${this.escapeHtml(row.category)}</td>
                <td>${this.escapeHtml(row.detailLimit)}</td>
                <td>${this.escapeHtml(row.requestNo)}</td>
                <td>${this.escapeHtml(row.requestDate)}</td>
                <td>${this.escapeHtml(row.status)}</td>
                <td>${this.escapeHtml(row.paymentDate)}</td>
                <td>${this.escapeHtml(row.subdivision)}</td>
                <td title="${this.escapeHtml(row.purpose)}">${this.escapeHtml(row.purpose).slice(0, 70)}</td>
            </tr>`;
        }).join('');

        if (!rows.length) {
            rowsHtml = `<tr><td colspan="13" class="pl-empty-filter">Нет заявок по фильтру «${this.escapeHtml(filterLabel)}»</td></tr>`;
        }

        return `
            <div class="pl-block pl-block-flat" id="plPreviewListBlock">
                <div class="pl-block-title"><i class="fas fa-list-alt"></i> Список загруженных заявок (превью)
                    <small>${this.fileName ? 'файл: ' + this.escapeHtml(this.fileName) : ''} · загрузка в ${this.periodLabel(this.currentPeriod)}</small>
                    <span class="pl-filter-chip">Выбрано: ${selectedCount} из ${this.previewRows.length} · ${this.fmtMoney(selectedSum)} ₽</span>
                    <span class="pl-select-btns">
                        <button type="button" class="pl-select-btn" data-pl-select="all" title="Отметить все заявки">Выбрать все</button>
                        <button type="button" class="pl-select-btn" data-pl-select="none" title="Снять все заявки — итоги обнулятся">Снять все</button>
                    </span>
                    ${hasFilter ? `<span class="pl-filter-chip">Фильтр: ${this.escapeHtml(filterLabel)} (${rows.length} из ${this.previewRows.length})
                        <button type="button" class="pl-filter-reset" data-pl-filter-reset="1" title="Сбросить фильтр">✕</button></span>` : ''}
                </div>
                <div class="data-table-wrapper pl-table-wrap">
                <table class="data-table pl-requests-table">
                    <thead><tr><th class="pl-check-col" title="Отметка заявок">${this.groupToggleHtml(rows, 'data-pl-select-toggle="1"', `Отметить/снять все заявки в списке (${rows.length})`)}</th><th>№</th><th>Сумма, ₽</th><th>Получатель</th><th>Статья движения ДС</th><th>Тип оплаты</th><th>Детальный лимит</th><th>Заявка</th><th>Дата заявки</th><th>Статус</th><th>Дата оплаты</th><th>Подразделение</th><th>Назначение платежа</th></tr></thead>
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
        const parts = [this.activeBlockHtml(), this.previewRowsHtml()];
        area.innerHTML = parts.filter(Boolean).join('');
        this.applyToggleStates();
        this.applyTabState();
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
