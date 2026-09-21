// app.js - Основной файл приложения
class App {
    constructor() {
        this.currentPage = 'upload';
        this.storage = new StorageManager();
        this.receiptsManager = new ReceiptsManager(this.storage);
        this.expensesReconciliationManager = new ExpensesReconciliationManager(this.storage);
        this.balancesManager = new BalancesManager(this.storage);
        this.debtManager = new DebtReconciliationManager(this.storage);
        this.contractorsLibrary = new ContractorsLibrary();
        this.supplierPayments = new SupplierPaymentsManager(this.contractorsLibrary);
        this.paymentLimits = new PaymentLimitsManager(this.storage);
        this.reportsManager = new ReportsManager();
        this.settingsManager = new SettingsManager(this.storage);
        window.settingsManager = this.settingsManager;
        this.parser = new BankStatementParser(); // Парсер банковских выписок
        this.originalWorkbook = null; // Для хранения оригинального файла поставщиков
        this.init();
    }

    async init() {
        console.log('Initializing Financial Analysis App...');
        this.setupNavigation();
        this.setupSidebarToggle();
        this.setupEventListeners();
        this.setupContractorsManager();
        this.reportsManager.init();
        this.settingsManager.init();
        this.setHeaderDate();

        // Устанавливаем сегодняшнюю дату по умолчанию
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('calculationDate').value = today;

        // Загружаем account_mapping из БД, чтобы accounts всегда были доступны
        // даже до загрузки выписок (для страницы «Остатки»)
        await this.storage.loadConfig();
        console.log(`✅ Account mapping загружен: ${Object.keys(this.storage.getAccounts()).length} счетов`);

        // Инициализируем депозиты из БД (с fallback на localStorage)
        await this.storage.initDeposits();

        // Загружаем библиотеку контрагентов через API (асинхронно)
        await this.contractorsLibrary.init();
        console.log(`✅ Библиотека контрагентов инициализирована: ${this.contractorsLibrary.contractors.length} записей`);

        // Загружаем маппинг счетов парсера с сервера
        this.parser.loadMappingFromServer().then(() => {
            console.log('✅ Маппинг счетов парсера загружен');
        });

        this.loadCurrentPage();
        this.updateStats();

        console.log('App initialized successfully');
    }

    setupNavigation() {
        const navButtons = document.querySelectorAll('.sidebar-nav-btn');
        const pages = document.querySelectorAll('.page');

        navButtons.forEach(button => {
            button.addEventListener('click', (e) => {
                e.preventDefault();
                const targetPage = button.getAttribute('data-page');
                this.switchPage(targetPage);
                // Закрываем мобильное меню после выбора
                document.getElementById('sidebar')?.classList.remove('open');
            });
        });
    }

    setupSidebarToggle() {
        const toggle = document.getElementById('sidebarToggle');
        const sidebar = document.getElementById('sidebar');
        if (toggle && sidebar) {
            toggle.addEventListener('click', () => {
                sidebar.classList.toggle('collapsed');
            });
        }

        const mobileBtn = document.getElementById('mobileMenuBtn');
        if (mobileBtn && sidebar) {
            mobileBtn.addEventListener('click', () => {
                sidebar.classList.toggle('open');
            });
        }
    }

    setHeaderDate() {
        const el = document.getElementById('headerDate');
        if (el) {
            const now = new Date();
            el.textContent = now.toLocaleDateString('ru-RU', {
                day: 'numeric', month: 'long', year: 'numeric'
            });
        }
    }

    switchPage(targetPage) {
        console.log('Switching to page:', targetPage);
        this.currentPage = targetPage;

        const navButtons = document.querySelectorAll('.sidebar-nav-btn');
        const pages = document.querySelectorAll('.page');

        navButtons.forEach(btn => btn.classList.remove('active'));

        const activeButton = Array.from(navButtons).find(btn =>
            btn.getAttribute('data-page') === targetPage
        );
        if (activeButton) activeButton.classList.add('active');

        pages.forEach(page => {
            page.style.display = 'none';
            page.classList.remove('active');
        });

        // Скрываем панель результатов сверки при переключении страниц
        const reconPanel = document.getElementById('reconciliationResultSection');
        if (reconPanel) reconPanel.style.display = 'none';

        const targetPageElement = document.getElementById(`${targetPage}-page`);
        if (targetPageElement) {
            targetPageElement.style.display = 'block';
            targetPageElement.classList.add('active');
            this.updateBreadcrumb(targetPage);
            this.updatePageData(targetPage);

            // Прокрутка к первому элементу страницы
            const firstHeader = targetPageElement.querySelector('.page-header h1');
            if (firstHeader) {
                firstHeader.scrollIntoView({ block: 'start', behavior: 'instant' });
            }
            const pc = document.querySelector('.page-content');
            if (pc) pc.scrollTop = 0;
            window.scrollTo(0, 0);
        }
    }

    updateBreadcrumb(pageName) {
        const el = document.getElementById('breadcrumbCurrent');
        if (!el) return;
        const names = {
            upload: 'Загрузка выписок',
            receipts: 'Поступления',
            'expenses-reconciliation': 'Сверка оплат',
            balances: 'Остатки',
            debt: 'Дебиторка',
            suppliers: 'Оплаты поставщикам',
            'payment-limits': 'Лимиты платежей',
            library: 'Библиотека',
            reports: 'Отчёты',
            settings: 'Настройки'
        };
        el.textContent = names[pageName] || pageName;
    }

    async updatePageData(pageName) {
        switch(pageName) {
            case 'receipts':
                this.receiptsManager.updateTable();
                break;
            case 'expenses-reconciliation':
                this.expensesReconciliationManager.updateTable();
                break;
            case 'balances':
                this.balancesManager.updateTable();
                break;
            case 'debt':
                await this.updateReconciliationUI();
                break;
            case 'suppliers':
                this.updateSuppliersUI();
                break;
            case 'payment-limits':
                await this.updatePaymentLimitsUI();
                break;
            case 'library':
                this.renderLibraryTable();
                break;
            case 'reports':
                break;
            case 'settings':
                await this.settingsManager.loadAllData();
                // Сброс прокрутки ПОСЛЕ загрузки данных
                const pc = document.querySelector('.page-content');
                if (pc) pc.scrollTop = 0;
                window.scrollTo(0, 0);
                break;
        }
    }

    updateSuppliersUI() {
        // Обновляем статистику на странице поставщиков
        const stats = document.getElementById('suppliersStats');
        const pivotTables = this.supplierPayments.getPivotTables();
        const libraryStats = this.contractorsLibrary.getStats();

        if (pivotTables.length > 0) {
            stats.style.display = 'grid';
            document.getElementById('suppliersRegistriesCount').textContent = pivotTables.length;
        }

        // Обновляем статистику на странице библиотеки
        this.updateLibraryStats();
    }

    updateLibraryStats() {
        const libraryStats = this.contractorsLibrary.getStats();
        document.getElementById('libTotal').textContent = libraryStats.total;
        document.getElementById('libWithExplanation').textContent = libraryStats.withExplanation;
    }

    // ✅ НОВЫЙ МЕТОД: Загрузка данных предыдущего дня из БД через API
    async loadPreviousDayDataFromServer() {
        try {
            console.log('🔄 Загрузка данных предыдущего дня из БД...');
            
            // Вычисляем вчерашнюю дату
            const yesterday = new Date();
            yesterday.setDate(yesterday.getDate() - 1);
            const dateStr = yesterday.toISOString().split('T')[0];
            
            const response = await fetch(`http://31.130.155.16:5000/api/previous-day-data?date=${dateStr}`);
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }
            
            const result = await response.json();
            
            if (result.success && result.data && Object.keys(result.data).length > 0) {
                const actualDate = result.date || dateStr;
                console.log(`✅ Загружено ${Object.keys(result.data).length} филиалов из БД за ${actualDate}`);

                // Если локально (вручную из файла) загружены данные за более позднюю дату,
                // чем последний снимок в БД, БД не перетирает их (иначе после загрузки
                // файла за 01.09 система снова показывает 26.08)
                const localInfo = this.debtManager.getPreviousDayData();
                const localData = localInfo.data || {};
                if (Object.keys(localData).length > 0 && localInfo.date && localInfo.date > actualDate) {
                    console.log(`ℹ️ Локально загружены данные за ${localInfo.date} — они свежее последнего снимка БД (${actualDate}). Данные из БД не применяем.`);
                    return { success: true, data: localData, date: localInfo.date, source: 'localStorage-newer' };
                }

                // ✅ ИСПРАВЛЕНИЕ: Загружаем сводные данные из ответа API
                if (result.summaryDT) {
                    this.debtManager.summaryDT = {
                        legal: result.summaryDT.legal || 0
                    };
                    this.debtManager.saveSummaryData();
                    console.log('✅ Загружены сводные данные ДТ из БД:', this.debtManager.summaryDT);
                }
                
                if (result.summarySIUAT) {
                    this.debtManager.summarySIUAT = {
                        totalDebt: result.summarySIUAT.totalDebt || 0,
                        totalOverdue: result.summarySIUAT.totalOverdue || 0,
                        legal: result.summarySIUAT.legal || 0
                    };
                    this.debtManager.saveSummaryData();
                    console.log('✅ Загружены сводные данные СИ УАТ из БД:', this.debtManager.summarySIUAT);
                }
                
                // Сохраняем в localStorage для кэша и офлайн-работы (включая сводные)
                localStorage.setItem('previousDayDebt_manual', JSON.stringify({
                    data: result.data,
                    date: actualDate,
                    summaryDT: result.summaryDT || {},
                    summarySIUAT: result.summarySIUAT || {}
                }));
                
                // Обновляем текущие данные менеджера
                this.debtManager.currentSubdivisionData = result.data;
                
                return { success: true, data: result.data, source: 'database' };
            } else {
                console.log('⚠️ Данные за вчерашний день не найдены в БД, пробуем localStorage');
                // Пробуем загрузить из localStorage
                return this.debtManager.getPreviousDayData();
            }
        } catch (error) {
            console.warn('❌ Ошибка загрузки из БД, используем localStorage:', error.message);
            // Fallback на localStorage если сервер недоступен
            return this.debtManager.getPreviousDayData();
        }
    }

    // ✅ ИЗМЕНЕНИЕ: Загрузка данных из БД при открытии страницы дебиторки
    async updateReconciliationUI() {
        const stats = this.debtManager.getStats();
        
        // ✅ Сначала загружаем данные из БД (независимо от браузера)
        await this.loadPreviousDayDataFromServer();
        
        if (stats.debtRows > 0 || stats.receiptsWithDates > 0) {
            this.showReconciliationStats(stats);
        }
        // Обновляем индикатор данных предыдущего дня при загрузке страницы
        this.updatePreviousDayIndicator();
    }

    setupEventListeners() {
        // Загрузка файлов (основная)
        const selectFilesBtn = document.getElementById('selectFilesBtn');
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = '.txt,.1c,.csv';
        fileInput.multiple = true;

        selectFilesBtn.addEventListener('click', () => {
            fileInput.click();
        });

        fileInput.addEventListener('change', (e) => {
            this.handleFileSelect(e.target.files);
            fileInput.value = '';
        });

        // Drag and drop для основной загрузки
        const dropArea = document.getElementById('dropArea');
        dropArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropArea.style.borderColor = '#2563eb';
            dropArea.style.background = 'rgba(37, 99, 235, 0.05)';
        });

        dropArea.addEventListener('dragleave', () => {
            dropArea.style.borderColor = '';
            dropArea.style.background = '';
        });

        dropArea.addEventListener('drop', (e) => {
            e.preventDefault();
            dropArea.style.borderColor = '';
            dropArea.style.background = '';

            if (e.dataTransfer.files.length) {
                this.handleFileSelect(e.dataTransfer.files);
            }
        });

        // Обработка файлов
        const processFilesBtn = document.getElementById('processFilesBtn');
        processFilesBtn.addEventListener('click', () => {
            this.processFiles();
        });

        // Очистка файлов
        const clearFilesBtn = document.getElementById('clearFilesBtn');
        clearFilesBtn.addEventListener('click', () => {
            this.clearFiles();
        });

        // Экспорт данных
        document.getElementById('exportReceiptsBtn').addEventListener('click', () => {
            this.receiptsManager.exportToExcel();
        });

        document.getElementById('exportBalancesBtn').addEventListener('click', () => {
            this.balancesManager.exportToExcel();
        });

        // Обновление данных
        document.getElementById('refreshReceiptsBtn').addEventListener('click', () => {
            this.receiptsManager.updateTable();
        });

        document.getElementById('calculateInterestsBtn').addEventListener('click', () => {
            this.balancesManager.calculateInterests();
        });

        // Продление срочных (длинных) депозитов до следующего рабочего дня
        const nextWorkingDayBtn = document.getElementById('nextWorkingDayBtn');
        if (nextWorkingDayBtn) {
            nextWorkingDayBtn.addEventListener('click', () => {
                this.balancesManager.extendTermDepositsToNextWorkingDay();
            });
        }

        // Загрузка ИНН
        document.getElementById('innUploadBtn').addEventListener('click', () => {
            document.getElementById('innFileUpload').click();
        });

        document.getElementById('innFileUpload').addEventListener('change', (e) => {
            this.receiptsManager.loadINNData(e.target.files[0]);
            e.target.value = '';
        });

        // Загрузка данных по депозитам
        document.getElementById('depositUploadBtn').addEventListener('click', () => {
            document.getElementById('depositFileUpload').click();
        });

        document.getElementById('depositFileUpload').addEventListener('change', (e) => {
            this.balancesManager.loadDepositData(e.target.files[0]);
            e.target.value = '';
        });

        // Очистка поиска
        document.getElementById('clearSearchBtn').addEventListener('click', () => {
            this.receiptsManager.clearSearch();
        });

        // Поиск
        document.getElementById('searchReceipts').addEventListener('input', (e) => {
            this.receiptsManager.searchTransactions(e.target.value);
        });

        // Модальное окно для депозитов
        const modal = document.getElementById('depositModal');
        const closeModalButtons = document.querySelectorAll('.close-modal');

        closeModalButtons.forEach(button => {
            button.addEventListener('click', () => {
                modal.classList.remove('active');
            });
        });

        document.getElementById('saveDepositBtn').addEventListener('click', () => {
            this.balancesManager.saveDeposit();
        });

        document.getElementById('addDepositBtn').addEventListener('click', () => {
            this.balancesManager.addDeposit();
        });

        // Enter/Escape в модальном окне депозитов
        modal.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && modal.classList.contains('active')) {
                e.preventDefault();
                this.balancesManager.saveDeposit();
            }
            if (e.key === 'Escape' && modal.classList.contains('active')) {
                modal.classList.remove('active');
            }
        });

        // Клик вне модального окна
        window.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.classList.remove('active');
            }
        });

        // ===== Обработчики для страницы сверки оплат =====
        this.setupExpensesReconciliationListeners();

        // ===== Обработчики для страницы сверки долгов =====
        this.setupDebtReconciliationListeners();

        // ===== Обработчики для страницы оплат поставщикам =====
        this.setupSupplierPaymentsListeners();

        // ===== Обработчики для страницы библиотеки =====
        this.setupLibraryListeners();

        // ===== Обработчики для страницы контроля лимитов платежей =====
        this.setupPaymentLimitsListeners();

        // ===== Обработчики для модального окна настроек парсера =====
        this.setupParserSettingsListeners();
    }

    setupDebtReconciliationListeners() {
        // Выбор файла реестра ДЗ
        const selectDebtRegistryBtn = document.getElementById('selectDebtRegistryBtn');
        const debtRegistryFile = document.getElementById('debtRegistryFile');
        const debtRegistryDropArea = document.getElementById('debtRegistryDropArea');
        const debtRegistryFileInfo = document.getElementById('debtRegistryFileInfo');

        selectDebtRegistryBtn.addEventListener('click', () => {
            debtRegistryFile.click();
        });

        debtRegistryFile.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.loadDebtRegistryFile(e.target.files[0]);
            }
        });

        // Drag and drop для реестра ДЗ
        debtRegistryDropArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            debtRegistryDropArea.style.borderColor = '#2563eb';
            debtRegistryDropArea.style.background = 'rgba(37, 99, 235, 0.05)';
        });

        debtRegistryDropArea.addEventListener('dragleave', () => {
            debtRegistryDropArea.style.borderColor = '';
            debtRegistryDropArea.style.background = '';
        });

        debtRegistryDropArea.addEventListener('drop', (e) => {
            e.preventDefault();
            debtRegistryDropArea.style.borderColor = '';
            debtRegistryDropArea.style.background = '';

            if (e.dataTransfer.files.length > 0) {
                this.loadDebtRegistryFile(e.dataTransfer.files[0]);
            }
        });

        // Выбор файла поступлений
        const selectReceiptsRegistryBtn = document.getElementById('selectReceiptsRegistryBtn');
        const receiptsRegistryFile = document.getElementById('receiptsRegistryFile');
        const receiptsRegistryDropArea = document.getElementById('receiptsRegistryDropArea');
        const receiptsRegistryFileInfo = document.getElementById('receiptsRegistryFileInfo');

        selectReceiptsRegistryBtn.addEventListener('click', () => {
            receiptsRegistryFile.click();
        });

        receiptsRegistryFile.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.loadReceiptsRegistryFile(e.target.files[0]);
            }
        });

        // Drag and drop для поступлений
        receiptsRegistryDropArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            receiptsRegistryDropArea.style.borderColor = '#2563eb';
            receiptsRegistryDropArea.style.background = 'rgba(37, 99, 235, 0.05)';
        });

        receiptsRegistryDropArea.addEventListener('dragleave', () => {
            receiptsRegistryDropArea.style.borderColor = '';
            receiptsRegistryDropArea.style.background = '';
        });

        receiptsRegistryDropArea.addEventListener('drop', (e) => {
            e.preventDefault();
            receiptsRegistryDropArea.style.borderColor = '';
            receiptsRegistryDropArea.style.background = '';

            if (e.dataTransfer.files.length > 0) {
                this.loadReceiptsRegistryFile(e.dataTransfer.files[0]);
            }
        });

        // Кнопка сверки
        document.getElementById('reconcileBtn').addEventListener('click', () => {
            this.performReconciliation();
        });

        // Кнопка экспорта
        document.getElementById('exportReconciledBtn').addEventListener('click', () => {
            this.exportReconciledFile();
        });

        // Кнопка отправки по почте
        document.getElementById('sendEmailBtn').addEventListener('click', () => {
            this.sendEmail();
        });

        // Кнопка очистки
        document.getElementById('clearReconciliationBtn').addEventListener('click', () => {
            this.clearReconciliationData();
        });

        // Кнопка настроек сводных таблиц
        document.getElementById('summarySettingsBtn').addEventListener('click', () => {
            this.openSummarySettingsModal();
        });

        // Закрытие модального окна настроек
        const closeSummaryModal = document.getElementById('closeSummarySettingsModal');
        closeSummaryModal.addEventListener('click', () => {
            document.getElementById('summarySettingsModal').classList.remove('active');
        });

        // Сохранение настроек сводных
        document.getElementById('saveSummarySettingsBtn').addEventListener('click', () => {
            this.saveSummarySettings();
        });

        // Enter/Escape в модальном окне настроек сводных
        const summarySettingsModal = document.getElementById('summarySettingsModal');
        summarySettingsModal.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && summarySettingsModal.classList.contains('active')) {
                e.preventDefault();
                this.saveSummarySettings();
            }
            if (e.key === 'Escape' && summarySettingsModal.classList.contains('active')) {
                summarySettingsModal.classList.remove('active');
            }
        });

        // Кнопка сохранения данных текущего дня (ручное сохранение)
        document.getElementById('saveCurrentDayBtn').addEventListener('click', () => {
            // Используем новый метод forceCollectAndSave для ручного сохранения
            const result = this.debtManager.forceCollectAndSave();
            if (result.success) {
                this.showNotification(`Сохранено ${result.count} подразделений`, 'success');
                this.updatePreviousDayIndicator();
            } else {
                this.showNotification(result.message, 'error');
            }
        });

        // Кнопка очистки данных предыдущего дня
        document.getElementById('clearPreviousDayDataBtn').addEventListener('click', () => {
            this.clearPreviousDayData();
        });

        // Загрузка данных предыдущего дня из Excel файла
        const loadPreviousDayFileBtn = document.getElementById('loadPreviousDayFileBtn');
        const previousDayFileInput = document.getElementById('previousDayFileInput');
        const previousDayFileInfo = document.getElementById('previousDayFileInfo');

        loadPreviousDayFileBtn.addEventListener('click', () => {
            previousDayFileInput.click();
        });

        previousDayFileInput.addEventListener('change', async (e) => {
            if (e.target.files.length > 0) {
                const file = e.target.files[0];
                // Передаём дату из поля «Дата предыдущего дня», если в имени файла её нет
                const dateInput = document.getElementById('previousDayDateInput');
                const result = await this.loadPreviousDayDataFromFile(file, { date: dateInput ? dateInput.value : '' });

                if (result.success) {
                    // Обновляем информацию о файле (с датой, если она определена)
                    const dateLabel = result.date ? ` за ${result.date}` : '';
                    previousDayFileInfo.innerHTML = `<i class="fas fa-check-circle" style="color: var(--success);"></i> ${file.name}${dateLabel} (${result.count} ДТ, ${this.formatCurrency(result.total)})`;

                    // Обновляем таблицу
                    this.renderPreviousDayDataTable();

                    // Обновляем индикатор
                    this.updatePreviousDayIndicator();

                    this.showNotification(result.message, 'success');
                } else {
                    this.showNotification(result.message, 'error');
                }

                // Сбрасываем input, чтобы можно было загрузить тот же файл повторно
                e.target.value = '';
            }
        });

        // Загрузка файла СИ УАТ
        const selectSiUatBtn = document.getElementById('selectSiUatBtn');
        const siUatFileInput = document.getElementById('siUatFile');
        const siUatDropArea = document.getElementById('siUatDropArea');
        const siUatFileInfo = document.getElementById('siUatFileInfo');

        selectSiUatBtn.addEventListener('click', () => {
            siUatFileInput.click();
        });

        siUatFileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.loadSiUatFile(e.target.files[0]);
            }
        });

        siUatDropArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            siUatDropArea.style.borderColor = '#2563eb';
            siUatDropArea.style.background = 'rgba(37, 99, 235, 0.05)';
        });

        siUatDropArea.addEventListener('dragleave', () => {
            siUatDropArea.style.borderColor = '';
            siUatDropArea.style.background = '';
        });

        siUatDropArea.addEventListener('drop', (e) => {
            e.preventDefault();
            siUatDropArea.style.borderColor = '';
            siUatDropArea.style.background = '';

            if (e.dataTransfer.files.length > 0) {
                this.loadSiUatFile(e.dataTransfer.files[0]);
            }
        });

        // Клик вне модального окна настроек
        window.addEventListener('click', (e) => {
            if (e.target === document.getElementById('summarySettingsModal')) {
                document.getElementById('summarySettingsModal').classList.remove('active');
            }
        });

        // ===== НОВЫЕ ОБРАБОТЧИКИ ДЛЯ РАБОТЫ С БД ЧЕРЕЗ API =====
        
        // Кнопка загрузки данных предыдущего дня из БД
        const loadPreviousDayFromDBBtn = document.getElementById('loadPreviousDayFromDBBtn');
        const previousDayDateInput = document.getElementById('previousDayDateInput');
        
        if (loadPreviousDayFromDBBtn && previousDayDateInput) {
            loadPreviousDayFromDBBtn.addEventListener('click', async () => {
                const date = previousDayDateInput.value;
                if (!date) {
                    this.showNotification('Укажите дату для загрузки', 'error');
                    return;
                }
                
                this.showLoading();
                try {
                    const result = await this.debtManager.loadPreviousDayDataFromAPI(date);
                    
                    if (result.success) {
                        // Показываем ФАКТИЧЕСКУЮ дату из БД: если за 01.09 данных нет,
                        // сервер вернёт последний доступный снимок (например, 26.08)
                        const actualDate = result.date || date;
                        // Обновляем информацию о файле
                        previousDayFileInfo.innerHTML = `<i class="fas fa-database" style="color: var(--success);"></i> Загружено из БД: ${actualDate} (${Object.keys(result.data).length} ДТ)`;

                        // Обновляем таблицу
                        this.renderPreviousDayDataTable();

                        // Обновляем индикатор
                        this.updatePreviousDayIndicator();

                        this.showNotification(`Загружено ${Object.keys(result.data).length} филиалов за ${actualDate}`, 'success');
                    } else {
                        this.showNotification(result.message || 'Ошибка загрузки данных', 'error');
                    }
                } catch (error) {
                    console.error('Ошибка загрузки из БД:', error);
                    this.showNotification('Ошибка подключения к серверу', 'error');
                } finally {
                    this.hideLoading();
                }
            });
        }

        // Кнопка сохранения данных текущего дня в БД
        const saveCurrentDayToDBBtn = document.getElementById('saveCurrentDayToDBBtn');
        
        if (saveCurrentDayToDBBtn) {
            saveCurrentDayToDBBtn.addEventListener('click', async () => {
                const date = previousDayDateInput.value || new Date().toISOString().split('T')[0];
                
                this.showLoading();
                try {
                    // Сначала собираем актуальные данные
                    this.debtManager.collectSubdivisionData(true);
                    
                    const result = await this.debtManager.saveCurrentDayDataToAPI(this.debtManager.currentSubdivisionData, date);
                    
                    if (result.success) {
                        this.showNotification(`Сохранено ${result.count} филиалов в БД за ${date}`, 'success');
                        this.updatePreviousDayIndicator();
                    } else {
                        const msg = result.fallback ? 'Данные сохранены локально (БД недоступна)' : (result.error || 'Ошибка сохранения');
                        this.showNotification(msg, result.fallback ? 'warning' : 'error');
                    }
                } catch (error) {
                    console.error('Ошибка сохранения в БД:', error);
                    this.showNotification('Ошибка подключения к серверу', 'error');
                } finally {
                    this.hideLoading();
                }
            });
        }
    }

    // ===== ОБРАБОТЧИКИ ДЛЯ СТРАНИЦЫ СВЕРКИ ОПЛАТ =====
    setupExpensesReconciliationListeners() {
        // Выбор Excel-файла для сверки
        const selectExcelBtn = document.getElementById('selectExpensesExcelBtn');
        const excelFileInput = document.getElementById('expensesExcelFile');

        selectExcelBtn.addEventListener('click', () => {
            excelFileInput.click();
        });

        excelFileInput.addEventListener('change', async (e) => {
            if (e.target.files.length > 0) {
                await this.loadExpensesExcelFile(e.target.files[0]);
            }
        });

        // Кнопка сверки
        document.getElementById('reconcileExpensesBtn').addEventListener('click', () => {
            this.performExpensesReconciliation();
        });

        // Кнопка экспорта
        document.getElementById('exportExpensesReconciliationBtn').addEventListener('click', () => {
            this.expensesReconciliationManager.exportToExcel();
        });

        // Кнопка очистки
        document.getElementById('clearExpensesReconciliationBtn').addEventListener('click', () => {
            this.expensesReconciliationManager.clearAll();
        });

        // Поиск
        document.getElementById('searchExpenses').addEventListener('input', (e) => {
            this.expensesReconciliationManager.searchTransactions(e.target.value);
        });

        // Очистка поиска
        document.getElementById('clearExpensesSearchBtn').addEventListener('click', () => {
            this.expensesReconciliationManager.clearSearch();
        });
    }

    async loadExpensesExcelFile(file) {
        this.showLoading();
        try {
            const result = await this.expensesReconciliationManager.loadExcelFile(file);
            if (result.success) {
                document.getElementById('expensesExcelFileInfo').innerHTML =
                    `<i class="fas fa-check-circle" style="color: var(--success);"></i> ${file.name} (${result.count} строк)`;
                document.getElementById('reconcileExpensesBtn').disabled = false;
                this.showNotification(result.message, 'success');
            } else {
                document.getElementById('expensesExcelFileInfo').innerHTML =
                    `<i class="fas fa-exclamation-circle" style="color: var(--danger);"></i> ${result.message}`;
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка загрузки Excel:', error);
            this.showNotification('Ошибка загрузки файла', 'error');
        } finally {
            this.hideLoading();
        }
    }

    performExpensesReconciliation() {
        this.showLoading();
        try {
            const result = this.expensesReconciliationManager.reconcile();
            if (result.success) {
                this.expensesReconciliationManager.renderReconciliationResult();
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка сверки:', error);
            this.showNotification('Ошибка при выполнении сверки', 'error');
        } finally {
            this.hideLoading();
        }
    }

    // Настройка управления контрагентами
    setupContractorsManager() {
        const manageContractorsBtn = document.getElementById('manageContractorsBtn');
        const contractorsModal = document.getElementById('contractorsModal');
        const closeModal = document.getElementById('closeContractorsModal');
        const cancelBtn = document.getElementById('cancelContractorsBtn');
        const saveBtn = document.getElementById('saveContractorsBtn');
        const addBtn = document.getElementById('addContractorBtn');
        const contractorsList = document.getElementById('contractorsList');

        // Открытие модального окна
        manageContractorsBtn.addEventListener('click', () => {
            this.renderContractorsList();
            contractorsModal.classList.add('active');
        });

        // Закрытие модального окна
        const closeModalFn = () => {
            contractorsModal.classList.remove('active');
        };

        closeModal.addEventListener('click', closeModalFn);
        cancelBtn.addEventListener('click', closeModalFn);

        // Клик вне модального окна
        window.addEventListener('click', (e) => {
            if (e.target === contractorsModal) {
                contractorsModal.classList.remove('active');
            }
        });

        // Добавление нового контрагента
        addBtn.addEventListener('click', () => {
            const contractor = prompt('Введите наименование контрагента:');
            if (contractor && contractor.trim() !== '') {
                if (this.debtManager.addTargetContractor(contractor)) {
                    this.renderContractorsList();
                    this.showNotification('Контрагент добавлен', 'success');
                } else {
                    this.showNotification('Контрагент уже существует или некорректное название', 'error');
                }
            }
        });

        // Сохранение изменений (фактически уже сохранено, но обновляем интерфейс)
        saveBtn.addEventListener('click', () => {
            contractorsModal.classList.remove('active');
            this.showNotification('Список контрагентов сохранен', 'success');
        });

        // Enter/Escape в модальном окне контрагентов (дебиторка)
        contractorsModal.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && contractorsModal.classList.contains('active')) {
                e.preventDefault();
                contractorsModal.classList.remove('active');
                this.showNotification('Список контрагентов сохранен', 'success');
            }
            if (e.key === 'Escape' && contractorsModal.classList.contains('active')) {
                contractorsModal.classList.remove('active');
            }
        });
    }

    // Отрисовка списка контрагентов в модальном окне
    renderContractorsList() {
        const contractorsList = document.getElementById('contractorsList');
        const contractors = this.debtManager.getTargetContractors();

        if (contractors.length === 0) {
            contractorsList.innerHTML = '<p class="empty-message">Список контрагентов пуст. Добавьте хотя бы одного.</p>';
            return;
        }

        let html = '';
        contractors.forEach(contractor => {
            html += `
                <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px; border-bottom: 1px solid var(--border);">
                    <span>${contractor}</span>
                    <button class="btn btn-danger btn-sm remove-contractor" data-contractor="${contractor}" style="padding: 4px 8px;">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            `;
        });

        contractorsList.innerHTML = html;

        // Добавляем обработчики для кнопок удаления
        contractorsList.querySelectorAll('.remove-contractor').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                const contractor = btn.dataset.contractor;
                if (confirm(`Удалить контрагента "${contractor}"?`)) {
                    if (this.debtManager.removeTargetContractor(contractor)) {
                        this.renderContractorsList();
                        this.showNotification('Контрагент удален', 'success');
                    }
                }
            });
        });
    }

    async loadDebtRegistryFile(file) {
        this.showLoading();
        try {
            const result = await this.debtManager.loadDebtRegistryFile(file);
            document.getElementById('debtRegistryFileInfo').innerHTML =
                `<i class="fas fa-check-circle" style="color: var(--success);"></i> ${file.name} (${result.message})`;

            // Активируем кнопку сверки, если оба файла загружены
            this.updateReconcileButtonState();

            if (result.success) {
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            this.showNotification('Ошибка загрузки файла', 'error');
        } finally {
            this.hideLoading();
        }
    }

    async loadReceiptsRegistryFile(file) {
        this.showLoading();
        try {
            const result = await this.debtManager.loadReceiptsFile(file);
            document.getElementById('receiptsRegistryFileInfo').innerHTML =
                `<i class="fas fa-check-circle" style="color: var(--success);"></i> ${file.name} (${result.message})`;

            // Активируем кнопку сверки, если оба файла загружены
            this.updateReconcileButtonState();

            if (result.success) {
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            this.showNotification('Ошибка загрузки файла', 'error');
        } finally {
            this.hideLoading();
        }
    }

    updateReconcileButtonState() {
        const stats = this.debtManager.getStats();
        const reconcileBtn = document.getElementById('reconcileBtn');
        reconcileBtn.disabled = !(stats.debtRows > 0 && stats.receiptsWithDates > 0);
    }

    performReconciliation() {
        this.showLoading();
        try {
            const result = this.debtManager.reconcile();
            if (result.success) {
                this.showReconciliationStats(this.debtManager.getStats());
                this.showReconciliationLog(this.debtManager.getProcessedLog());
                document.getElementById('exportReconciledBtn').disabled = false;
                document.getElementById('saveCurrentDayBtn').disabled = false;
                document.getElementById('sendEmailBtn').disabled = false;
                
                // ✅ ИСПРАВЛЕНИЕ: УБРАЛИ преждевременное сохранение!
                // Сохранение происходит ТОЛЬКО в exportToExcel() после ответа сервера
                
                // Обновляем индикатор данных предыдущего дня (показывает текущее состояние localStorage)
                this.updatePreviousDayIndicator();
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка сверки:', error);
            this.showNotification('Ошибка при выполнении сверки', 'error');
        } finally {
            this.hideLoading();
        }
    }

    showReconciliationStats(stats) {
        document.getElementById('statTotalDocuments').textContent = stats.totalDocuments || 0;
        document.getElementById('statFoundDocuments').textContent = stats.foundDocuments || 0;
        document.getElementById('statUpdatedDocuments').textContent = stats.updatedDocuments || 0;
        document.getElementById('statReceiptsWithDates').textContent = stats.receiptsWithDates || 0;
        document.getElementById('reconciliationStats').style.display = 'block';
    }

    showReconciliationLog(log) {
        const logContainer = document.getElementById('logContainer');
        const logSection = document.getElementById('reconciliationLog');

        if (!log || log.length === 0) {
            logContainer.innerHTML = '<p class="empty-message">Нет обработанных документов</p>';
        } else {
            let html = '<table class="log-table"><tr><th>Документ</th><th>Действие</th><th>Дата</th><th>Сумма</th></tr>';
            log.slice(0, 50).forEach(item => {
                const dateStr = item.date ? new Date(item.date).toLocaleDateString('ru-RU') : '';
                html += `<tr>
                    <td>${item.documentName}</td>
                    <td><span class="badge badge-success">${item.action}</span></td>
                    <td>${dateStr}</td>
                    <td class="number-cell">${this.storage.formatNumber(item.amount || 0)}</td>
                </tr>`;
            });
            if (log.length > 50) {
                html += `<tr><td colspan="4" class="text-center">... и еще ${log.length - 50} записей</td></tr>`;
            }
            html += '</table>';
            logContainer.innerHTML = html;
        }

        logSection.style.display = 'block';
    }

    async exportReconciledFile() {
        const progressBar = document.getElementById('saveProgressBar');
        const progressFill = document.getElementById('saveProgressFill');
        const progressText = document.getElementById('saveProgressText');
        const exportBtn = document.getElementById('exportReconciledBtn');

        try {
            // Показываем прогресс-бар
            progressBar.style.display = 'block';
            exportBtn.disabled = true;
            progressFill.style.width = '10%';
            progressText.textContent = 'Подготовка данных...';

            // Анимация прогресса
            let progress = 10;
            const progressInterval = setInterval(() => {
                if (progress < 80) {
                    progress += Math.random() * 15;
                    if (progress > 80) progress = 80;
                    progressFill.style.width = progress + '%';
                    if (progress < 40) {
                        progressText.textContent = 'Обработка документов...';
                    } else if (progress < 60) {
                        progressText.textContent = 'Формирование отчёта...';
                    } else {
                        progressText.textContent = 'Создание файлов...';
                    }
                }
            }, 500);

            const result = await this.debtManager.exportToExcel();

            // Останавливаем анимацию
            clearInterval(progressInterval);
            progressFill.style.width = '100%';
            progressText.textContent = 'Готово!';

            if (result.success) {
                this.showNotification(result.message, 'success');
            } else if (result.message) {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка при экспорте:', error);
            this.showNotification('Ошибка при сохранении файла', 'error');
        } finally {
            // Скрываем прогресс-бар через небольшую задержку
            setTimeout(() => {
                progressBar.style.display = 'none';
                progressFill.style.width = '0%';
                exportBtn.disabled = false;
            }, 1500);
        }
    }

    async sendEmail() {
        const sendEmailBtn = document.getElementById('sendEmailBtn');
        this.showLoading();
        try {
            sendEmailBtn.disabled = true;
            const result = await this.debtManager.sendToEmail();
            if (result.success) {
                this.showNotification(result.message, 'success');
                // После отправки писем снимок текущего дня сохранён — обновляем индикатор,
                // чтобы показать дату, которая станет «предыдущим днём» для следующей сверки
                this.updatePreviousDayIndicator();
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка при отправке по почте:', error);
            this.showNotification('Ошибка при формировании писем', 'error');
        } finally {
            sendEmailBtn.disabled = false;
            this.hideLoading();
        }
    }

    clearReconciliationData() {
        this.debtManager.clearData();
        document.getElementById('debtRegistryFileInfo').innerHTML = 'Файл не выбран';
        document.getElementById('receiptsRegistryFileInfo').innerHTML = 'Файл не выбран';
        document.getElementById('siUatFileInfo').innerHTML = 'Файл не выбран';
        document.getElementById('reconciliationStats').style.display = 'none';
        document.getElementById('reconciliationLog').style.display = 'none';
        document.getElementById('exportReconciledBtn').disabled = true;
        document.getElementById('reconcileBtn').disabled = true;
        document.getElementById('sendEmailBtn').disabled = true;
        document.getElementById('previousDayIndicator').style.display = 'none';
        this.showNotification('Данные очищены', 'info');
    }

    // Загрузка файла СИ УАТ
    async loadSiUatFile(file) {
        this.showLoading();
        try {
            const result = await this.debtManager.loadSiUatFile(file);
            document.getElementById('siUatFileInfo').innerHTML =
                `<i class="fas fa-check-circle" style="color: var(--success);"></i> ${file.name}`;

            if (result.success) {
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            this.showNotification('Ошибка загрузки файла СИ УАТ', 'error');
        } finally {
            this.hideLoading();
        }
    }

    // Загрузка данных предыдущего дня из Excel файла
    async loadPreviousDayDataFromFile(file, options) {
        this.showLoading();
        try {
            const result = await this.debtManager.loadPreviousDayDataFromFile(file, options || {});
            const fileInfoEl = document.getElementById('previousDayFileInfo');

            if (result.success) {
                const dateLabel = result.date ? ` за ${result.date}` : '';
                fileInfoEl.innerHTML =
                    `<i class="fas fa-check-circle" style="color: var(--success);"></i> ${file.name}${dateLabel} (${result.count} ДТ, ${this.formatCurrency(result.total)})`;
                this.showNotification(result.message, 'success');
            } else {
                fileInfoEl.innerHTML =
                    `<i class="fas fa-exclamation-circle" style="color: var(--danger);"></i> ${result.message}`;
                this.showNotification(result.message, 'error');
            }

            return result;
        } catch (error) {
            this.showNotification('Ошибка загрузки файла: ' + error.message, 'error');
            return { success: false, message: error.message };
        } finally {
            this.hideLoading();
        }
    }

    // Открытие модального окна настроек сводных
    openSummarySettingsModal() {
        // Заполняем поля сводных данных
        document.getElementById('summaryDtLegal').value = this.debtManager.summaryDT.legal || '';

        document.getElementById('summarySiuatTotalDebt').value = this.debtManager.summarySIUAT.totalDebt || '';
        document.getElementById('summarySiuatTotalOverdue').value = this.debtManager.summarySIUAT.totalOverdue || '';
        document.getElementById('summarySiuatLegal').value = this.debtManager.summarySIUAT.legal || '';

        // Заполняем дату предыдущего дня
        const previousInfo = this.debtManager.getPreviousDayData();
        document.getElementById('previousDayDateInput').value = previousInfo.date || '';

        // Заполняем таблицу данных предыдущего дня из localStorage
        this.renderPreviousDayDataTable();

        document.getElementById('summarySettingsModal').classList.add('active');
    }

    // Отрисовка таблицы данных предыдущего дня
    renderPreviousDayDataTable() {
        const tbody = document.getElementById('previousDayDataBody');
        const statsEl = document.getElementById('previousDayStats');
        const previousInfo = this.debtManager.getPreviousDayData();
        const previousData = previousInfo.data || previousInfo; // совместимость со старым форматом

        if (Object.keys(previousData).length === 0) {
            tbody.innerHTML = '<tr class="empty-row"><td colspan="2">Данные не загружены. Загрузите Excel файл.</td></tr>';
            statsEl.style.display = 'none';
            return;
        }

        // Показываем статистику
        statsEl.style.display = 'block';
        const dtCount = Object.keys(previousData).length;
        const totalAmount = Object.values(previousData).reduce((sum, val) => sum + val, 0);

        document.getElementById('previousDayDtCount').textContent = dtCount;
        document.getElementById('previousDayDtTotal').textContent = this.formatCurrency(totalAmount);

        // Таблица — только для просмотра (данные из файла, не редактируемые)
        let html = '';
        const sortedFilials = Object.keys(previousData).sort();

        sortedFilials.forEach(filial => {
            const amount = previousData[filial] || 0;
            html += `<tr>
                <td>${this.escapeHtml(filial)}</td>
                <td class="number-cell">${this.formatNumber(amount)}</td>
            </tr>`;
        });

        tbody.innerHTML = html;
    }

    // Форматирование валюты для отображения
    formatCurrency(amount) {
        return new Intl.NumberFormat('ru-RU', {
            style: 'currency',
            currency: 'RUB',
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(amount);
    }

    // Форматирование числа
    formatNumber(num) {
        return new Intl.NumberFormat('ru-RU', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(num);
    }

    // Очистка данных предыдущего дня
    clearPreviousDayData() {
        if (!confirm('Вы уверены, что хотите удалить все данные предыдущего дня?')) {
            return;
        }

        const result = this.debtManager.clearPreviousDayData();

        if (result.success) {
            // Сбрасываем информацию о файле
            document.getElementById('previousDayFileInfo').innerHTML = '<i class="fas fa-info-circle"></i> Файл не загружен';
            document.getElementById('previousDayFileInput').value = '';
            document.getElementById('previousDayDateInput').value = '';

            this.renderPreviousDayDataTable();
            this.updatePreviousDayIndicator();

            this.showNotification('Данные предыдущего дня очищены', 'success');
        } else {
            this.showNotification(result.message, 'error');
        }
    }

    // Сохранение настроек сводных
    async saveSummarySettings() {
        // Читаем данные предыдущего дня ДО изменения сводных значений.
        // ВАЖНО: getPreviousDayData() подхватывает сводные из localStorage
        // (previousDayDebt_manual) и перезаписывает ими this.summaryDT/summarySIUAT.
        // Если вызвать его после ввода новых значений — новая сумма «в т.ч. Судебная»
        // будет затёрта старым значением ещё до сохранения в localStorage.
        const previousInfo = this.debtManager.getPreviousDayData();
        const previousDate = document.getElementById('previousDayDateInput').value;

        // Сохраняем сводные данные ДТ
        this.debtManager.summaryDT = {
            legal: parseFloat(document.getElementById('summaryDtLegal').value) || 0
        };

        // Сохраняем сводные данные СИ УАТ
        this.debtManager.summarySIUAT = {
            totalDebt: parseFloat(document.getElementById('summarySiuatTotalDebt').value) || 0,
            totalOverdue: parseFloat(document.getElementById('summarySiuatTotalOverdue').value) || 0,
            legal: parseFloat(document.getElementById('summarySiuatLegal').value) || 0
        };

        // ✅ ИСПРАВЛЕНИЕ: Отправляем сводные данные на сервер в БД
        // чтобы они были доступны при открытии в другом браузере
        const date = previousDate || new Date().toISOString().split('T')[0];
        const filialData = this.debtManager.currentSubdivisionData || {};
        
        try {
            await this.debtManager.saveCurrentDayDataToAPI(filialData, date);
            console.log('✅ Сводные данные отправлены на сервер');
        } catch (e) {
            console.warn('⚠️ Не удалось отправить сводные на сервер:', e);
        }

        // Сохраняем дату предыдущего дня
        if (Object.keys(previousInfo.data || previousInfo).length > 0 || previousDate) {
            localStorage.setItem('previousDayDebt_manual', JSON.stringify({
                data: previousInfo.data || previousInfo,
                date: previousDate,
                summaryDT: this.debtManager.summaryDT,
                summarySIUAT: this.debtManager.summarySIUAT
            }));
        }

        // Данные предыдущего дня сохраняются автоматически при загрузке из файла
        // Здесь просто сохраняем сводные данные
        this.debtManager.saveSummaryData();

        // Закрываем модалку
        document.getElementById('summarySettingsModal').classList.remove('active');

        // Обновляем индикатор
        this.updatePreviousDayIndicator();

        this.showNotification('Настройки сводных таблиц сохранены', 'success');
    }

    // Обновление индикатора данных предыдущего дня
    updatePreviousDayIndicator() {
        const indicator = document.getElementById('previousDayIndicator');
        const indicatorText = document.getElementById('previousDayIndicatorText');
        const previousInfo = this.debtManager.getPreviousDayData();
        const previousData = previousInfo.data || previousInfo; // совместимость со старым форматом

        if (Object.keys(previousData).length > 0) {
            indicator.style.display = 'flex';
            indicator.className = 'previous-day-indicator loaded';
            const dateStr = previousInfo.date ? ` (${previousInfo.date})` : '';
            indicatorText.textContent = `Данные за предыдущий день загружены${dateStr} (${Object.keys(previousData).length} подразделений)`;
        } else {
            indicator.style.display = 'flex';
            indicator.className = 'previous-day-indicator warning';
            indicatorText.textContent = 'Данные за предыдущий день не заполнены. Откройте Настройки сводных.';
        }
    }

    handleFileSelect(files) {
        if (!files || files.length === 0) return;

        this.storage.addFiles(Array.from(files));
        this.updateFileList();
        this.updateStats();
    }

    updateFileList() {
        const fileListContent = document.getElementById('fileListContent');
        const fileListSection = document.getElementById('fileListSection');
        const statsSection = document.getElementById('statsSection');
        const files = this.storage.getFiles();

        // Показываем/скрываем секции в зависимости от наличия файлов
        if (files.length === 0) {
            fileListSection.style.display = 'none';
            statsSection.style.display = 'none';
            return;
        }

        // Показываем секции при наличии файлов
        fileListSection.style.display = 'block';
        if (this.storage.getTransactions().length > 0) {
            statsSection.style.display = 'block';
        }

        if (files.length === 0) {
            fileListContent.innerHTML = '<p class="empty-message">Файлы не загружены</p>';
            return;
        }

        let html = '';
        files.forEach((file) => {
            const size = file.size > 1024 * 1024
                ? `${(file.size / (1024 * 1024)).toFixed(2)} MB`
                : `${(file.size / 1024).toFixed(2)} KB`;

            html += `
                <div class="file-item">
                    <div class="file-name">${file.name}</div>
                    <div class="file-size">${size}</div>
                </div>
            `;
        });

        fileListContent.innerHTML = html;
    }

    async processFiles() {
        const files = this.storage.getFiles();
        if (files.length === 0) {
            alert('Пожалуйста, сначала загрузите файлы выписок');
            return;
        }

        this.showLoading();

        try {
            // Загружаем настройки (исключения, категоризация, синонимы) с сервера перед парсингом
            await this.storage.loadConfig();
            console.log('Настройки загружены:', {
                exclusionRules: this.storage.getExclusionRules().length,
                categorizationRules: this.storage.getCategorizationRules().length,
                companyAliases: this.storage.getCompanyAliases().length
            });

            // Инициализируем парсер настройками из БД (company aliases для нормализации)
            await this.parser.init(this.storage);

            const results = await this.parser.processFiles(files);

            // Сохраняем данные
            this.storage.setStatements(results.statements);
            this.storage.setTransactions(results.transactions);

            // Обновляем счета данными из выписок (merge, не перезапись),
            // чтобы сохранить все счета из account_mapping в таблице «Остатки»
            for (const [accountNumber, accountData] of Object.entries(results.accounts)) {
                this.storage.updateAccount(accountNumber, accountData);
            }

            // Обновляем интерфейс
            this.updateStats();
            this.receiptsManager.updateTable();
            this.expensesReconciliationManager.updateTable();
            this.balancesManager.updateTable();

            // Показываем уведомление
            this.showNotification(`Обработано ${files.length} файлов. Найдено ${results.transactions.length} операций.`, 'success');

        } catch (error) {
            console.error('Error processing files:', error);
            this.showNotification('Ошибка при обработке файлов', 'error');
        } finally {
            this.hideLoading();
        }
    }

    async clearFiles() {
        if (confirm('Вы уверены, что хотите очистить список файлов?')) {
            this.storage.clearFiles();
            // Восстанавливаем account_mapping из БД после очистки,
            // чтобы счета всегда отображались в таблице «Остатки»
            await this.storage.loadConfig();
            this.updateFileList();
            this.updateStats();
            this.receiptsManager.updateTable();
            this.expensesReconciliationManager.updateTable();
            this.balancesManager.updateTable();
        }
    }

    updateStats() {
        const files = this.storage.getFiles();
        const transactions = this.storage.getTransactions();
        const accounts = this.storage.getAccounts();

        document.getElementById('filesCount').textContent = files.length;
        document.getElementById('operationsCount').textContent = transactions.length;
        document.getElementById('accountsCount').textContent = Object.keys(accounts).length;

        // Показываем статистику только если есть обработанные данные
        const statsSection = document.getElementById('statsSection');
        if (transactions.length > 0) {
            statsSection.style.display = 'block';
        }
    }

    showLoading() {
        const overlay = document.createElement('div');
        overlay.className = 'loading-overlay active';
        overlay.innerHTML = '<div class="loading-spinner"></div>';
        document.body.appendChild(overlay);
    }

    hideLoading() {
        const overlay = document.querySelector('.loading-overlay');
        if (overlay) {
            overlay.remove();
        }
    }

    showNotification(message, type = 'info') {
        // Создаем уведомление
        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.innerHTML = `
            <div class="notification-content">
                <i class="fas fa-${type === 'success' ? 'check-circle' : type === 'error' ? 'exclamation-circle' : 'info-circle'}"></i>
                <span>${message}</span>
            </div>
        `;

        document.body.appendChild(notification);

        // Удаляем уведомление через 5 секунд
        setTimeout(() => {
            notification.style.animation = 'fadeOut 0.3s ease';
            setTimeout(() => notification.remove(), 300);
        }, 5000);
    }

    loadCurrentPage() {
        this.updatePageData(this.currentPage);
    }

    // ===== ОБРАБОТЧИКИ ДЛЯ СТРАНИЦЫ ОПЛАТ ПОСТАВЩИКАМ =====
    setupSupplierPaymentsListeners() {
        // Загрузка файла
        const uploadBtn = document.getElementById('uploadSuppliersBtn');
        const fileInput = document.getElementById('suppliersFileUpload');

        uploadBtn.addEventListener('click', () => {
            fileInput.click();
        });

        fileInput.addEventListener('change', async (e) => {
            if (e.target.files.length > 0) {
                await this.handleSuppliersFileUpload(e.target.files[0]);
                e.target.value = '';
            }
        });

        // Переход в библиотеку
        document.getElementById('goToLibraryBtn').addEventListener('click', () => {
            this.switchPage('library');
        });

        // Закрытие модальных окон (только для депозитов и контрагентов дебиторки)
        document.querySelectorAll('.close-modal').forEach(btn => {
            btn.addEventListener('click', () => {
                btn.closest('.modal').classList.remove('active');
            });
        });

        // Закрытие по клику вне модального окна
        window.addEventListener('click', (e) => {
            if (e.target.classList.contains('modal')) {
                e.target.classList.remove('active');
            }
        });
    }

    async handleSuppliersFileUpload(file) {
        this.showLoading();
        try {
            // Сохраняем оригинальный файл для экспорта
            this.originalSuppliersFile = file;

            const result = await this.supplierPayments.loadExcelFile(file);
            if (result.success) {
                this.renderSuppliersContent();
                this.updateSuppliersUI();
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка загрузки файла поставщиков:', error);
            this.showNotification('Ошибка загрузки файла', 'error');
        } finally {
            this.hideLoading();
        }
    }

    renderSuppliersContent() {
        const emptyState = document.getElementById('suppliersEmptyState');
        const dataArea = document.getElementById('suppliersDataArea');
        const pivotTables = this.supplierPayments.getPivotTables();

        if (pivotTables.length === 0) {
            // Показываем empty-state, скрываем данные
            emptyState.style.display = '';
            dataArea.innerHTML = '';
            return;
        }

        // Скрываем empty-state, показываем данные
        emptyState.style.display = 'none';

        // ✅ Кнопка экспорта — отображается вверху, перед таблицами
        let html = `
            <div style="display: flex; justify-content: flex-end; gap: 12px; margin-bottom: 24px;">
                <button id="exportSuppliersBtn" class="btn btn-success">
                    <i class="fas fa-file-excel"></i> Экспорт в Excel
                </button>
            </div>
        `;

        pivotTables.forEach((pivot, index) => {
            html += `
                <div style="background: var(--bg-card); border: 1px solid var(--border-light); border-radius: var(--radius-lg); padding: 24px; margin-bottom: 24px; box-shadow: var(--shadow-sm);">
                    <h3 style="margin-bottom: 16px; font-size: 16px; font-weight: 600;">
                        <i class="fas fa-table" style="color: var(--primary);"></i> ${pivot.sheetName}
                    </h3>
                    <div class="table-container">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>Контрагент</th>
                                    ${pivot.pivotHeaders.map(h => `<th>${h}</th>`).join('')}
                                    <th>Итого</th>
                                    <th>Пояснение</th>
                                </tr>
                            </thead>
                            <tbody>
            `;

            let totals = {};
            pivot.pivotHeaders.forEach(h => totals[h] = 0);
            let grandTotal = 0;

            pivot.pivotData.forEach(row => {
                const rowClass = row.needsReview ? ' class="row-needs-review"' : '';
                html += `<tr${rowClass}>`;
                html += `<td><strong>${row.contractor}</strong></td>`;

                pivot.pivotHeaders.forEach(h => {
                    const value = row[h] || 0;
                    totals[h] = (totals[h] || 0) + value;
                    html += `<td class="number-cell">${this.storage.formatNumber(value)}</td>`;
                });

                grandTotal += row.total;
                html += `<td class="number-cell" style="font-weight: 600;">${this.storage.formatNumber(row.total)}</td>`;

                // Подсветка пояснения, требующего проверки
                const explStyle = row.needsReview
                    ? ' style="color: var(--warning); font-size: 13px; font-style: italic;" title="Требует проверки — смысл не извлечён"'
                    : ' style="color: var(--text-secondary); font-size: 13px;"';
                html += `<td${explStyle}>${row.explanation || ''}</td>`;
                html += `</tr>`;
            });

            // Итоговая строка
            html += `<tr style="background: var(--bg-tertiary); font-weight: 600;">`;
            html += `<td>ИТОГО</td>`;
            pivot.pivotHeaders.forEach(h => {
                html += `<td class="number-cell">${this.storage.formatNumber(totals[h] || 0)}</td>`;
            });
            html += `<td class="number-cell">${this.storage.formatNumber(grandTotal)}</td>`;
            html += `<td></td>`;
            html += `</tr>`;

            html += `
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        });

        dataArea.innerHTML = html;

        // Обработчик экспорта
        document.getElementById('exportSuppliersBtn').addEventListener('click', () => {
            this.exportSuppliersToExcel();
        });
    }

    async exportSuppliersToExcel() {
        this.showLoading();
        try {
            const result = await this.supplierPayments.exportToExcel(this.originalSuppliersFile);
            if (result.success) {
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка экспорта:', error);
            this.showNotification('Ошибка при экспорте', 'error');
        } finally {
            this.hideLoading();
        }
    }

    // ===== СТРАНИЦА КОНТРОЛЯ ЛИМИТОВ ПЛАТЕЖЕЙ =====
    setupPaymentLimitsListeners() {
        const uploadBtn = document.getElementById('uploadPaymentLimitsBtn');
        const fileInput = document.getElementById('paymentLimitsFileUpload');

        if (uploadBtn && fileInput) {
            uploadBtn.addEventListener('click', () => {
                fileInput.click();
            });

            fileInput.addEventListener('change', async (e) => {
                if (e.target.files.length > 0) {
                    await this.handlePaymentLimitsFileUpload(e.target.files[0]);
                    e.target.value = '';
                }
            });
        }

        const confirmBtn = document.getElementById('confirmPaymentLimitsBtn');
        if (confirmBtn) {
            confirmBtn.addEventListener('click', () => {
                this.handlePaymentLimitsConfirm();
            });
        }

        // Сохранение стадий: черновик (стадия 1) и финальный вариант (стадия 2)
        const draftBtn = document.getElementById('saveDraftPaymentLimitsBtn');
        if (draftBtn) {
            draftBtn.addEventListener('click', () => this.handlePaymentLimitsSave('draft'));
        }
        const finalBtn = document.getElementById('saveFinalPaymentLimitsBtn');
        if (finalBtn) {
            finalBtn.addEventListener('click', () => this.handlePaymentLimitsSave('final'));
        }

        const cancelBtn = document.getElementById('cancelPaymentLimitsBtn');
        if (cancelBtn) {
            cancelBtn.addEventListener('click', () => {
                this.handlePaymentLimitsCancel();
            });
        }

        // Настройка лимитов
        const settingsBtn = document.getElementById('plLimitsSettingsBtn');
        if (settingsBtn) {
            settingsBtn.addEventListener('click', () => {
                this.openPaymentLimitsSettings();
            });
        }

        const limitsSaveBtn = document.getElementById('plLimitsSaveBtn');
        if (limitsSaveBtn) {
            limitsSaveBtn.addEventListener('click', () => {
                this.savePaymentLimitsSettings();
            });
        }

        const limitsResetBtn = document.getElementById('plLimitsResetBtn');
        if (limitsResetBtn) {
            limitsResetBtn.addEventListener('click', () => {
                this.paymentLimits.resetEditorToDefaults();
            });
        }

        // Чекбокс «Только текущий период» (без переноса неиспользованного остатка)
        const onlyCurrentPeriod = document.getElementById('plOnlyCurrentPeriod');
        if (onlyCurrentPeriod) {
            onlyCurrentPeriod.checked = this.paymentLimits.onlyCurrentPeriod;
            onlyCurrentPeriod.addEventListener('change', () => {
                this.paymentLimits.setOnlyCurrentPeriod(onlyCurrentPeriod.checked);
                this.renderPaymentLimitsUI();
            });
        }

        // Выбор периода загрузки реестра (Период 1..4 или авто по дате)
        const loadPeriodSelect = document.getElementById('plLoadPeriod');
        if (loadPeriodSelect) {
            loadPeriodSelect.addEventListener('change', () => {
                const period = this.paymentLimits.setLoadPeriod(loadPeriodSelect.value);
                this.renderPaymentLimitsUI();
                this.showNotification(`Период загрузки: ${this.paymentLimits.periodLabel(period)}`, 'info');
            });
        }

        // Горизонтальное меню блоков (радио-вкладки: активной всегда ровно одна)
        const tabsBar = document.getElementById('plTabs');
        if (tabsBar) {
            tabsBar.querySelectorAll('input[name="plTabs"]').forEach(radio => {
                radio.addEventListener('change', () => {
                    if (!radio.checked) return;
                    this.paymentLimits.setActiveTab(radio.value);
                    this.renderPaymentLimitsUI();
                    // Если блок оказался выше экрана — поднимаем его в поле зрения
                    const area = document.getElementById('plControlArea');
                    if (area && area.getBoundingClientRect().top < 0) {
                        area.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                });
            });
        }

        // Клик по категории/подстатье/группе — фильтрация списка загруженных заявок
        const plControlArea = document.getElementById('plControlArea');
        if (plControlArea) {
            plControlArea.addEventListener('click', (e) => this.handlePaymentLimitsControlClick(e));
            // Отметка/снятие заявок чекбоксами — пересчёт итогов во всех блоках
            plControlArea.addEventListener('change', (e) => this.handlePaymentLimitsRowCheck(e));
            // Enter в поле пояснения — сохраняем сразу
            plControlArea.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' && e.target && e.target.hasAttribute
                    && e.target.hasAttribute('data-pl-expl')) {
                    e.preventDefault();
                    e.target.blur();
                }
            });
        }

        // Экспорт реестра в исходном формате (без снятых заявок)
        const exportBtn = document.getElementById('exportPaymentLimitsBtn');
        if (exportBtn) {
            exportBtn.addEventListener('click', () => this.handlePaymentLimitsExport());
        }

        // Обнуление ранее загруженных данных месяца (для повторного прохождения цикла)
        const resetBtn = document.getElementById('resetPaymentLimitsBtn');
        if (resetBtn) {
            resetBtn.addEventListener('click', () => this.handlePaymentLimitsReset());
        }

        document.querySelectorAll('#paymentLimitsSettingsModal .close-modal').forEach(btn => {
            btn.addEventListener('click', () => {
                document.getElementById('paymentLimitsSettingsModal').classList.remove('active');
            });
        });
    }

    // Фильтрация списка заявок по клику на категорию/подстатью/группу
    handlePaymentLimitsControlClick(e) {
        const manager = this.paymentLimits;

        const resetBtn = e.target.closest('[data-pl-filter-reset]');
        if (resetBtn) {
            manager.setPreviewFilter(null, null, null);
            this.renderPaymentLimitsUI();
            return;
        }

        // «Выбрать все» / «Снять все» заявки
        const selectBtn = e.target.closest('[data-pl-select]');
        if (selectBtn) {
            e.preventDefault();
            if (!manager.previewRows.length) return;
            const selectAll = selectBtn.getAttribute('data-pl-select') === 'all';
            manager.setAllSelected(selectAll);
            this.renderPaymentLimitsUI();
            this.showNotification(selectAll
                ? `Отмечены все заявки: ${manager.selectedCount()}, сумма ${manager.fmtMoney(manager.selectedTotal())} ₽`
                : 'Все заявки сняты — итоги по блокам обнулены', 'info');
            return;
        }

        // Клик по группе в блоке «Разбивка на категории»
        const groupLink = e.target.closest('[data-pl-group]');
        if (groupLink) {
            e.preventDefault();
            if (!manager.previewRows.length) {
                this.showNotification('Загрузите реестр — список заявок пуст', 'info');
                return;
            }
            manager.setPreviewFilter(null, null, groupLink.getAttribute('data-pl-group'));
            this.renderPaymentLimitsUI();
            const blockG = document.getElementById('plPreviewListBlock');
            if (blockG) blockG.scrollIntoView({ behavior: 'smooth', block: 'start' });
            return;
        }

        // Клик по подразделению в блоке «Пояснения» — фильтр списка заявок
        const subdivisionLink = e.target.closest('[data-pl-subdivision]');
        if (subdivisionLink) {
            e.preventDefault();
            if (!manager.previewRows.length) {
                this.showNotification('Загрузите реестр — список заявок пуст', 'info');
                return;
            }
            manager.setPreviewFilter(null, null, null, subdivisionLink.getAttribute('data-pl-subdivision'), null);
            this.renderPaymentLimitsUI();
            const blockS = document.getElementById('plPreviewListBlock');
            if (blockS) blockS.scrollIntoView({ behavior: 'smooth', block: 'start' });
            return;
        }

        // Клик по контрагенту в блоке «Пояснения» — фильтр списка заявок
        const recipientLink = e.target.closest('[data-pl-recipient]');
        if (recipientLink) {
            e.preventDefault();
            if (!manager.previewRows.length) {
                this.showNotification('Загрузите реестр — список заявок пуст', 'info');
                return;
            }
            manager.setPreviewFilter(null, null, null, null, recipientLink.getAttribute('data-pl-recipient'));
            this.renderPaymentLimitsUI();
            const blockR = document.getElementById('plPreviewListBlock');
            if (blockR) blockR.scrollIntoView({ behavior: 'smooth', block: 'start' });
            return;
        }

        const link = e.target.closest('[data-pl-category]');
        if (!link) return;
        e.preventDefault();

        if (!manager.previewRows.length) {
            this.showNotification('Загрузите реестр — список заявок пуст', 'info');
            return;
        }

        manager.setPreviewFilter(
            link.getAttribute('data-pl-category'),
            link.getAttribute('data-pl-detail'),
            null
        );
        this.renderPaymentLimitsUI();

        // Прокручиваем к отфильтрованному списку заявок
        const block = document.getElementById('plPreviewListBlock');
        if (block) block.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    // Отметка/снятие заявки чекбоксом — итоги во всех блоках пересчитываются.
    // Поддерживает: отдельную заявку, всю категорию/подстатью/группу и все заявки в списке.
    handlePaymentLimitsRowCheck(e) {
        const input = e.target;
        if (!input || input.tagName !== 'INPUT') return;
        const manager = this.paymentLimits;

        // Правка пояснения прямо в сводной таблице блока «Пояснения»
        if (input.hasAttribute && input.hasAttribute('data-pl-expl')) {
            this.handlePaymentLimitsExplanation(input);
            return;
        }

        // Кнопки-категории внутри блока «Пояснения»: перестраивают сводную
        // и одновременно фильтруют список заявок ниже
        if (input.type === 'radio' && input.name === 'plExpGroup') {
            if (!input.checked) return;
            manager.setExplanationGroup(input.value);
            manager.setPreviewFilter(null, null, input.value);
            this.renderPaymentLimitsUI();
            return;
        }

        if (input.type !== 'checkbox') return;
        let changed = 0;

        if (input.hasAttribute('data-pl-select-toggle')) {
            changed = manager.setRowsSelected(manager.filteredPreviewRows(), input.checked);
        } else if (input.hasAttribute('data-pl-cat-toggle')) {
            changed = manager.setRowsSelected(manager.rowsByCategory(input.getAttribute('data-pl-cat-toggle')), input.checked);
        } else if (input.hasAttribute('data-pl-detail-toggle')) {
            changed = manager.setRowsSelected(
                manager.rowsByDetail(input.getAttribute('data-pl-cat'), input.getAttribute('data-pl-detail-toggle')),
                input.checked);
        } else if (input.hasAttribute('data-pl-group-toggle')) {
            changed = manager.setRowsSelected(manager.rowsByGroup(input.getAttribute('data-pl-group-toggle')), input.checked);
        } else if (input.hasAttribute('data-pl-row')) {
            if (!manager.setRowSelected(input.getAttribute('data-pl-row'), input.checked)) return;
            changed = 1;
        } else {
            return;
        }

        if (changed) this.renderPaymentLimitsUI();
    }

    // Сохранение пояснения из сводной таблицы → сразу в «Библиотеку контрагентов»
    async handlePaymentLimitsExplanation(input) {
        const manager = this.paymentLimits;
        const contractor = input.getAttribute('data-pl-expl');
        const previous = manager.explanationFor(contractor);
        const value = String(input.value || '').trim();
        if (value === previous) return;

        try {
            const result = await manager.saveExplanation(contractor, value);
            if (result.success) {
                input.classList.toggle('pl-exp-missing', !value);
                input.classList.add('pl-exp-saved');
                setTimeout(() => input.classList.remove('pl-exp-saved'), 1500);
                this.showNotification(result.message, 'success');
            } else {
                input.value = previous;
                input.classList.toggle('pl-exp-missing', !previous);
                this.showNotification(result.message || 'Не удалось сохранить пояснение', 'error');
            }
        } catch (error) {
            console.error('Ошибка сохранения пояснения:', error);
            input.value = previous;
            this.showNotification('Ошибка сохранения пояснения', 'error');
        }
    }

    // Экспорт реестра заявок в исходном формате (только отмеченные)
    async handlePaymentLimitsExport() {
        const manager = this.paymentLimits;
        this.showLoading();
        try {
            const result = await manager.exportToExcel();
            if (result.success) {
                this.renderPaymentLimitsUI();   // обновляем кружок-уведомление на кнопке
                this.showNotification(result.message, result.warning ? 'info' : 'success');
            } else {
                this.showNotification(result.message || 'Ошибка экспорта', 'error');
            }
        } catch (error) {
            console.error('Ошибка экспорта реестра лимитов:', error);
            this.showNotification('Ошибка экспорта', 'error');
        } finally {
            this.hideLoading();
        }
    }

    // Открытие/сохранение настроек лимитов
    openPaymentLimitsSettings() {
        const manager = this.paymentLimits;
        manager.renderEditorHtml();
        const modal = document.getElementById('paymentLimitsSettingsModal');
        if (modal) modal.classList.add('active');
    }

    async savePaymentLimitsSettings() {
        const manager = this.paymentLimits;
        const config = manager.collectEditorConfig();
        if (!config.categories.length || config.categories.some(c => !c.name || !c.name.trim())) {
            this.showNotification('Название лимита не может быть пустым', 'error');
            return;
        }
        this.showLoading();
        try {
            const result = await manager.saveConfig(config);
            if (result.success) {
                const modal = document.getElementById('paymentLimitsSettingsModal');
                if (modal) modal.classList.remove('active');
                // Названия могли измениться — перечитываем подтверждённые данные
                await manager.loadConfirmed(manager.currentMonth);
                this.renderPaymentLimitsUI();
                let message = 'Настройки лимитов сохранены';
                if (result.renames && result.renames.length) {
                    const affected = result.renames.reduce((a, r) => a + (r.rows || 0), 0);
                    message += ` (переименований: ${result.renames.length}, обновлено записей: ${affected})`;
                }
                this.showNotification(message, 'success');
            } else {
                this.showNotification(result.message || 'Ошибка сохранения настроек', 'error');
            }
        } catch (error) {
            console.error('Ошибка сохранения настроек лимитов:', error);
            this.showNotification('Ошибка сохранения настроек', 'error');
        } finally {
            this.hideLoading();
        }
    }

    // Показ/обновление страницы лимитов (вызывается при переключении на страницу)
    async updatePaymentLimitsUI() {
        const manager = this.paymentLimits;
        manager.renderHeaderInfo();

        // Сначала настройки лимитов (нужны для нормализации категорий/подстатей)
        if (!manager.configLoaded) {
            const cfgRes = await manager.loadConfig();
            if (!cfgRes.success) {
                console.warn('Настройки лимитов недоступны, используются значения по умолчанию');
            }
        }
        if (!manager.confirmedLoaded) {
            const res = await manager.loadConfirmed();
            if (!res.success) {
                console.warn('Не удалось загрузить подтверждённые платежи:', res.message);
            }
        }

        // Справочник контрагентов (пояснения) для блока «Пояснения»
        const libRes = await manager.loadContractorsLibrary();
        if (!libRes.success) {
            console.warn('Библиотека контрагентов недоступна:', libRes.message);
        }

        // Сохранённый вариант месяца (черновик/финал) — подгружается автоматически,
        // чтобы коллега на другом компьютере сразу видел данные
        const stateRes = await manager.loadRegistryState();
        if (stateRes.success && stateRes.registry) {
            if (!manager.previewRows.length) {
                const loaded = await manager.loadRegistryPreview();
                if (loaded.success) {
                    this.showNotification(loaded.message, 'info');
                } else if (loaded.message && !/не найден/.test(loaded.message)) {
                    console.warn('Автозагрузка реестра лимитов:', loaded.message);
                }
            }
        }

        this.renderPaymentLimitsUI();
    }

    renderPaymentLimitsUI() {
        const manager = this.paymentLimits;
        manager.renderHeaderInfo();
        manager.renderControlArea();

        // Синхронизируем чекбокс режима «только текущий период»
        const onlyCurrentPeriod = document.getElementById('plOnlyCurrentPeriod');
        if (onlyCurrentPeriod) onlyCurrentPeriod.checked = manager.onlyCurrentPeriod;

        // Синхронизируем селект периода загрузки
        const loadPeriodSelect = document.getElementById('plLoadPeriod');
        if (loadPeriodSelect) loadPeriodSelect.value = manager.selectedPeriod ? String(manager.selectedPeriod) : '';

        const actions = document.getElementById('plActions');
        const hint = document.getElementById('plEmptyHint');

        // Кнопка экспорта доступна только при загруженном реестре
        const exportBtn = document.getElementById('exportPaymentLimitsBtn');
        if (exportBtn) exportBtn.disabled = manager.previewRows.length === 0 && !(manager.registry && manager.registry.has_file);

        // Красный кружок на кнопке экспорта: финальный вариант согласован и ещё не выгружен
        const exportBadge = document.getElementById('exportPaymentLimitsBadge');
        if (exportBadge) exportBadge.hidden = !manager.finalExportPending();

        // Стадия 1 (черновик) и стадия 2 (финал): сохраняем только отмеченные заявки
        const draftSaveBtn = document.getElementById('saveDraftPaymentLimitsBtn');
        if (draftSaveBtn) draftSaveBtn.disabled = manager.previewRows.length === 0;
        const finalSaveBtn = document.getElementById('saveFinalPaymentLimitsBtn');
        if (finalSaveBtn) finalSaveBtn.disabled = manager.previewRows.length === 0 || manager.selectedCount() === 0;

        // Состояние сохранённого варианта месяца
        const registryState = document.getElementById('plRegistryState');
        if (registryState) registryState.textContent = manager.registryStatusText();

        if (actions) {
            actions.style.display = manager.previewRows.length > 0 ? 'flex' : 'none';
        }

        // Подпись слева в закреплённой полосе кнопок
        const actionsInfo = document.getElementById('plActionsInfo');
        if (actionsInfo) {
            actionsInfo.textContent = manager.previewRows.length
                ? `Отмечено: ${manager.selectedCount()} из ${manager.previewRows.length} · `
                  + `${manager.fmtMoney(manager.selectedTotal())} ₽ · ${manager.registryStatusText()}`
                : '';
        }
        if (hint) {
            if (manager.previewRows.length > 0) {
                const stage = manager.savedStage();
                const stageText = stage === 'final' ? 'открыт сохранённый финальный вариант'
                    : (stage === 'draft' ? 'открыт сохранённый черновик' : 'данные не сохранены');
                hint.textContent = `Период загрузки: ${manager.periodLabel(manager.currentPeriod)} · ${stageText}. `
                    + `Отмечено заявок: ${manager.selectedCount()} из ${manager.previewRows.length} `
                    + `(${manager.fmtMoney(manager.selectedTotal())} ₽)`;
            } else {
                hint.textContent = `Загрузите Excel-реестр или дождитесь автозагрузки сохранённого варианта. `
                    + `Сохранено за месяц: ${manager.confirmedRows.length} платежей.`;
            }
            hint.style.color = 'var(--text-secondary)';
        }
    }

    async handlePaymentLimitsFileUpload(file) {
        this.showLoading();
        try {
            const manager = this.paymentLimits;
            const result = await manager.loadExcelFile(file);
            if (result.success) {
                // Убеждаемся, что подтверждённые данные загружены (после загрузки файла)
                if (!manager.confirmedLoaded) {
                    await manager.loadConfirmed(manager.currentMonth);
                }
                // Обновляем список сохранённых загрузок месяца
                await manager.loadRegistryState();
                // Ручная загрузка — это новая партия: не продолжаем ранее сохранённую
                manager.markAsNewUpload();
                this.renderPaymentLimitsUI();
                this.showNotification(result.message + '. Период загрузки: ' + manager.periodLabel(manager.currentPeriod), 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка загрузки файла лимитов:', error);
            this.showNotification('Ошибка загрузки файла', 'error');
        } finally {
            this.hideLoading();
        }
    }

    // Сохранение стадии: 'draft' — черновик (стадия 1), 'final' — финальный вариант (стадия 2)
    async handlePaymentLimitsSave(stage) {
        const manager = this.paymentLimits;
        if (!manager.previewRows.length) return;
        this.showLoading();
        try {
            const result = await manager.saveStage(stage);
            if (result.success) {
                // После сохранения обновляем список сохранённых загрузок периода
                // и перечитываем сохранённые данные (для накопительных итогов)
                await manager.loadRegistryState();
                await manager.loadConfirmed(manager.currentMonth);
                this.renderPaymentLimitsUI();
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message || 'Ошибка сохранения', 'error');
            }
        } catch (error) {
            console.error('Ошибка сохранения варианта лимитов:', error);
            this.showNotification('Ошибка сохранения', 'error');
        } finally {
            this.hideLoading();
        }
    }

    // Обнуление ранее загруженных данных месяца (чтобы пройти полный цикл заново)
    async handlePaymentLimitsReset() {
        const manager = this.paymentLimits;
        const confirmed = window.confirm(
            'Обнулить ранее загруженные данные за текущий месяц?\n\n'
            + 'Будут удалены: сохранённый черновик, финальный вариант, платежи и исходный файл реестра. '
            + 'После этого можно снова загрузить файл и пройти обе стадии согласования.');
        if (!confirmed) return;

        this.showLoading();
        try {
            const result = await manager.resetSavedData();
            if (result.success) {
                await manager.loadConfirmed();
                this.renderPaymentLimitsUI();
                this.showNotification(result.message, 'success');
            } else {
                this.showNotification(result.message || 'Ошибка обнуления', 'error');
            }
        } catch (error) {
            console.error('Ошибка обнуления данных лимитов:', error);
            this.showNotification('Ошибка обнуления', 'error');
        } finally {
            this.hideLoading();
        }
    }

    // Совместимость: старое «Подтвердить и сохранить» = финальный вариант (стадия 2)
    handlePaymentLimitsConfirm() {
        return this.handlePaymentLimitsSave('final');
    }

    handlePaymentLimitsCancel() {
        const manager = this.paymentLimits;
        manager.cancelPreview();
        this.renderPaymentLimitsUI();
        this.showNotification('Загрузка отменена, данные не сохранены', 'info');
    }

    // ===== СТРАНИЦА БИБЛИОТЕКИ КОНТРАГЕНТОВ =====
    setupLibraryListeners() {
        // Инициализация фильтров
        this.libraryFilters = {
            name: new Set(),
            organization: new Set(),
            explanation: new Set()
        };
        this.librarySearchTerm = '';
        this.selectedRowId = null;

        // Добавление контрагента
        document.getElementById('libAddBtn').addEventListener('click', () => {
            this.openAddContractorModal();
        });

        // Импорт
        document.getElementById('libImportBtn').addEventListener('click', () => {
            document.getElementById('libraryImportFile').click();
        });

        document.getElementById('libraryImportFile').addEventListener('change', async (e) => {
            if (e.target.files.length > 0) {
                await this.importLibraryFromExcel(e.target.files[0]);
                e.target.value = '';
            }
        });

        // Экспорт
        document.getElementById('libExportBtn').addEventListener('click', () => {
            this.contractorsLibrary.exportToExcel();
        });

        // Очистка
        document.getElementById('libClearBtn').addEventListener('click', async () => {
            if (!confirm('Вы уверены, что хотите очистить всю библиотеку контрагентов?')) return;
            const result = await this.contractorsLibrary.clearAll();
            if (result.success) {
                this.renderLibraryTable();
                this.updateLibraryStats();
                this.showNotification(result.message, 'info');
            }
        });

        // Поиск
        document.getElementById('librarySearch').addEventListener('input', (e) => {
            this.librarySearchTerm = e.target.value;
            this.renderLibraryTable();
        });

        // Очистка поиска
        document.getElementById('libClearSearchBtn').addEventListener('click', () => {
            document.getElementById('librarySearch').value = '';
            this.librarySearchTerm = '';
            this.renderLibraryTable();
        });

        // Кнопки фильтров в заголовках
        document.querySelectorAll('.filter-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const filterType = btn.dataset.filter;
                this.toggleFilterDropdown(filterType, btn);
            });
        });

        // Закрытие фильтров
        document.querySelectorAll('.filter-close').forEach(btn => {
            btn.addEventListener('click', () => {
                const filterType = btn.dataset.filter;
                this.closeFilterDropdown(filterType);
            });
        });

        // Поиск внутри фильтров
        document.querySelectorAll('[data-filter-search]').forEach(input => {
            input.addEventListener('input', (e) => {
                const filterType = e.target.dataset.filterSearch;
                this.filterOptionsSearch(filterType, e.target.value);
            });
        });

        // Чекбокс "Выбрать все"
        document.querySelectorAll('.select-all-cb').forEach(cb => {
            cb.addEventListener('change', (e) => {
                const filterType = e.target.dataset.filterSelect;
                this.toggleSelectAll(filterType, e.target.checked);
            });
        });

        // Применить фильтр
        document.querySelectorAll('[data-filter-apply]').forEach(btn => {
            btn.addEventListener('click', () => {
                const filterType = btn.dataset.filterApply;
                this.applyFilter(filterType);
            });
        });

        // Сбросить фильтр
        document.querySelectorAll('[data-filter-clear]').forEach(btn => {
            btn.addEventListener('click', () => {
                const filterType = btn.dataset.filterClear;
                this.clearFilter(filterType);
            });
        });

        // Сохранение контрагента (используем новый id из модального окна)
        document.getElementById('saveContractorBtn').addEventListener('click', async () => {
            await this.saveContractorFromModal();
        });

        // Закрытие модального окна контрагента
        document.getElementById('closeContractorModal').addEventListener('click', () => {
            document.getElementById('contractorModal').style.display = 'none';
        });
        document.getElementById('cancelContractorBtn').addEventListener('click', () => {
            document.getElementById('contractorModal').style.display = 'none';
        });

        // Enter в модальном окне контрагента → сохранить
        const contractorModal = document.getElementById('contractorModal');
        contractorModal.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && contractorModal.style.display !== 'none') {
                e.preventDefault();
                this.saveContractorFromModal();
            }
            if (e.key === 'Escape' && contractorModal.style.display !== 'none') {
                contractorModal.style.display = 'none';
            }
        });

        // Подтверждение удаления
        document.getElementById('closeConfirmDeleteModal').addEventListener('click', () => {
            document.getElementById('confirmDeleteModal').style.display = 'none';
        });
        document.getElementById('cancelDeleteBtn').addEventListener('click', () => {
            document.getElementById('confirmDeleteModal').style.display = 'none';
        });
        document.getElementById('confirmDeleteBtn').addEventListener('click', async () => {
            await this.performDeleteContractor();
        });

        // Enter/Escape в модальном окне подтверждения удаления
        const confirmDeleteModal = document.getElementById('confirmDeleteModal');
        confirmDeleteModal.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && confirmDeleteModal.style.display !== 'none') {
                e.preventDefault();
                this.performDeleteContractor();
            }
            if (e.key === 'Escape' && confirmDeleteModal.style.display !== 'none') {
                confirmDeleteModal.style.display = 'none';
            }
        });

        // Закрытие фильтров при клике на оверлей
        document.getElementById('filterDropdowns').addEventListener('click', (e) => {
            if (e.target === document.getElementById('filterDropdowns')) {
                this.closeAllFilterDropdowns();
            }
        });

        // Закрытие фильтров при клике вне
        document.addEventListener('click', (e) => {
            if (!e.target.closest('.filter-dropdown') && !e.target.closest('.filter-btn')) {
                this.closeAllFilterDropdowns();
            }
        });

        // Горячие клавиши
        document.addEventListener('keydown', (e) => {
            // Escape закрывает фильтры
            if (e.key === 'Escape') {
                this.closeAllFilterDropdowns();
            }
            // Enter открывает редактирование выбранной строки (только если не открыто модальное окно)
            const anyModalOpen = document.querySelector('.modal-overlay[style*="display: flex"], .modal-overlay[style*="display: block"], .modal.active');
            if (e.key === 'Enter' && this.selectedRowId && !anyModalOpen) {
                this.openEditContractorModal(this.selectedRowId);
            }
        });
    }

    toggleFilterDropdown(filterType, btn) {
        const dropdown = document.getElementById(`filter${this.capitalizeFirst(filterType)}`);
        const isOpen = dropdown.style.display === 'block';

        // Закрываем все фильтры
        this.closeAllFilterDropdowns();

        if (!isOpen) {
            // Позиционируем фильтр под кнопкой
            const rect = btn.getBoundingClientRect();
            dropdown.style.left = rect.left + 'px';
            dropdown.style.top = (rect.bottom + 4) + 'px';
            dropdown.style.display = 'block';
            document.getElementById('filterDropdowns').style.display = 'flex';

            // Заполняем опции
            this.populateFilterOptions(filterType);

            // Подсвечиваем кнопку
            btn.classList.add('active');
        }
    }

    closeFilterDropdown(filterType) {
        const dropdown = document.getElementById(`filter${this.capitalizeFirst(filterType)}`);
        dropdown.style.display = 'none';
        document.querySelector(`.filter-btn[data-filter="${filterType}"]`)?.classList.remove('active');

        // Проверяем, есть ли открытые фильтры
        const anyOpen = document.querySelectorAll('.filter-dropdown[style*="block"]');
        if (anyOpen.length === 0) {
            document.getElementById('filterDropdowns').style.display = 'none';
        }
    }

    toggleSelectAll(filterType, checked) {
        const container = document.getElementById(`filterOptions${this.capitalizeFirst(filterType)}`);
        const checkboxes = container.querySelectorAll('input[type="checkbox"]');
        checkboxes.forEach(cb => {
            cb.checked = checked;
        });
    }

    updateSelectAllState(filterType) {
        const container = document.getElementById(`filterOptions${this.capitalizeFirst(filterType)}`);
        const checkboxes = container.querySelectorAll('input[type="checkbox"]');
        const selectAllCb = document.querySelector(`.select-all-cb[data-filter-select="${filterType}"]`);

        if (checkboxes.length === 0) return;

        const allChecked = Array.from(checkboxes).every(cb => cb.checked);
        selectAllCb.checked = allChecked;
    }

    closeAllFilterDropdowns() {
        ['name', 'organization', 'explanation'].forEach(type => {
            this.closeFilterDropdown(type);
        });
    }

    populateFilterOptions(filterType) {
        const contractors = this.contractorsLibrary.getAll();
        const container = document.getElementById(`filterOptions${this.capitalizeFirst(filterType)}`);

        // Собираем уникальные значения
        const values = new Map();
        contractors.forEach(c => {
            const val = c[filterType] || '(пусто)';
            values.set(val, (values.get(val) || 0) + 1);
        });

        // Сортируем
        const sorted = Array.from(values.entries()).sort((a, b) => a[0].localeCompare(b[0]));

        // Получаем текущие выбранные
        const selected = this.libraryFilters[filterType];

        let html = '';
        sorted.forEach(([value, count]) => {
            const isChecked = selected.size === 0 || selected.has(value);
            html += `
                <label class="filter-option">
                    <input type="checkbox" value="${this.escapeHtml(value)}" ${isChecked ? 'checked' : ''}>
                    <span class="filter-option-label" title="${this.escapeHtml(value)}">${this.escapeHtml(value)}</span>
                    <span class="filter-option-count">${count}</span>
                </label>
            `;
        });

        container.innerHTML = html;

        // Обработчики чекбоксов
        container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
            cb.addEventListener('change', () => {
                this.updateSelectAllState(filterType);
            });
            cb.addEventListener('click', (e) => {
                e.stopPropagation();
            });
        });
    }

    filterOptionsSearch(filterType, searchTerm) {
        const container = document.getElementById(`filterOptions${this.capitalizeFirst(filterType)}`);
        const options = container.querySelectorAll('.filter-option');
        const term = searchTerm.toLowerCase();

        options.forEach(option => {
            const label = option.querySelector('.filter-option-label').textContent.toLowerCase();
            option.style.display = label.includes(term) ? '' : 'none';
        });
    }

    applyFilter(filterType) {
        const container = document.getElementById(`filterOptions${this.capitalizeFirst(filterType)}`);
        const checkboxes = container.querySelectorAll('input[type="checkbox"]');
        const selected = new Set();

        checkboxes.forEach(cb => {
            if (cb.checked) {
                selected.add(cb.value);
            }
        });

        // Если выбраны все — очищаем фильтр (показываем всё)
        const totalOptions = container.querySelectorAll('.filter-option').length;
        if (selected.size === totalOptions || selected.size === 0) {
            this.libraryFilters[filterType] = new Set();
        } else {
            this.libraryFilters[filterType] = selected;
        }

        this.closeFilterDropdown(filterType);
        this.renderLibraryTable();
    }

    clearFilter(filterType) {
        this.libraryFilters[filterType] = new Set();
        this.closeFilterDropdown(filterType);
        this.renderLibraryTable();
    }

    capitalizeFirst(str) {
        return str.charAt(0).toUpperCase() + str.slice(1);
    }

    escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    renderLibraryTable() {
        const tbody = document.getElementById('libraryTableBody');
        let contractors = this.contractorsLibrary.getAll();

        // Фильтрация по поиску
        if (this.librarySearchTerm) {
            const term = this.librarySearchTerm.toLowerCase();
            contractors = contractors.filter(c =>
                c.name.toLowerCase().includes(term) ||
                (c.organization && c.organization.toLowerCase().includes(term)) ||
                (c.explanation && c.explanation.toLowerCase().includes(term))
            );
        }

        // Фильтрация по чекбоксам
        for (const [filterType, selectedValues] of Object.entries(this.libraryFilters)) {
            if (selectedValues.size > 0) {
                contractors = contractors.filter(c => {
                    const val = c[filterType] || '(пусто)';
                    return selectedValues.has(val);
                });
            }
        }

        this.updateLibraryStats();

        if (contractors.length === 0) {
            tbody.innerHTML = `
                <tr class="empty-row">
                    <td colspan="4">${this.librarySearchTerm || Object.values(this.libraryFilters).some(s => s.size > 0) ? 'Ничего не найдено' : 'Библиотека пуста'}</td>
                </tr>
            `;
            this.selectedRowId = null;
            return;
        }

        let html = '';
        contractors.forEach(contractor => {
            const isSelected = contractor.id === this.selectedRowId;
            html += `
                <tr data-id="${contractor.id}" ${isSelected ? 'class="selected"' : ''}>
                    <td>${contractor.name}</td>
                    <td>${contractor.organization || '<span style="color: var(--text-tertiary);">—</span>'}</td>
                    <td style="color: var(--text-secondary); font-size: 12px;">${contractor.explanation || '<span style="color: var(--text-tertiary);">—</span>'}</td>
                    <td class="actions-cell">
                        <button class="btn btn-sm btn-secondary edit-contractor" data-id="${contractor.id}" title="Редактировать">
                            <i class="fas fa-edit"></i>
                        </button>
                        <button class="btn btn-sm btn-danger delete-contractor" data-id="${contractor.id}" title="Удалить">
                            <i class="fas fa-trash"></i>
                        </button>
                    </td>
                </tr>
            `;
        });

        tbody.innerHTML = html;

        // Выделение строк по клику
        tbody.querySelectorAll('tr[data-id]').forEach(row => {
            // Одинарный клик — выделение
            row.addEventListener('click', (e) => {
                // Не выделяем если клик по кнопкам
                if (e.target.closest('.btn')) return;
                this.selectRow(row.dataset.id);
            });

            // Двойной клик — редактирование
            row.addEventListener('dblclick', (e) => {
                if (e.target.closest('.btn')) return;
                this.openEditContractorModal(row.dataset.id);
            });
        });

        // Обработчики редактирования
        tbody.querySelectorAll('.edit-contractor').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                this.openEditContractorModal(btn.dataset.id);
            });
        });

        // Обработчики удаления
        tbody.querySelectorAll('.delete-contractor').forEach(btn => {
            btn.addEventListener('click', async (e) => {
                e.stopPropagation();
                const contractor = this.contractorsLibrary.contractors.find(c => c.id === btn.dataset.id);
                if (contractor && confirm(`Удалить "${contractor.name}"?`)) {
                    await this.contractorsLibrary.removeContractor(btn.dataset.id);
                    if (this.selectedRowId === btn.dataset.id) {
                        this.selectedRowId = null;
                    }
                    this.renderLibraryTable();
                    this.updateLibraryStats();
                    this.showNotification('Контрагент удалён', 'success');
                }
            });
        });
    }

    selectRow(id) {
        this.selectedRowId = id;

        // Обновляем визуальное выделение — только строки с data-id
        const tbody = document.getElementById('libraryTableBody');
        tbody.querySelectorAll('tr[data-id]').forEach(row => {
            row.classList.toggle('selected', row.dataset.id === id);
        });
    }

    async performDeleteContractor() {
        const id = document.getElementById('deleteContractorId').value;
        if (id) {
            await this.contractorsLibrary.removeContractor(id);
            this.renderLibraryTable();
            this.updateLibraryStats();
            this.showNotification('Контрагент удалён', 'info');
        }
        document.getElementById('confirmDeleteModal').style.display = 'none';
    }

    // Метод для работы с модальным окном contractorModal
    async saveContractorFromModal() {
        const id = document.getElementById('editContractorId').value;
        const name = document.getElementById('editContractorName').value.trim();
        const organization = document.getElementById('editContractorOrg').value.trim();
        const explanation = document.getElementById('editContractorExplanation').value.trim();

        if (!name) {
            this.showNotification('Введите наименование контрагента', 'error');
            return;
        }

        let result;
        if (id) {
            result = await this.contractorsLibrary.updateContractor(id, { name, organization, explanation });
        } else {
            result = await this.contractorsLibrary.addContractor({ name, organization, explanation });
        }

        if (result.success || result.updated) {
            document.getElementById('contractorModal').style.display = 'none';
            this.renderLibraryTable();
            this.updateSuppliersUI();
            this.showNotification(result.message, 'success');
        } else {
            this.showNotification(result.message || 'Ошибка сохранения', 'error');
        }
    }

    // Обновлённый метод открытия модалки для добавления (использует новое окно contractorModal)
    openAddContractorModal() {
        document.getElementById('contractorModalTitle').textContent = 'Добавить контрагента';
        document.getElementById('editContractorId').value = '';
        document.getElementById('editContractorName').value = '';
        document.getElementById('editContractorOrg').value = '';
        document.getElementById('editContractorExplanation').value = '';
        document.getElementById('contractorModal').style.display = 'flex';

        // Фокус на первое поле
        setTimeout(() => document.getElementById('editContractorName').focus(), 100);
    }

    // Обновлённый метод открытия модалки для редактирования
    openEditContractorModal(id) {
        const contractor = this.contractorsLibrary.contractors.find(c => c.id === id || String(c.id) === String(id));
        if (!contractor) return;

        document.getElementById('contractorModalTitle').textContent = 'Редактировать контрагента';
        document.getElementById('editContractorId').value = contractor.id;
        document.getElementById('editContractorName').value = contractor.name;
        document.getElementById('editContractorOrg').value = contractor.organization || '';
        document.getElementById('editContractorExplanation').value = contractor.explanation || '';
        document.getElementById('contractorModal').style.display = 'flex';

        // Фокус на первое поле
        setTimeout(() => document.getElementById('editContractorName').focus(), 100);
    }

    async importLibraryFromExcel(file) {
        this.showLoading();
        try {
            const result = await this.contractorsLibrary.importFromExcel(file);
            if (result.success) {
                this.renderLibraryTable();
                this.updateSuppliersUI();
                // Показываем расширенное уведомление
                const message = `${result.message}`;
                this.showNotification(message, 'success');
            } else {
                this.showNotification(result.message, 'error');
            }
        } catch (error) {
            console.error('Ошибка импорта библиотеки:', error);
            this.showNotification('Ошибка импорта: ' + error.message, 'error');
        } finally {
            this.hideLoading();
        }
    }

    // ===== ОБРАБОТЧИКИ МОДАЛЬНОГО ОКНА НАСТРОЕК ПАРСЕРА =====
    setupParserSettingsListeners() {
        const modal = document.getElementById('parserSettingsModal');

        // Открытие модального окна по кнопке на странице загрузки
        const openBtn = document.getElementById('openParserSettingsBtn');
        if (openBtn) {
            openBtn.addEventListener('click', () => this.openParserSettings());
        }

        // Закрытие модального окна
        const closeBtn = document.getElementById('closeParserSettingsModal');
        if (closeBtn) {
            closeBtn.addEventListener('click', () => modal.classList.remove('active'));
        }

        document.querySelectorAll('#parserSettingsModal .close-modal').forEach(btn => {
            btn.addEventListener('click', () => modal.classList.remove('active'));
        });

        // Клик вне модального окна
        window.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.classList.remove('active');
            }
        });

        // Добавление нового счёта
        document.getElementById('addMappingBtn').addEventListener('click', () => {
            this.addMappingEntry();
        });

        // Enter в полях формы
        ['newAccountNumber', 'newCompanyName', 'newBankName'].forEach(id => {
            document.getElementById(id).addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    this.addMappingEntry();
                }
            });
        });

        // Синхронизация с сервером
        document.getElementById('syncMappingBtn').addEventListener('click', async () => {
            const btn = document.getElementById('syncMappingBtn');
            btn.disabled = true;
            btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Синхронизация...';
            try {
                const result = await this.parser.syncMappingToServer();
                if (result.success) {
                    this.showNotification(`Синхронизировано: добавлено ${result.added}, обновлено ${result.updated}`, 'success');
                } else {
                    this.showNotification('Ошибка синхронизации: ' + (result.error || 'неизвестная ошибка'), 'error');
                }
            } catch (error) {
                this.showNotification('Ошибка подключения к серверу', 'error');
            } finally {
                btn.disabled = false;
                btn.innerHTML = '<i class="fas fa-sync-alt"></i> Синхронизировать с сервером';
            }
        });

        // Загрузка с сервера
        document.getElementById('reloadMappingBtn').addEventListener('click', async () => {
            const btn = document.getElementById('reloadMappingBtn');
            btn.disabled = true;
            btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Загрузка...';
            try {
                await this.parser.loadMappingFromServer();
                this.renderMappingTable();
                this.loadMappingDatalists();
                this.showNotification('Маппинг загружен с сервера', 'success');
            } catch (error) {
                this.showNotification('Ошибка загрузки с сервера', 'error');
            } finally {
                btn.disabled = false;
                btn.innerHTML = '<i class="fas fa-download"></i> Загрузить с сервера';
            }
        });

        // Поисковый фильтр в модальном окне
        const searchInput = document.getElementById('mappingSearchInput');
        if (searchInput) {
            searchInput.addEventListener('input', () => this.filterMappingTable());
        }
    }

    async addMappingEntry() {
        const accountNumber = document.getElementById('newAccountNumber').value.trim();
        const companyName = document.getElementById('newCompanyName').value.trim();
        const bankName = document.getElementById('newBankName').value.trim();

        if (!accountNumber || !companyName) {
            this.showNotification('Номер счёта и название компании обязательны', 'error');
            return;
        }

        // Обновляем локальный маппинг
        const mapping = { ...this.parser.accountMapping };
        mapping[accountNumber] = { company: companyName, bank: bankName };
        this.parser.saveAccountMapping(mapping);

        // Отправляем на сервер
        try {
            const response = await fetch('/api/account-mapping', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    account_number: accountNumber,
                    company_name: companyName,
                    bank_name: bankName
                })
            });

            if (response.ok) {
                const result = await response.json();
                if (result.success) {
                    this.showNotification(`Счёт ${accountNumber} добавлен`, 'success');
                }
            } else {
                console.warn('Сервер недоступен, сохранено локально');
                this.showNotification('Сохранено локально (сервер недоступен)', 'warning');
            }
        } catch (error) {
            console.warn('Ошибка отправки на сервер:', error);
            this.showNotification('Сохранено локально (сервер недоступен)', 'warning');
        }

        // Очищаем поля
        document.getElementById('newAccountNumber').value = '';
        document.getElementById('newCompanyName').value = '';
        document.getElementById('newBankName').value = '';

        // Обновляем таблицу
        this.renderMappingTable();
        this.loadMappingDatalists();

        // Фокус на поле счёта
        document.getElementById('newAccountNumber').focus();
    }

    renderMappingTable() {
        const tbody = document.getElementById('mappingTableBody');
        const mapping = this.parser.getMappingForUI();

        document.getElementById('mappingCount').textContent = `Записей: ${mapping.length}`;

        if (mapping.length === 0) {
            tbody.innerHTML = '<tr class="empty-row"><td colspan="4">Маппинг пуст</td></tr>';
            return;
        }

        let html = '';
        mapping.forEach(item => {
            html += `
                <tr>
                    <td style="font-family: monospace; font-size: 12px;">${item.account_number}</td>
                    <td>${this.escapeHtml(item.company_name)}</td>
                    <td>${this.escapeHtml(item.bank_name) || '<span style="color: var(--text-tertiary);">—</span>'}</td>
                    <td>
                        <button class="btn btn-xs btn-danger delete-mapping" data-account="${item.account_number}" title="Удалить">
                            <i class="fas fa-trash"></i>
                        </button>
                    </td>
                </tr>
            `;
        });

        tbody.innerHTML = html;

        // Обработчики удаления
        tbody.querySelectorAll('.delete-mapping').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const account = btn.dataset.account;
                if (confirm(`Удалить счёт ${account} из маппинга?`)) {
                    this.deleteMappingEntry(account);
                }
            });
        });
    }

    async deleteMappingEntry(accountNumber) {
        // Удаляем из локального маппинга
        const mapping = { ...this.parser.accountMapping };
        delete mapping[accountNumber];
        this.parser.saveAccountMapping(mapping);

        // Отправляем на сервер (мягкое удаление через API)
        try {
            // Сначала получаем id записи
            const response = await fetch('/api/account-mapping?all=true');
            if (response.ok) {
                const result = await response.json();
                if (result.success && result.data) {
                    const entry = result.data.find(e => e.account_number === accountNumber && e.is_active);
                    if (entry) {
                        await fetch(`/api/account-mapping/${entry.id}`, { method: 'DELETE' });
                    }
                }
            }
        } catch (error) {
            console.warn('Ошибка удаления на сервере:', error);
        }

        // Обновляем таблицу
        this.renderMappingTable();
        this.loadMappingDatalists();
        this.showNotification(`Счёт ${accountNumber} удалён`, 'success');
    }

    async loadMappingDatalists() {
        // Загружаем списки компаний и банков для автодополнения
        try {
            const [companiesRes, banksRes] = await Promise.all([
                fetch('/api/companies'),
                fetch('/api/banks')
            ]);

            if (companiesRes.ok) {
                const companies = await companiesRes.json();
                const datalist = document.getElementById('companiesList');
                if (datalist && companies.data) {
                    datalist.innerHTML = companies.data.map(c => `<option value="${c}">`).join('');
                }
            }

            if (banksRes.ok) {
                const banks = await banksRes.json();
                const datalist = document.getElementById('banksList');
                if (datalist && banks.data) {
                    datalist.innerHTML = banks.data.map(b => `<option value="${b}">`).join('');
                }
            }
        } catch (error) {
            // Fallback: используем локальные данные
            const mapping = this.parser.getMappingForUI();
            const companies = [...new Set(mapping.map(m => m.company_name).filter(Boolean))].sort();
            const banks = [...new Set(mapping.map(m => m.bank_name).filter(Boolean))].sort();

            const companiesDatalist = document.getElementById('companiesList');
            if (companiesDatalist) {
                companiesDatalist.innerHTML = companies.map(c => `<option value="${c}">`).join('');
            }

            const banksDatalist = document.getElementById('banksList');
            if (banksDatalist) {
                banksDatalist.innerHTML = banks.map(b => `<option value="${b}">`).join('');
            }
        }

        // Также заполняем локальными данными для надёжности
        const mapping = this.parser.getMappingForUI();
        const companies = [...new Set(mapping.map(m => m.company_name).filter(Boolean))].sort();
        const banks = [...new Set(mapping.map(m => m.bank_name).filter(Boolean))].sort();

        // Дополняем datalist локальными значениями, если их нет
        const companiesDatalist = document.getElementById('companiesList');
        if (companiesDatalist) {
            const existingOptions = new Set(Array.from(companiesDatalist.options).map(o => o.value));
            companies.forEach(c => {
                if (!existingOptions.has(c)) {
                    companiesDatalist.innerHTML += `<option value="${c}">`;
                }
            });
        }

        const banksDatalist = document.getElementById('banksList');
        if (banksDatalist) {
            const existingOptions = new Set(Array.from(banksDatalist.options).map(o => o.value));
            banks.forEach(b => {
                if (!existingOptions.has(b)) {
                    banksDatalist.innerHTML += `<option value="${b}">`;
                }
            });
        }
    }

    openParserSettings() {
        // Загружаем данные перед открытием
        this.renderMappingTable();
        this.loadMappingDatalists();
        document.getElementById('parserSettingsModal').classList.add('active');
        // Очищаем поисковый фильтр при открытии
        const searchInput = document.getElementById('mappingSearchInput');
        if (searchInput) searchInput.value = '';
    }

    filterMappingTable() {
        const query = document.getElementById('mappingSearchInput').value.toLowerCase().trim();
        const rows = document.querySelectorAll('#mappingTableBody tr:not(.empty-row)');
        let visibleCount = 0;

        rows.forEach(row => {
            const text = row.textContent.toLowerCase();
            if (!query || text.includes(query)) {
                row.style.display = '';
                visibleCount++;
            } else {
                row.style.display = 'none';
            }
        });

        document.getElementById('mappingCount').textContent = 
            `Показано: ${visibleCount} из ${this.parser.getMappingForUI().length}`;
    }
}

// Инициализация приложения после загрузки страницы
document.addEventListener('DOMContentLoaded', () => {
    window.app = new App();
});
