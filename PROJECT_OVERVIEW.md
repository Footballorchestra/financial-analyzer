# FINANCIAL-ANALYZER — Обзор проекта

> Файл-справочник для быстрого входа в проект. Создан автоматически на основе анализа кода.
> Прочитав этот файл, можно сразу понимать архитектуру, модули, взаимосвязи и ключевые потоки данных,
> не изучая весь код заново.

---

## 1. Что это за проект

**Financial Analyzer («ФинАналитик»)** — веб-приложение для финансового анализа деятельности группы компаний
(«Сервис-Интегратор», «СИ УАТ», «СОИР», «Сервис ЦМ», «Управляющая компания Сервис-Интегратор» и др.).

Основные задачи:

1. **Парсинг банковских выписок** (формат 1С / текстовые выписки) в транзакции.
2. **Сверка дебиторской задолженности (ДЗ)** — сопоставление реестра ДЗ с реестром фактических поступлений,
   расчёт просрочки и сохранение истории сверок.
3. **Распределение налогов по подразделениям (ДДС)** — пропорционально выручке, с формированием Excel-отчёта ДДС.
4. **Сверка расходных операций** (оплаты) с данными из Excel.
5. **Мониторинг остатков по счетам и расчёт процентов по депозитам**.
6. **Сводные таблицы по оплатам поставщикам**.
7. **Библиотека контрагентов** с типовыми пояснениями.
8. **Отчёты/дашборд** по динамике просроченной задолженности (филиалы/контрагенты, сравнение периодов).
9. **Гибкие настройки** — маппинг счетов, правила исключения/категоризации, синонимы компаний.

---

## 2. Архитектура (общая схема)

```
[Браузер: SPA на vanilla JS]
   ├── index.html + styles.css
   ├── app.js (главный контроллер App)
   └── 12 модулей-менеджеров (JS-классы)
          │  fetch() → REST API (JSON / файлы Excel)
          ▼
[Backend: Python Flask 2.3]
   ├── server.py  — HTTP-маршруты, манипуляция Excel (openpyxl)
   ├── allocator.py — движок распределения налогов (CashFlowAllocator)
   └── db.py      — слой доступа к БД (пул соединений psycopg2)
          │  SQL
          ▼
[PostgreSQL]  — schema.sql (9 таблиц)
```

- **Фронтенд** — классический SPA без фреймворка: один `index.html`, переключение «страниц» через скрытие/показ `<div class="page">` в боковом меню.
- **Бэкенд** — Flask, раздаёт статику и предоставляет REST API. Тяжёлые операции (сверка ДЗ, распределение налогов) выполняются на сервере через openpyxl.
- **БД** — PostgreSQL, используется для персистентности: история сверок, настройки, библиотека контрагентов, депозиты.
- **localStorage** — используется для промежуточных/кэшируемых данных (файлы, транзакции, маппинг счетов как fallback, данные «предыдущего дня»).

---

## 3. Технологический стек

| Слой | Технологии |
|------|------------|
| Backend | Python 3.12, Flask 2.3.3, Flask-CORS 4.0.0, Flask-Limiter 3.5.0, openpyxl 3.1.2 |
| БД | PostgreSQL, psycopg2-binary 2.9.9 (ThreadedConnectionPool) |
| Сервер | gunicorn 21.2.0 (systemd-юнит) |
| Frontend | Vanilla JS (ES6-классы), Chart.js (графики), SheetJS (`lib/xlsx.full.min.js`) + ExcelJS (CDN) для Excel, FileSaver (CDN) |
| Прочее | psutil (не используется напрямую в основном коде) |

Внешние CDN (в `index.html`): FileSaver, ExcelJS, Font Awesome.

## 4. Структура файлов

### 4.1. Backend (Python)

| Файл | Строк | Назначение |
|------|------:|------------|
| `server.py` | 1931 | Flask-приложение: все REST-эндпоинты, логика сверки ДЗ поверх Excel (переименование листов, обновление строк, пересчёт итогов), генерация сводных таблиц |
| `db.py` | 2109 | Слой БД: пул соединений, все CRUD-функции для 9 таблиц |
| `allocator.py` | 776 | `CashFlowAllocator` + `allocate_cashflow()`: загрузка Excel «Источник», разделение по подразделениям/месяцам, расчёт долей по выручке, распределение налогов, генерация ДДС Excel |
| `schema.sql` | 143 | SQL-схема БД (9 таблиц + индексы + комментарии) |
| `gunicorn.conf.py` | 61 | Конфигурация gunicorn (workers, таймауты, логи, pid) |
| `financial-analyzer.service` | ~30 | systemd-юнит для запуска под gunicorn |
| `requirements.txt` | 8 | Python-зависимости |
| `.env.example` | ~30 | Пример переменных окружения БД |

### 4.2. Frontend (JS + HTML + CSS)

| Файл | Строк | Назначение |
|------|------:|------------|
| `index.html` | 1738 | Разметка SPA: боковое меню, 10 страниц, модальные окна, подключение скриптов |
| `styles.css` | 3118 | Единая дизайн-система (CSS-переменные, кнопки, таблицы, модалки) |
| `app.js` | 2499 | Класс `App` — точка входа: инициализация, навигация, обработчики событий, UI библиотеки и парсера |
| `storage.js` | 273 | `StorageManager` — хранилище в памяти + localStorage + синхронизация с API |
| `parser.js` | 587 | `BankStatementParser` — парсинг банковских выписок, маппинг счетов, нормализация компаний |
| `receipts-manager.js` | 295 | `ReceiptsManager` — страница «Поступления» (входящие платежи) |
| `expenses-reconciliation.js` | 720 | `ExpensesReconciliationManager` — сверка расходов с Excel |
| `balances-manager.js` | 426 | `BalancesManager` — остатки и проценты по депозитам |
| `debt-reconciliation.js` | 1515 | `DebtReconciliationManager` — сверка ДЗ (ядро «Дебиторки») |
| `contractors-library.js` | 474 | `ContractorsLibrary` — библиотека контрагентов |
| `supplier-payments.js` | 480 | `SupplierPaymentsManager` — сводные таблицы оплат поставщикам |
| `cashflow-allocator.js` | 584 | `CashFlowAllocator` (фронт) — интерфейс распределения налогов/ДДС |
| `reports-manager.js` | 1250 | `ReportsManager` — отчёты/дашборд по истории сверок |
| `settings-manager.js` | 818 | `SettingsManager` — страница настроек (4 вкладки) |

### 4.3. Библиотеки

| Файл | Назначение |
|------|------------|
| `lib/chart.min.js` | Chart.js (графики) |
| `lib/xlsx.full.min.js` | SheetJS — чтение/запись Excel на клиенте |

## 5. Страницы приложения (боковое меню)

Переключение страниц реализовано в `app.js` (`setupNavigation` / `switchPage`), атрибут `data-page` в кнопках меню соответствует `id` блока-страницы.

| # | data-page | Страница | Файл-менеджер | Что делает |
|---|-----------|----------|---------------|------------|
| 1 | `upload` | Загрузка выписок | `parser.js` | Загрузка файлов банковских выписок, парсинг в транзакции |
| 2 | `receipts` | Поступления | `receipts-manager.js` | Просмотр/фильтр/экспорт входящих платежей, загрузка ИНН |
| 3 | `expenses-reconciliation` | Сверка оплат | `expenses-reconciliation.js` | Сверка расходных операций с Excel |
| 4 | `balances` | Остатки | `balances-manager.js` | Остатки по счетам, депозиты, расчёт процентов |
| 5 | `debt` | Дебиторка | `debt-reconciliation.js` | Сверка ДЗ с поступлениями, сохранение истории |
| 6 | `suppliers` | Оплаты поставщикам | `supplier-payments.js` | Сводные таблицы по контрагентам |
| 7 | `cashflow` | ДДС по подразделениям | `cashflow-allocator.js` | Распределение налогов → Excel ДДС |
| 8 | `library` | Библиотека | `contractors-library.js` | CRUD контрагентов, пояснения, импорт/экспорт |
| 9 | `reports` | Отчёты | `reports-manager.js` | Дашборд динамики просроченной задолженности |
| 10 | `settings` | Настройки | `settings-manager.js` | Маппинг счетов, правила, синонимы |

---

## 6. Классы фронтенда и их ответственность

### 6.1. `App` (app.js)
Точка входа. В конструкторе создаёт все менеджеры, в `init()` настраивает навигацию, обработчики, загружает конфиг из БД (`storage.loadConfig()`), инициализирует депозиты и библиотеку контрагентов. Методы-обработчики: `handleFileSelect`, `processFiles`, `performReconciliation`, `performExpensesReconciliation`, `renderLibraryTable`, `renderMappingTable`, `setupParserSettingsListeners` и др.

### 6.2. `StorageManager` (storage.js)
Хранилище в памяти с кэшем: `files`, `statements`, `transactions`, `accounts`, `innData`, `depositData`, `exclusionRules`, `categorizationRules`, `companyAliases`. Ключевые методы:
- `loadConfig()` — грузит конфиг через `GET /api/config` (маппинг счетов + правила).
- `initDeposits()` / `syncDepositsToServer()` — депозиты из БД (`/api/term-deposits`) с fallback на localStorage. При синхронизации отправляются все счета локального состояния, в т.ч. с пустым списком, иначе удаление последнего срочного депозита не сохранилось бы в БД.
- Вспомогательные: `formatCurrency`, `parseDate`, `calculateInterest`, `getDaysBetween`.

### 6.3. `BankStatementParser` (parser.js)
Парсит текстовые банковские выписки (кодировка Windows-1251). Определяет счёт/банк/компанию по маппингу, направление транзакции, нормализует названия через `company_aliases` из БД. Маппинг счетов хранится в localStorage с большим fallback-списком (хардкод ~50 счетов компаний группы).

### 6.4. `DebtReconciliationManager` (debt-reconciliation.js) — ЯДРО
Отвечает за сверку дебиторской задолженности:
- Хранит `debtData` (реестр ДЗ), `receiptsData` (реестр поступлений), `processedDocuments`, `TARGET_CONTRAGENTS` (целевые контрагенты для сверки).
- `reconcile()` — проход по документным строкам, сопоставление с картой поступлений по имени документа, определение просрочки относительно `currentDate`, обновление строк.
- `collectSubdivisionData()` — сбор сумм просрочки по филиалам (строки начинаются с `«ДТ »`).
- `updateDocumentRow()` — записывает сумму/дни просрочки в нужные колонки.
- Сохранение «данных предыдущего дня» в localStorage, отправка на сервер (`POST /save-excel`).
- **Отправка по почте**: `sendToEmail()` отправляет файл на сервер с `mode=email`, получает данные таблиц из заголовка `X-Email-Data` и генерирует 3 файла `.eml` (MIME) с HTML-таблицами и вложением. Письмо 1 — 4 таблицы (Свод ДТ, Динамика по подразделениям, Свод СИ УАТ, таблица из «Лист1» A5:D8 файла СИ УАТ), письмо 2 — 2 таблицы (Свод ДТ, Свод СИ УАТ), письмо 3 — 1 таблица (Свод ДТ). Получатели/тема/подпись заданы в `_getEmailRecipients()`/`_getEmailSignature()`, отправитель — `minenkov.a@s-int.ru`.

### 6.5. `ReportsManager` (reports-manager.js)
Дашборд с 3 вкладками: `overview` (сводка + графики), `details` (таблица детализации по филиалам), `comparison` (сравнение филиалов). Загружает данные с API (`/api/summary`, `/api/swipe-dates`, `/api/swipe-raw`, `/api/filial-trend`, `/api/counterparty-trend`, `/api/filial-list`, `/api/counterparty-list`). Поддерживает детализацию по периоду (день/декада/месяц/квартал/год).

> ⚠️ Важная особенность: в `reports-manager.js` захардкожен внешний адрес API `this.apiBase = 'http://31.130.155.16:5000'`, тогда как остальные модули используют относительные URL (`fetch('/api/...')`).

### 6.6. `SettingsManager` (settings-manager.js)
4 вкладки: **Счета** (account_mapping), **Исключения** (exclusion_rules), **Категоризация** (categorization_rules), **Синонимы** (company_aliases). CRUD через API, поиск, модальное редактирование, импорт счетов из Excel.

### 6.7. Прочие менеджеры
- `ReceiptsManager` — таблица поступлений, ИНН-справочник, экспорт Excel.
- `ExpensesReconciliationManager` — сопоставление расходов с Excel (поиск совпадений, комиссии банка).
- `BalancesManager` — остатки, депозиты (срочные + до востребования), расчёт процентов.
- `SupplierPaymentsManager` — pivot-таблицы оплат по контрагентам, подтягивание пояснений из библиотеки.
- `ContractorsLibrary` — CRUD контрагентов (название/организация/пояснение), импорт/экспорт Excel.
- `CashFlowAllocator` (фронт) — выбор периода/стратегии, предпросмотр, отправка `/api/allocate-cashflow`, скачивание Excel.

## 7. Backend: REST API (server.py)

### 7.1. Статика
| Метод | Маршрут | Назначение |
|-------|---------|------------|
| GET | `/` | `index.html` |
| GET | `/lib/<filename>` | файлы из `lib/` |
| GET | `/<filename>` | статические файлы из корня |

### 7.2. Health
| GET | `/api/health` | Проверка сервера и доступности БД (`db.check_health()`) |

### 7.3. Сверка ДЗ (тяжёлые операции)
| Метод | Маршрут | Rate limit | Назначение |
|-------|---------|------------|------------|
| POST | `/save-excel` | 5/мин | Принимает Excel + JSON `updatedDocuments`, обновляет строки просрочки, пересчитывает итоги, возвращает исправленный Excel. При `mode=email` скрывает вкладку «Сводные таблицы» и возвращает данные таблиц письма в заголовке `X-Email-Data` (Base64 JSON) |
| GET/POST | `/api/previous-day-data` | — | Загрузка/сохранение данных «предыдущего дня» для сравнения |
| POST | `/api/allocate-cashflow` | 3/мин | Распределение налогов (возвращает Excel + заголовок `X-Allocation-Summary`) |
| POST | `/save-suppliers` | 5/мин | Сводные таблицы оплат поставщикам |

### 7.4. История сверок (отчёты)
| Метод | Маршрут | Назначение |
|-------|---------|------------|
| POST | `/api/save-swipe-data` | Сохранить снимок сверки в БД |
| POST | `/api/delete-swipe` | Удалить сверку по дате |
| GET | `/api/swipe-dates` | Список дат сверок |
| GET | `/api/swipe-raw` | Сырые данные сверок за период (фильтр по филиалу) |
| GET | `/api/filial-trend` | Динамика по филиалу |
| GET | `/api/counterparty-trend` | Динамика по контрагенту |
| GET | `/api/filial-list` | Список филиалов |
| GET | `/api/counterparty-list` | Список контрагентов (фильтр по филиалу) |
| GET | `/api/summary` | Сводная статистика за период |

### 7.5. Маппинг счетов (account_mapping)
| Метод | Маршрут | Назначение |
|-------|---------|------------|
| GET/POST | `/api/account-mapping` | Список / добавление |
| PUT/DELETE | `/api/account-mapping/<id>` | Обновление / мягкое удаление |
| PUT | `/api/account-mapping/<id>/toggle` | Вкл/выкл счёт |
| POST | `/api/account-mapping/sync` | Синхронизация из парсера |
| GET | `/api/companies` | Уникальные компании (автодополнение) |
| GET | `/api/banks` | Уникальные банки |

### 7.6. Библиотека контрагентов
| Метод | Маршрут | Назначение |
|-------|---------|------------|
| GET/POST | `/api/contractors` | Список / добавление |
| PUT/DELETE | `/api/contractors/<id>` | Обновление / удаление |
| GET | `/api/contractors/find` | Поиск по имени |
| POST | `/api/contractors/import` | Пакетный импорт |
| POST | `/api/contractors/clear` | Очистить всё |
| GET | `/api/contractors/stats` | Статистика |

### 7.7. Настройки (правила)
| Маршрут | Назначение |
|---------|------------|
| `GET /api/config` | Весь конфиг (маппинг + все правила) |
| `GET/POST /api/exclusion-rules`, `PUT/DELETE /api/exclusion-rules/<id>` | Правила исключения |
| `GET/POST /api/categorization-rules`, `PUT/DELETE /api/categorization-rules/<id>` | Правила категоризации |
| `GET/POST /api/company-aliases`, `PUT/DELETE /api/company-aliases/<id>` | Синонимы компаний |

### 7.8. Срочные депозиты
| GET | `/api/term-deposits` | Получить депозиты (сгруппированы по счёту) |
| POST | `/api/term-deposits/sync` | Синхронизация: по каждому счёту из payload строки удаляются и вставляются заново. Счёт с пустым массивом = удалить все его срочные депозиты, поэтому фронтенд шлёт **все** счета локального состояния, включая очищенные. |

## 8. База данных (schema.sql — 9 таблиц)

| # | Таблица | Назначение | Ключевые поля |
|---|---------|------------|---------------|
| 1 | `swipe_history` | Сводка по дате сверки ДЗ | `swipe_date` (UNIQUE), `total_overdue`, `total_debt`, `filial_count`, `counterparty_count`, поля ДТ/СИ УАТ (судебная/не подлежащая/подлежащая взысканию) |
| 2 | `filial_snapshots` | Снимки по филиалам на дату | `swipe_id` FK, `swipe_date`, `filial_name`, `overdue_amount`, `total_debt_amount` |
| 3 | `counterparty_snapshots` | Снимки по контрагентам (2-й уровень) | `swipe_id` FK, `filial_name`, `counterparty_name`, `debt_amount` |
| 4 | `account_mapping` | Маппинг счёт → компания+банк | `account_number` (UNIQUE), `company_name`, `bank_name`, `is_active` |
| 5 | `contractors_library` | Библиотека контрагентов | `name` (UNIQUE), `organization`, `explanation` |
| 6 | `exclusion_rules` | Правила исключения транзакций | `rule_type` (purpose/counterparty), `pattern`, `is_regex` |
| 7 | `categorization_rules` | Правила категоризации | `field` (purpose/counterparty), `pattern`, `display_name` |
| 8 | `company_aliases` | Синонимы компаний | `pattern`, `canonical`, `match_type` (exact/contains/regex) |
| 9 | `term_deposits` | Срочные депозиты | `account_number`, `amount`, `rate`, `start_date`, `end_date` |

Связи: `filial_snapshots` и `counterparty_snapshots` ссылаются на `swipe_history.id` (ON DELETE CASCADE), денормализованы полем `swipe_date` для удобства запросов.

---

## 9. Ключевые бизнес-потоки

### 9.1. Загрузка и парсинг банковских выписок
1. Пользователь загружает файлы выписок (страница «Загрузка выписок»).
2. `BankStatementParser.processFiles()` читает файлы (Windows-1251), `parseStatement()` разбирает строки, определяет счёт/банк/компанию по `accountMapping`, направление транзакции, нормализует компании через `company_aliases`.
3. Результат (statements, transactions, accounts) сохраняется в `StorageManager`.
4. Данные отображаются на страницах «Поступления», «Сверка оплат», «Остатки».

### 9.2. Сверка дебиторской задолженности (страница «Дебиторка»)
1. Загружаются два файла: **Реестр ДЗ** и **Реестр поступлений с датами**.
2. `DebtReconciliationManager.reconcile()` строит карту документов из поступлений (`documentName → expectedDate`), проходит по документным строкам реестра ДЗ, для целевых контрагентов находит дату, рассчитывает просрочку (сегодня − ожидаемая дата) и обновляет колонки просрочки/дней/интервалов.
3. `collectSubdivisionData()` собирает суммы по филиалам (строки `«ДТ …»`).
4. Отправка на сервер `POST /save-excel` (файл + `updatedDocuments`), сервер переименовывает листы, пересчитывает итоги (`recalc_totals`, `update_top_table`) и возвращает исправленный Excel.
5. Снимок сохраняется в БД через `POST /api/save-swipe-data` (или `/api/previous-day-data`).
6. **Отправка по почте** (кнопка «Отправить по почте»): повторно вызывает `/save-excel` с `mode=email`; сервер дополнительно извлекает таблицу из вкладки «Лист1» (диапазон A5:D8) файла СИ УАТ (`extract_siuat_sheet1_table`) и возвращает все данные таблиц в заголовке `X-Email-Data`; фронтенд генерирует 3 `.eml` с HTML-таблицами и вложением (письмо 1 — 4 таблицы, письмо 2 — 2, письмо 3 — 1).

### 9.3. Распределение налогов / ДДС (страница «ДДС по подразделениям»)
1. Пользователь загружает Excel с листом **«Источник»**, задаёт период/стратегию (`revenue` или равные доли) и статьи-исключения.
2. Фронт `CashFlowAllocator` формирует настройки и отправляет `POST /api/allocate-cashflow`.
3. Сервер вызывает `allocator.allocate_cashflow()`: `load_excel()` читает «Источник», `separate_data()` раскладывает по подразделениям/месяцам, `calculate_shares()` считает доли по выручке, `allocate_taxes()` распределяет налоги (НДФЛ/ФОТ/НДС/прибыль) пропорционально долям, `generate_excel()` формирует ДДС (поступления/расходы/остаток по каждому филиалу).
4. Возвращается Excel + метаданные в заголовке `X-Allocation-Summary`.

### 9.4. Отчёты/дашборд (страница «Отчёты»)
1. Пользователь выбирает период/филиал/контрагента → `buildDashboard()`.
2. Параллельно грузятся сводка, история сверок, тренды, сырые данные с API.
3. Рендерятся графики (Chart.js) и таблицы по 3 вкладкам.

### 9.5. Настройки
CRUD-операции над `account_mapping`, `exclusion_rules`, `categorization_rules`, `company_aliases` через соответствующие API; данные кэшируются в `StorageManager` и используются парсером выписок.

## 10. Взаимосвязи модулей (что кого вызывает)

```
index.html
  └── App (app.js)
        ├── StorageManager (storage.js)      ← данные + кэш, sync с API
        ├── BankStatementParser (parser.js)  ← использует storage.getAccounts()/getCompanyAliases()
        ├── ReceiptsManager                   ← storage.transactions
        ├── ExpensesReconciliationManager     ← storage.transactions
        ├── BalancesManager                   ← storage.accounts/depositData
        ├── DebtReconciliationManager         ← server.py /save-excel, /api/save-swipe-data
        ├── SupplierPaymentsManager           ← ContractorsLibrary (пояснения)
        ├── CashFlowAllocator (фронт)         ← server.py /api/allocate-cashflow
        ├── ReportsManager                    ← server.py /api/* (история сверок)
        ├── SettingsManager                   ← server.py /api/config, /api/*-rules, /api/account-mapping
        └── ContractorsLibrary                ← server.py /api/contractors
```

Backend-цепочки:
```
server.py  →  db.py  →  PostgreSQL
server.py  →  allocator.py (openpyxl) → Excel
server.py  →  openpyxl (сверка ДЗ, сводные таблицы)
```

---

## 11. Конфигурация и запуск

- **Переменные окружения** (`db.py` / `.env.example`): `DB_HOST`, `DB_PORT`, `DB_NAME`, `DB_USER`, `DB_PASSWORD`, `DB_CONNECT_TIMEOUT`, `DB_POOL_MIN`, `DB_POOL_MAX`, `GUNICORN_WORKERS`.
- **Запуск вручную**: `python server.py` (порт 5000, host 0.0.0.0).
- **Продакшн**: gunicorn через systemd (`financial-analyzer.service`, `gunicorn.conf.py`). Логи в `/var/log/financial-analyzer/` (ротация 10 МБ × 5).
- **Лимит размера файла**: 100 МБ (`MAX_CONTENT_LENGTH`).
- **Rate limiting**: Flask-Limiter, storage `memory://`, default 60/мин; тяжёлые эндпоинты — 3–5/мин.

---

## 12. Важные замечания и «подводные камни»

1. **Захардкоженный адрес API в `reports-manager.js`**: `this.apiBase = 'http://31.130.155.16:5000'`. Остальной код использует относительные URL. При переносе/настройке это надо выравнивать.
2. **Пароль БД хранится только в `.env`** (файл в `.gitignore`, шаблон — `.env.example`). `db.py` читает `DB_PASSWORD` из переменных окружения и сам подгружает `.env` (без внешних зависимостей, приоритет у переменных из systemd/оболочки). `financial-analyzer.service` подключает `.env` через `EnvironmentFile`. В коде и в юните пароля нет.
3. **Кодировка**: `JSON_AS_ASCII = False`, `JSON_SORT_KEYS = False`; CORS с `expose_headers=['X-Filial-Data', 'X-Allocation-Summary']` — эти заголовки используются фронтендом для передачи метаданных.
4. **Сверка ДЗ на сервере** переименовывает листы Excel: `Свод ДЗ`, `Свод ДЗ СИ УАТ`, `Сводные таблицы`; лишние листы удаляются. Логика привязана к именам листов (`Свод ДЗ СИ УАТ`, `Сводные таблицы`) и префиксу `«ДТ »` в первом столбце.
5. **Иерархия ДЗ** строится по отступам (indent) ячеек, а не по ключевым словам — функции `build_hierarchy`, `recalc_totals`, `find_structure` в `server.py`.
6. **localStorage** активно используется для данных «предыдущего дня» (`previousDayDebt_manual`) и маппинга счетов (`accountMapping`) как fallback.
7. **Файл `test_dds_output.xlsx`** — тестовый выходной файл ДДС (пример результата).
8. Часть зависимостей указана, но может не использоваться напрямую в основном коде (например, `psutil`).

---

## 13. Быстрый старт для новых сессий

1. Прочитать этот файл.
2. Для правок бэкенда смотреть: `server.py` (маршруты), `db.py` (SQL), `allocator.py` (распределение налогов), `schema.sql` (схема).
3. Для правок фронтенда смотреть: `index.html` (разметка), `app.js` (общая логика/навигация), конкретный `*-manager.js` для страницы.
4. Проверка работоспособности: `GET /api/health`; запуск `python server.py` (нужен PostgreSQL и `venv`).





