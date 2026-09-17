-- Схема базы данных для хранения истории сверок дебиторской задолженности
-- Финансовый анализатор - Динамика просроченной задолженности

-- Таблица 1: Сводка по дате сверки
CREATE TABLE IF NOT EXISTS swipe_history (
    id SERIAL PRIMARY KEY,
    swipe_date DATE NOT NULL UNIQUE,           -- Дата сверки (уникальная — одна сверка в день)
    total_overdue NUMERIC(20, 2) DEFAULT 0,    -- Общая просроченная задолженность
    total_debt NUMERIC(20, 2) DEFAULT 0,       -- Общая дебиторская задолженность
    filial_count INT DEFAULT 0,                -- Количество филиалов
    counterparty_count INT DEFAULT 0,          -- Количество контрагентов
    legal_dt NUMERIC(20, 2) DEFAULT 0,         -- Судебная задолженность ДТ
    not_recoverable_dt NUMERIC(20, 2) DEFAULT 0, -- Не подлежащая взысканию ДТ
    recoverable_dt NUMERIC(20, 2) DEFAULT 0,   -- Подлежащая взысканию ДТ
    legal_siuat NUMERIC(20, 2) DEFAULT 0,      -- Судебная задолженность СИ УАТ
    not_recoverable_siuat NUMERIC(20, 2) DEFAULT 0, -- Не подлежащая взысканию СИ УАТ
    recoverable_siuat NUMERIC(20, 2) DEFAULT 0, -- Подлежащая взысканию СИ УАТ
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Таблица 2: Данные по филиалам (ДТ) на дату сверки
CREATE TABLE IF NOT EXISTS filial_snapshots (
    id SERIAL PRIMARY KEY,
    swipe_id INT NOT NULL REFERENCES swipe_history(id) ON DELETE CASCADE,
    swipe_date DATE NOT NULL,                   -- Денормализация для удобства запросов
    filial_name VARCHAR(255) NOT NULL,         -- Название филиала (например, "ДТ ТУРУХАНСК")
    overdue_amount NUMERIC(20, 2) DEFAULT 0,   -- Просроченная задолженность
    total_debt_amount NUMERIC(20, 2) DEFAULT 0, -- Общая задолженность
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(swipe_date, filial_name)             -- Один филиал — одна запись на дату
);

-- Таблица 3: Данные по контрагентам на дату сверки (2-й уровень детализации)
CREATE TABLE IF NOT EXISTS counterparty_snapshots (
    id SERIAL PRIMARY KEY,
    swipe_id INT NOT NULL REFERENCES swipe_history(id) ON DELETE CASCADE,
    swipe_date DATE NOT NULL,                   -- Денормализация для удобства запросов
    filial_name VARCHAR(255) NOT NULL,         -- Название филиала (родитель)
    counterparty_name VARCHAR(255) NOT NULL,   -- Название контрагента (например, "РН-Ванкор ООО")
    debt_amount NUMERIC(20, 2) DEFAULT 0,      -- Задолженность контрагента
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(swipe_date, filial_name, counterparty_name) -- Одна запись на комбинацию
);

-- Индексы для ускорения запросов
CREATE INDEX IF NOT EXISTS idx_filial_swipe_date ON filial_snapshots(swipe_date);
CREATE INDEX IF NOT EXISTS idx_filial_name ON filial_snapshots(filial_name);
CREATE INDEX IF NOT EXISTS idx_counterparty_swipe_date ON counterparty_snapshots(swipe_date);
CREATE INDEX IF NOT EXISTS idx_counterparty_filial ON counterparty_snapshots(filial_name);
CREATE INDEX IF NOT EXISTS idx_counterparty_name ON counterparty_snapshots(counterparty_name);
CREATE INDEX IF NOT EXISTS idx_swipe_history_date ON swipe_history(swipe_date);

-- Таблица 4: Маппинг расчётных счетов → компания + банк (настройки парсера выписок)
CREATE TABLE IF NOT EXISTS account_mapping (
    id SERIAL PRIMARY KEY,
    account_number VARCHAR(20) NOT NULL UNIQUE,
    company_name VARCHAR(255) NOT NULL,
    bank_name VARCHAR(100) DEFAULT '',
    is_active BOOLEAN DEFAULT true,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_account_mapping_number ON account_mapping(account_number);
CREATE INDEX IF NOT EXISTS idx_account_mapping_company ON account_mapping(company_name);
CREATE INDEX IF NOT EXISTS idx_account_mapping_active ON account_mapping(is_active);

-- Таблица 5: Библиотека контрагентов (пояснения к платежам)
CREATE TABLE IF NOT EXISTS contractors_library (
    id SERIAL PRIMARY KEY,
    name VARCHAR(500) NOT NULL UNIQUE,
    organization VARCHAR(500) DEFAULT '',
    explanation TEXT DEFAULT '',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_contractors_library_name ON contractors_library(name);

-- ============================================================
-- НОВЫЕ ТАБЛИЦЫ: ГИБКИЕ НАСТРОЙКИ (аналог BankFlow settings)
-- ============================================================

-- Таблица 6: Правила исключения (фильтрация ненужных записей)
CREATE TABLE IF NOT EXISTS exclusion_rules (
    id SERIAL PRIMARY KEY,
    rule_type VARCHAR(20) NOT NULL CHECK (rule_type IN ('purpose', 'counterparty')),
    pattern VARCHAR(500) NOT NULL,
    is_regex BOOLEAN DEFAULT false,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_exclusion_rules_type ON exclusion_rules(rule_type);

-- Таблица 7: Правила категоризации (автоопределение категории по полю)
CREATE TABLE IF NOT EXISTS categorization_rules (
    id SERIAL PRIMARY KEY,
    field VARCHAR(20) NOT NULL CHECK (field IN ('purpose', 'counterparty')),
    pattern VARCHAR(500) NOT NULL,
    display_name VARCHAR(255) NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_categorization_rules_field ON categorization_rules(field);

-- Таблица 8: Синонимы компаний (унификация названий)
CREATE TABLE IF NOT EXISTS company_aliases (
    id SERIAL PRIMARY KEY,
    pattern VARCHAR(500) NOT NULL,
    canonical VARCHAR(255) NOT NULL,
    match_type VARCHAR(20) DEFAULT 'contains' CHECK (match_type IN ('exact', 'contains', 'regex')),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_company_aliases_canonical ON company_aliases(canonical);

-- Таблица 9: Срочные депозиты (длинные) — сохраняются между сессиями
CREATE TABLE IF NOT EXISTS term_deposits (
    id SERIAL PRIMARY KEY,
    account_number VARCHAR(20) NOT NULL,
    amount NUMERIC(20, 2) NOT NULL DEFAULT 0,
    rate NUMERIC(10, 4) NOT NULL DEFAULT 0,
    start_date DATE NOT NULL,
    end_date DATE NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_term_deposits_account ON term_deposits(account_number);
CREATE INDEX IF NOT EXISTS idx_term_deposits_end_date ON term_deposits(end_date);

-- Таблица 10: Контроль лимитов платежей поставщикам
-- Одна строка = один подтверждённый платёж/заявка загрузки.
-- Месяц и период (P1..P4) определяются по дате подтверждения (дате загрузки файла),
-- а НЕ по дате платежа из файла.
CREATE TABLE IF NOT EXISTS payment_limits_payments (
    id SERIAL PRIMARY KEY,
    limit_month DATE NOT NULL,               -- 'YYYY-MM-01' — месяц учёта (по дате подтверждения)
    period SMALLINT NOT NULL CHECK (period BETWEEN 1 AND 4),  -- P1..P4 по дате подтверждения
    stage VARCHAR(10) NOT NULL DEFAULT 'final', -- стадия: 'draft' (черновик) | 'final' (финальный вариант)
    confirmed_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, -- дата подтверждения
    article VARCHAR(255) NOT NULL,           -- Статья движения денежных средств
    detail_limit VARCHAR(255) DEFAULT '',    -- подстатья / детальный лимит (для «Аренда, Коммуналка, Связь»)
    amount NUMERIC(20, 2) NOT NULL DEFAULT 0,-- сумма из файла, рубли
    recipient VARCHAR(500) DEFAULT '',
    request_no VARCHAR(255) DEFAULT '',
    request_date DATE,
    status VARCHAR(255) DEFAULT '',
    organization VARCHAR(255) DEFAULT '',
    subdivision VARCHAR(255) DEFAULT '',
    priority VARCHAR(100) DEFAULT '',
    purpose TEXT DEFAULT '',
    payment_date DATE,
    applicant VARCHAR(255) DEFAULT '',
    payment_type VARCHAR(255) DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_plp_month_period ON payment_limits_payments(limit_month, period);
CREATE INDEX IF NOT EXISTS idx_plp_article ON payment_limits_payments(article);
CREATE INDEX IF NOT EXISTS idx_plp_stage ON payment_limits_payments(limit_month, period, stage);

-- Таблица 11а: Реестры загрузок лимитов — исходный файл + снятые чекбоксами строки по стадиям.
-- Нужна, чтобы варианты (черновик/финал) и финальный файл были доступны с любого компьютера,
-- и обнуляется вместе с платежами по истечении месяца.
CREATE TABLE IF NOT EXISTS payment_limits_registry (
    id SERIAL PRIMARY KEY,
    limit_month DATE NOT NULL,
    period SMALLINT NOT NULL CHECK (period BETWEEN 1 AND 4),
    file_name VARCHAR(500) DEFAULT '',
    sheet_name VARCHAR(255) DEFAULT '',
    file_bytes BYTEA,
    row_count INTEGER DEFAULT 0,
    draft_removed_rows JSONB DEFAULT '[]'::jsonb,
    draft_saved_at TIMESTAMP,
    draft_author VARCHAR(255) DEFAULT '',
    final_removed_rows JSONB DEFAULT '[]'::jsonb,
    final_saved_at TIMESTAMP,
    final_author VARCHAR(255) DEFAULT '',
    final_exported_at TIMESTAMP,          -- когда выгрузили финальный файл (для уведомления)
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (limit_month, period)
);

-- Таблица 11: Настройки лимитов платежей (редактируются в интерфейсе; одна строка id = 1)
CREATE TABLE IF NOT EXISTS payment_limits_config (
    id SMALLINT PRIMARY KEY DEFAULT 1 CHECK (id = 1),
    config JSONB NOT NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Комментарии
COMMENT ON TABLE swipe_history IS 'История сверок дебиторской задолженности — сводка по дате';
COMMENT ON TABLE filial_snapshots IS 'Снимки данных по филиалам (ДТ) на дату сверки';
COMMENT ON TABLE counterparty_snapshots IS 'Снимки данных по контрагентам на дату сверки (2-й уровень)';
COMMENT ON TABLE account_mapping IS 'Маппинг расчётных счетов → компания и банк для парсера банковских выписок';
COMMENT ON TABLE contractors_library IS 'Библиотека контрагентов — типовые пояснения к платежам поставщикам';
COMMENT ON TABLE exclusion_rules IS 'Правила исключения транзакций из отчётов (фильтрация)';
COMMENT ON TABLE categorization_rules IS 'Правила категоризации — автоопределение категории по содержимому поля';
COMMENT ON TABLE company_aliases IS 'Синонимы названий компаний для унификации';
COMMENT ON TABLE payment_limits_payments IS 'Контроль лимитов платежей поставщикам — подтверждённые платежи (месяц/период по дате подтверждения)';
COMMENT ON TABLE payment_limits_config IS 'Настройки лимитов платежей поставщикам (редактируются в интерфейсе, JSONB)';