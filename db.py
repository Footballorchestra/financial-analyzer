# db.py — Модуль работы с PostgreSQL для хранения истории сверок
import psycopg2
import psycopg2.extras
import psycopg2.pool
from datetime import date, datetime
from decimal import Decimal
import json
import logging
import time
import os

# ============================================================
# ЛОГГИРОВАНИЕ
# ============================================================
logger = logging.getLogger('financial_analyzer.db')

# ============================================================
# НАСТРОЙКИ ПОДКЛЮЧЕНИЯ (из переменных окружения с fallback)
# ============================================================
DB_CONFIG = {
    'host': os.environ.get('DB_HOST', '127.0.0.1'),
    'port': int(os.environ.get('DB_PORT', 5432)),
    'database': os.environ.get('DB_NAME', 'financial_analyzer'),
    'user': os.environ.get('DB_USER', 'postgres'),
    'password': os.environ.get('DB_PASSWORD', 'Kapapa661109'),
    'connect_timeout': int(os.environ.get('DB_CONNECT_TIMEOUT', 5)),
}

# ============================================================
# ПУЛ СОЕДИНЕНИЙ (один на весь процесс)
# ============================================================
_pool = None

def _init_pool():
    """Инициализирует пул соединений (вызывается один раз)."""
    global _pool
    if _pool is None:
        try:
            _pool = psycopg2.pool.ThreadedConnectionPool(
                minconn=int(os.environ.get('DB_POOL_MIN', 2)),
                maxconn=int(os.environ.get('DB_POOL_MAX', 10)),
                **DB_CONFIG
            )
            logger.info(f"Пул соединений БД создан: min={_pool.minconn}, max={_pool.maxconn}")
        except Exception as e:
            logger.critical(f"Не удалось создать пул соединений БД: {e}")
            raise


def get_connection():
    """
    Возвращает соединение из пула с retry-логикой.

    При недоступности БД делает до 3 повторных попыток
    с экспоненциальной задержкой (1с, 2с, 4с).
    """
    _init_pool()

    last_error = None
    for attempt in range(3):
        try:
            conn = _pool.getconn()
            # Проверяем, живо ли соединение
            try:
                cur = conn.cursor()
                cur.execute("SELECT 1")
                cur.close()
            except Exception:
                # Соединение умерло — закрываем и пробуем снова
                try:
                    _pool.putconn(conn, close=True)
                except Exception:
                    pass
                raise

            return conn
        except psycopg2.OperationalError as e:
            last_error = e
            if attempt < 2:
                delay = 2 ** attempt  # 1с, 2с, 4с
                logger.warning(
                    f"БД недоступна (попытка {attempt + 1}/3), "
                    f"повтор через {delay}с: {e}"
                )
                time.sleep(delay)
            else:
                logger.error(f"БД недоступна после 3 попыток: {e}")
                raise

    raise last_error or RuntimeError("Не удалось подключиться к БД")


def return_connection(conn, close=False):
    """Возвращает соединение в пул."""
    if _pool and conn:
        try:
            _pool.putconn(conn, close=close)
        except Exception as e:
            logger.warning(f"Ошибка возврата соединения в пул: {e}")


def close_all_connections():
    """Закрывает все соединения в пуле (вызывается при остановке)."""
    global _pool
    if _pool:
        try:
            _pool.closeall()
            logger.info("Пул соединений БД закрыт")
        except Exception as e:
            logger.error(f"Ошибка закрытия пула: {e}")
        _pool = None


# ============================================================
# ПОМОЩНИК ДЛЯ ВЫПОЛНЕНИЯ ФУНКЦИЙ С АВТО-ПОДКЛЮЧЕНИЕМ
# ============================================================

def with_connection(func):
    """
    Декоратор: автоматически получает соединение из пула,
    передаёт его первым аргументом, и возвращает в пул после.
    Также обрабатывает ошибки и делает rollback.
    """
    def wrapper(*args, **kwargs):
        conn = None
        try:
            conn = get_connection()
            return func(conn, *args, **kwargs)
        except psycopg2.OperationalError as e:
            logger.error(f"Ошибка БД в {func.__name__}: {e}")
            return _error_result(func, str(e))
        except Exception as e:
            logger.error(f"Ошибка в {func.__name__}: {e}", exc_info=True)
            if conn:
                try:
                    conn.rollback()
                except Exception:
                    pass
            return _error_result(func, str(e))
        finally:
            if conn:
                return_connection(conn)
    return wrapper


def _error_result(func, error_msg):
    """Возвращает результат ошибки в зависимости от типа функции."""
    name = func.__name__
    if name.startswith('save_') or name.startswith('add_') or name.startswith('insert_'):
        return {'success': False, 'error': error_msg}
    if name.startswith('delete_') or name.startswith('clear_'):
        return {'success': False, 'error': error_msg}
    if name.startswith('get_') or name.startswith('find_') or name.startswith('load_'):
        if 'list' in name:
            return []
        if 'stats' in name or 'summary' in name:
            return {}
        return {} if 'data' in name or 'config' in name else None
    return {'success': False, 'error': error_msg}


# ============================================================
# ФУНКЦИЯ ИСПРАВЛЕНИЯ КОДИРОВКИ
# ============================================================

def fix_encoding(text):
    """
    Исправляет двойное кодирование UTF-8 в строках из БД.

    Проблема: если кириллица была сохранена как байты UTF-8, а прочитана как latin1,
    то "ДТ" превращается в "Ð\x94Ð¢". Эта функция исправляет такие строки.

    Args:
        text: str или None — текст для исправления

    Returns:
        str — исправленный текст или оригинал, если исправление невозможно
    """
    if text is None:
        return None
    if not isinstance(text, str):
        return text
    try:
        return text.encode('latin1').decode('utf-8')
    except (UnicodeEncodeError, UnicodeDecodeError, AttributeError):
        return text


# ============================================================
# СОХРАНЕНИЕ ДАННЫХ СВЕРКИ
# ============================================================

def save_swipe_data(swipe_date, filial_data, counterparty_data,
                    total_debt=0, total_overdue=0,
                    summary_dt=None, summary_siuat=None):
    """
    Сохраняет данные сверки в БД.

    Args:
        swipe_date: str — дата сверки (YYYY-MM-DD)
        filial_data: dict — {filial_name: overdue_amount}
        counterparty_data: dict — {(filial, counterparty): debt_amount}
        total_debt: float — общая ДЗ
        total_overdue: float — общая ПДЗ
        summary_dt: dict — сводные данные ДТ {legal, notRecoverable, recoverable}
        summary_siuat: dict — сводные данные СИ УАТ {legal, notRecoverable, recoverable}
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        cur.execute("""
            INSERT INTO swipe_history (swipe_date, total_overdue, total_debt,
                filial_count, counterparty_count, updated_at,
                legal_dt, not_recoverable_dt, recoverable_dt,
                legal_siuat, not_recoverable_siuat, recoverable_siuat)
            VALUES (%s, %s, %s, %s, %s, NOW(), %s, %s, %s, %s, %s, %s)
            ON CONFLICT (swipe_date) DO UPDATE SET
                total_overdue = EXCLUDED.total_overdue,
                total_debt = EXCLUDED.total_debt,
                filial_count = EXCLUDED.filial_count,
                counterparty_count = EXCLUDED.counterparty_count,
                updated_at = NOW(),
                legal_dt = EXCLUDED.legal_dt,
                not_recoverable_dt = EXCLUDED.not_recoverable_dt,
                recoverable_dt = EXCLUDED.recoverable_dt,
                legal_siuat = EXCLUDED.legal_siuat,
                not_recoverable_siuat = EXCLUDED.not_recoverable_siuat,
                recoverable_siuat = EXCLUDED.recoverable_siuat
            RETURNING id
        """, (
            swipe_date,
            total_overdue,
            total_debt,
            len(filial_data),
            len(counterparty_data),
            summary_dt.get('legal', 0) if summary_dt else 0,
            summary_dt.get('notRecoverable', 0) if summary_dt else 0,
            summary_dt.get('recoverable', 0) if summary_dt else 0,
            summary_siuat.get('legal', 0) if summary_siuat else 0,
            summary_siuat.get('notRecoverable', 0) if summary_siuat else 0,
            summary_siuat.get('recoverable', 0) if summary_siuat else 0
        ))

        swipe_row = cur.fetchone()
        swipe_id = swipe_row[0] if swipe_row else None

        if not swipe_id:
            raise Exception("Не удалось получить ID сверки")

        conn.commit()

        # 2. Сохраняем данные по филиалам
        filial_rows = []
        for filial_name, overdue in filial_data.items():
            fixed_filial = fix_encoding(filial_name)
            filial_rows.append((swipe_id, swipe_date, fixed_filial, float(overdue), 0))

        if filial_rows:
            psycopg2.extras.execute_values(
                cur,
                """
                INSERT INTO filial_snapshots (swipe_id, swipe_date, filial_name,
                    overdue_amount, total_debt_amount)
                VALUES %s
                ON CONFLICT (swipe_date, filial_name) DO UPDATE SET
                    swipe_id = EXCLUDED.swipe_id,
                    overdue_amount = EXCLUDED.overdue_amount,
                    total_debt_amount = EXCLUDED.total_debt_amount
                """,
                filial_rows
            )
        conn.commit()

        # 3. Сохраняем данные по контрагентам
        cp_rows = []
        for key, debt in counterparty_data.items():
            # Ключ может быть кортежем (filial_name, cp_name) или строкой "filial||cp"
            if isinstance(key, tuple):
                filial_name, cp_name = key
            else:
                parts = str(key).split('||', 1)
                filial_name = parts[0] if len(parts) > 0 else ''
                cp_name = parts[1] if len(parts) > 1 else ''
            fixed_filial = fix_encoding(filial_name)
            fixed_cp = fix_encoding(cp_name)
            cp_rows.append((swipe_id, swipe_date, fixed_filial, fixed_cp, float(debt)))

        if cp_rows:
            psycopg2.extras.execute_values(
                cur,
                """
                INSERT INTO counterparty_snapshots (swipe_id, swipe_date,
                    filial_name, counterparty_name, debt_amount)
                VALUES %s
                ON CONFLICT (swipe_date, filial_name, counterparty_name) DO UPDATE SET
                    swipe_id = EXCLUDED.swipe_id,
                    debt_amount = EXCLUDED.debt_amount
                """,
                cp_rows
            )
        conn.commit()

        cur.close()

        logger.info(
            f"Данные сверки за {swipe_date} сохранены: "
            f"{len(filial_data)} филиалов, {len(counterparty_data)} контрагентов"
        )

        return {'success': True, 'swipe_id': swipe_id}

    except Exception as e:
        logger.error(f"Ошибка сохранения данных сверки: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def save_previous_day_data(swipe_date, filial_data, summary_dt=None, summary_siuat=None):
    """
    Сохраняет данные по филиалам на указанную дату (для использования на следующий день).

    Args:
        swipe_date: str — дата (YYYY-MM-DD)
        filial_data: dict — {filial_name: overdue_amount}
        summary_dt: dict — сводные данные ДТ
        summary_siuat: dict — сводные данные СИ УАТ

    Returns:
        dict — {'success': True, 'count': N} или {'success': False, 'error': '...'}
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        total_debt = summary_dt.get('totalDebt', 0) if summary_dt else 0
        total_overdue = summary_dt.get('totalOverdue', 0) if summary_dt else 0

        cur.execute("""
            INSERT INTO swipe_history (swipe_date, total_overdue, total_debt,
                filial_count, counterparty_count, updated_at,
                legal_dt, not_recoverable_dt, recoverable_dt,
                legal_siuat, not_recoverable_siuat, recoverable_siuat)
            VALUES (%s, %s, %s, %s, %s, NOW(), %s, %s, %s, %s, %s, %s)
            ON CONFLICT (swipe_date) DO UPDATE SET
                total_overdue = EXCLUDED.total_overdue,
                total_debt = EXCLUDED.total_debt,
                filial_count = EXCLUDED.filial_count,
                updated_at = NOW(),
                legal_dt = EXCLUDED.legal_dt,
                not_recoverable_dt = EXCLUDED.not_recoverable_dt,
                recoverable_dt = EXCLUDED.recoverable_dt,
                legal_siuat = EXCLUDED.legal_siuat,
                not_recoverable_siuat = EXCLUDED.not_recoverable_siuat,
                recoverable_siuat = EXCLUDED.recoverable_siuat
            RETURNING id
        """, (
            swipe_date,
            total_overdue,
            total_debt,
            len(filial_data),
            0,
            summary_dt.get('legal', 0) if summary_dt else 0,
            summary_dt.get('notRecoverable', 0) if summary_dt else 0,
            summary_dt.get('recoverable', 0) if summary_dt else 0,
            summary_siuat.get('legal', 0) if summary_siuat else 0,
            summary_siuat.get('notRecoverable', 0) if summary_siuat else 0,
            summary_siuat.get('recoverable', 0) if summary_siuat else 0
        ))

        swipe_row = cur.fetchone()
        swipe_id = swipe_row[0] if swipe_row else None
        if not swipe_id:
            raise Exception("Не удалось получить ID записи swipe_history")

        conn.commit()

        # Сохраняем данные по филиалам (swipe_id обязателен для filial_snapshots)
        filial_rows = []
        for filial_name, overdue in filial_data.items():
            fixed_filial = fix_encoding(filial_name)
            filial_rows.append((swipe_id, swipe_date, fixed_filial, float(overdue), 0))

        if filial_rows:
            psycopg2.extras.execute_values(
                cur,
                """
                INSERT INTO filial_snapshots (swipe_id, swipe_date, filial_name,
                    overdue_amount, total_debt_amount)
                VALUES %s
                ON CONFLICT (swipe_date, filial_name) DO UPDATE SET
                    swipe_id = EXCLUDED.swipe_id,
                    overdue_amount = EXCLUDED.overdue_amount,
                    total_debt_amount = EXCLUDED.total_debt_amount
                """,
                filial_rows
            )
        conn.commit()
        cur.close()

        logger.info(f"Данные за {swipe_date} сохранены: {len(filial_data)} филиалов")
        return {'success': True, 'count': len(filial_data)}

    except Exception as e:
        logger.error(f"Ошибка сохранения данных за {swipe_date}: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def get_last_available_swipe_date(before_date):
    """
    Находит последнюю доступную дату с данными по филиалам перед указанной датой.

    Используется для обработки выходных дней: если запрошена дата понедельника,
    а данных за него нет, функция вернет дату последней сверки (например, пятницы).

    Args:
        before_date: str — дата (YYYY-MM-DD), перед которой ищем

    Returns:
        str или None — дата последней доступной сверки в формате YYYY-MM-DD, или None если не найдено
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        cur.execute("""
            SELECT DISTINCT swipe_date
            FROM filial_snapshots
            WHERE swipe_date < %s
            ORDER BY swipe_date DESC
            LIMIT 1
        """, (before_date,))

        result = cur.fetchone()
        cur.close()

        if result and result[0]:
            found_date = result[0]
            if isinstance(found_date, (date, datetime)):
                found_date = found_date.isoformat()
            logger.info(f"Найдена последняя доступная дата: {found_date} (перед {before_date})")
            return found_date
        else:
            logger.warning(f"Не найдено данных перед датой {before_date}")
            return None

    except Exception as e:
        logger.error(f"Ошибка поиска последней доступной даты: {e}")
        return None
    finally:
        if conn:
            return_connection(conn)


def get_previous_day_data(swipe_date):
    """
    Получает данные по филиалам за указанную дату.

    Если данных за запрошенную дату нет, автоматически ищет
    последнюю доступную дату через get_last_available_swipe_date().

    Args:
        swipe_date: str — дата (YYYY-MM-DD)

    Returns:
        dict — {filial_name: overdue_amount, ...} или пустой dict если данных нет
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        cur.execute("""
            SELECT filial_name, overdue_amount
            FROM filial_snapshots
            WHERE swipe_date = %s
            ORDER BY filial_name ASC
        """, (swipe_date,))

        rows = cur.fetchall()

        if not rows:
            logger.warning(f"Данные за {swipe_date} не найдены, ищем последнюю доступную дату...")
            last_date = get_last_available_swipe_date(swipe_date)

            if last_date:
                logger.info(f"Повторный запрос данных за {last_date}")
                cur.execute("""
                    SELECT filial_name, overdue_amount
                    FROM filial_snapshots
                    WHERE swipe_date = %s
                    ORDER BY filial_name ASC
                """, (last_date,))
                rows = cur.fetchall()
                swipe_date = last_date
            else:
                cur.close()
                logger.warning(f"Данные не найдены ни за {swipe_date}, ни за предыдущие даты")
                return {}

        cur.close()

        if not rows:
            logger.warning(f"Данные за {swipe_date} не найдены в БД")
            return {}

        result = {}
        for row in rows:
            fn = fix_encoding(row['filial_name'])
            amount = float(row['overdue_amount']) if row['overdue_amount'] else 0
            result[fn] = amount

        logger.info(f"Загружено {len(result)} филиалов за {swipe_date}")
        return result

    except Exception as e:
        logger.error(f"Ошибка чтения данных за {swipe_date}: {e}")
        return {}
    finally:
        if conn:
            return_connection(conn)


def get_summary_data(swipe_date):
    """
    Получает сводные данные (legal, notRecoverable, recoverable) за указанную дату.

    Args:
        swipe_date: str — дата (YYYY-MM-DD)

    Returns:
        dict — {summaryDT: {...}, summarySIUAT: {...}} или пустые dict если данных нет
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        cur.execute("""
            SELECT legal_dt, not_recoverable_dt, recoverable_dt,
                   legal_siuat, not_recoverable_siuat, recoverable_siuat
            FROM swipe_history
            WHERE swipe_date = %s
            LIMIT 1
        """, (swipe_date,))

        row = cur.fetchone()
        cur.close()

        if row:
            result = {
                'summaryDT': {
                    'legal': float(row['legal_dt']) if row['legal_dt'] else 0,
                    'notRecoverable': float(row['not_recoverable_dt']) if row['not_recoverable_dt'] else 0,
                    'recoverable': float(row['recoverable_dt']) if row['recoverable_dt'] else 0
                },
                'summarySIUAT': {
                    'legal': float(row['legal_siuat']) if row['legal_siuat'] else 0,
                    'notRecoverable': float(row['not_recoverable_siuat']) if row['not_recoverable_siuat'] else 0,
                    'recoverable': float(row['recoverable_siuat']) if row['recoverable_siuat'] else 0
                }
            }
            logger.info(f"Загружены сводные данные за {swipe_date}")
            return result
        else:
            logger.warning(f"Сводные данные за {swipe_date} не найдены")
            return {
                'summaryDT': {'legal': 0, 'notRecoverable': 0, 'recoverable': 0},
                'summarySIUAT': {'legal': 0, 'notRecoverable': 0, 'recoverable': 0}
            }

    except Exception as e:
        logger.error(f"Ошибка чтения сводных данных за {swipe_date}: {e}")
        return {
            'summaryDT': {'legal': 0, 'notRecoverable': 0, 'recoverable': 0},
            'summarySIUAT': {'legal': 0, 'notRecoverable': 0, 'recoverable': 0}
        }
    finally:
        if conn:
            return_connection(conn)


def get_swipe_dates(from_date=None, to_date=None):
    """
    Возвращает список дат сверок.

    Returns:
        list of dict — [{'date': '2026-04-15', 'total_overdue': 586395480.60, ...}, ...]
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        query = ("SELECT swipe_date, total_overdue, total_debt, filial_count, "
                 "counterparty_count, created_at FROM swipe_history")
        params = []
        conditions = []

        if from_date:
            conditions.append("swipe_date >= %s")
            params.append(from_date)
        if to_date:
            conditions.append("swipe_date <= %s")
            params.append(to_date)

        if conditions:
            query += " WHERE " + " AND ".join(conditions)

        query += " ORDER BY swipe_date ASC"

        cur.execute(query, params)
        rows = cur.fetchall()

        result = []
        for row in rows:
            result.append({
                'date': row['swipe_date'].isoformat() if isinstance(row['swipe_date'], (date, datetime)) else str(row['swipe_date']),
                'total_overdue': float(row['total_overdue']) if row['total_overdue'] else 0,
                'total_debt': float(row['total_debt']) if row['total_debt'] else 0,
                'filial_count': row['filial_count'],
                'counterparty_count': row['counterparty_count'],
            })

        cur.close()
        return result

    except Exception as e:
        logger.error(f"Ошибка чтения дат сверок: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def get_swipe_raw_data(from_date, to_date, filial_name=None):
    """
    Возвращает сырые данные сверок для расчета взвешенного среднего.

    Args:
        from_date: str — дата начала периода (YYYY-MM-DD)
        to_date: str — дата конца периода (YYYY-MM-DD)
        filial_name: str или None

    Returns:
        list of dict
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        query = """
            SELECT swipe_date, filial_name, overdue_amount
            FROM filial_snapshots
            WHERE swipe_date >= %s AND swipe_date <= %s
        """
        params = [from_date, to_date]

        if filial_name:
            query += " AND filial_name = %s"
            params.append(filial_name)

        query += " ORDER BY filial_name ASC, swipe_date ASC"

        cur.execute(query, params)
        rows = cur.fetchall()

        result = []
        for row in rows:
            d = row['swipe_date'].isoformat() if isinstance(row['swipe_date'], (date, datetime)) else str(row['swipe_date'])
            fn = fix_encoding(row['filial_name'])
            amount = float(row['overdue_amount']) if row['overdue_amount'] else 0

            result.append({
                'date': d,
                'filial': fn,
                'overdue': amount
            })

        cur.close()
        logger.info(f"get_swipe_raw_data: загружено {len(result)} записей за период {from_date} - {to_date}")
        return result

    except Exception as e:
        logger.error(f"Ошибка чтения сырых данных сверок: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def get_filial_trend(from_date, to_date, filial_name=None):
    """
    Возвращает данные для графика динамики ПДЗ по филиалам.

    Returns:
        {
            'dates': ['2026-04-10', '2026-04-11', ...],
            'series': [{'name': 'ДТ ТУРУХАНСК', 'data': [...]}, ...]
        }
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        query = """
            SELECT swipe_date, filial_name, overdue_amount
            FROM filial_snapshots
            WHERE swipe_date >= %s AND swipe_date <= %s
        """
        params = [from_date, to_date]

        if filial_name:
            query += " AND filial_name = %s"
            params.append(filial_name)

        query += " ORDER BY swipe_date ASC, filial_name ASC"

        cur.execute(query, params)
        rows = cur.fetchall()

        dates_set = set()
        filial_data = {}

        for row in rows:
            d = row['swipe_date'].isoformat() if isinstance(row['swipe_date'], (date, datetime)) else str(row['swipe_date'])
            fn = fix_encoding(row['filial_name'])
            amount = float(row['overdue_amount']) if row['overdue_amount'] else 0

            dates_set.add(d)
            if fn not in filial_data:
                filial_data[fn] = {}
            filial_data[fn][d] = amount

        dates = sorted(dates_set)

        series = []
        for fn, data in sorted(filial_data.items()):
            series.append({
                'name': fn,
                'data': [data.get(d, 0) for d in dates]
            })

        cur.close()
        return {
            'dates': dates,
            'series': series
        }

    except Exception as e:
        logger.error(f"Ошибка чтения тренда филиалов: {e}")
        return {'dates': [], 'series': []}
    finally:
        if conn:
            return_connection(conn)


def get_counterparty_trend(from_date, to_date, filial_name=None, counterparty_name=None):
    """
    Возвращает данные для графика динамики по контрагентам.

    Returns:
        {'dates': [...], 'series': [...]}
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        query = """
            SELECT swipe_date, filial_name, counterparty_name, debt_amount
            FROM counterparty_snapshots
            WHERE swipe_date >= %s AND swipe_date <= %s
        """
        params = [from_date, to_date]

        if filial_name:
            query += " AND filial_name = %s"
            params.append(filial_name)
        if counterparty_name:
            query += " AND counterparty_name = %s"
            params.append(counterparty_name)

        query += " ORDER BY swipe_date ASC"

        cur.execute(query, params)
        rows = cur.fetchall()

        dates_set = set()
        cp_data = {}

        for row in rows:
            d = row['swipe_date'].isoformat() if isinstance(row['swipe_date'], (date, datetime)) else str(row['swipe_date'])
            fn = fix_encoding(row['filial_name'])
            cp = fix_encoding(row['counterparty_name'])
            amount = float(row['debt_amount']) if row['debt_amount'] else 0

            label = f"{cp} ({fn})"
            dates_set.add(d)
            if label not in cp_data:
                cp_data[label] = {}
            cp_data[label][d] = amount

        dates = sorted(dates_set)

        series = []
        for label, data in sorted(cp_data.items()):
            series.append({
                'name': label,
                'data': [data.get(d, 0) for d in dates]
            })

        cur.close()
        return {
            'dates': dates,
            'series': series
        }

    except Exception as e:
        logger.error(f"Ошибка чтения тренда контрагентов: {e}")
        return {'dates': [], 'series': []}
    finally:
        if conn:
            return_connection(conn)


def get_filial_list(from_date=None, to_date=None):
    """Возвращает список уникальных филиалов за период"""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        query = "SELECT DISTINCT filial_name FROM filial_snapshots"
        params = []
        conditions = []

        if from_date:
            conditions.append("swipe_date >= %s")
            params.append(from_date)
        if to_date:
            conditions.append("swipe_date <= %s")
            params.append(to_date)

        if conditions:
            query += " WHERE " + " AND ".join(conditions)

        query += " ORDER BY filial_name ASC"

        cur.execute(query, params)
        result = [fix_encoding(row[0]) for row in cur.fetchall()]
        cur.close()
        return result

    except Exception as e:
        logger.error(f"Ошибка чтения списка филиалов: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def get_counterparty_list(filial_name=None):
    """Возвращает список уникальных контрагентов (опционально фильтруя по филиалу)"""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        query = "SELECT DISTINCT counterparty_name FROM counterparty_snapshots"
        params = []

        if filial_name:
            query += " WHERE filial_name = %s"
            params.append(filial_name)

        query += " ORDER BY counterparty_name ASC"

        cur.execute(query, params)
        result = [fix_encoding(row[0]) for row in cur.fetchall()]
        cur.close()
        return result

    except Exception as e:
        logger.error(f"Ошибка чтения списка контрагентов: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def get_summary(from_date, to_date):
    """
    Возвращает сводную статистику за период.

    Returns:
        {
            'min_overdue': ..., 'max_overdue': ..., 'avg_overdue': ...,
            'latest_overdue': ..., 'trend': 'up' | 'down' | 'stable', 'swipe_count': ...
        }
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        cur.execute("""
            SELECT
                COUNT(*) as swipe_count,
                MIN(total_overdue) as min_overdue,
                MAX(total_overdue) as max_overdue,
                AVG(total_overdue) as avg_overdue,
                (SELECT total_overdue FROM swipe_history WHERE swipe_date <= %s ORDER BY swipe_date DESC LIMIT 1) as first_overdue,
                (SELECT total_overdue FROM swipe_history WHERE swipe_date <= %s ORDER BY swipe_date DESC LIMIT 1) as last_overdue
            FROM swipe_history
            WHERE swipe_date >= %s AND swipe_date <= %s
        """, (from_date, to_date, from_date, to_date))

        row = cur.fetchone()
        cur.close()

        if not row or row['swipe_count'] == 0:
            return None

        first = float(row['first_overdue']) if row['first_overdue'] else 0
        last = float(row['last_overdue']) if row['last_overdue'] else 0

        if first > 0 and last > 0:
            change = ((last - first) / first) * 100
            trend = 'up' if change > 5 else ('down' if change < -5 else 'stable')
        else:
            trend = 'stable'

        return {
            'min_overdue': float(row['min_overdue']) if row['min_overdue'] else 0,
            'max_overdue': float(row['max_overdue']) if row['max_overdue'] else 0,
            'avg_overdue': float(row['avg_overdue']) if row['avg_overdue'] else 0,
            'latest_overdue': last,
            'trend': trend,
            'swipe_count': row['swipe_count'],
        }

    except Exception as e:
        logger.error(f"Ошибка чтения сводки: {e}")
        return None
    finally:
        if conn:
            return_connection(conn)


def delete_swipe_data(swipe_date):
    """Удаляет данные сверки за указанную дату (CASCADE удалит связанные записи)"""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        cur.execute("DELETE FROM swipe_history WHERE swipe_date = %s RETURNING id", (swipe_date,))
        deleted = cur.fetchone()
        conn.commit()
        cur.close()

        if deleted:
            logger.info(f"Удалены данные за {swipe_date}")
            return {'success': True, 'message': f'Данные за {swipe_date} удалены'}
        else:
            return {'success': False, 'message': f'Данные за {swipe_date} не найдены'}

    except Exception as e:
        logger.error(f"Ошибка удаления: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


# ============================================================
# ФУНКЦИИ ДЛЯ РАБОТЫ С МАППИНГОМ СЧЕТОВ (account_mapping)
# ============================================================

def get_account_mapping(active_only=True):
    """
    Возвращает все записи маппинга счетов.

    Args:
        active_only: bool — если True, возвращает только активные (is_active=true)

    Returns:
        list of dict
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        query = ("SELECT id, account_number, company_name, bank_name, "
                 "is_active, created_at, updated_at FROM account_mapping")
        params = []

        if active_only:
            query += " WHERE is_active = true"

        query += " ORDER BY company_name ASC, account_number ASC"

        cur.execute(query, params)
        rows = cur.fetchall()
        cur.close()

        result = []
        for row in rows:
            result.append({
                'id': row['id'],
                'account_number': row['account_number'],
                'company_name': fix_encoding(row['company_name']),
                'bank_name': fix_encoding(row['bank_name']) if row['bank_name'] else '',
                'is_active': row['is_active'],
                'created_at': row['created_at'].isoformat() if row['created_at'] else None,
                'updated_at': row['updated_at'].isoformat() if row['updated_at'] else None
            })

        return result

    except Exception as e:
        logger.error(f"Ошибка чтения маппинга счетов: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def add_account_mapping(account_number, company_name, bank_name=''):
    """
    Добавляет новую запись в маппинг счетов.

    Returns:
        dict — {'success': True, 'id': N} или {'success': False, 'error': '...'}
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        clean_account = account_number.replace(' ', '')

        cur.execute("""
            INSERT INTO account_mapping (account_number, company_name, bank_name, updated_at)
            VALUES (%s, %s, %s, NOW())
            ON CONFLICT (account_number) DO UPDATE SET
                company_name = EXCLUDED.company_name,
                bank_name = EXCLUDED.bank_name,
                is_active = true,
                updated_at = NOW()
            RETURNING id
        """, (clean_account, company_name, bank_name))

        row = cur.fetchone()
        conn.commit()
        cur.close()

        new_id = row[0] if row else None
        logger.info(f"Добавлен/обновлён счёт {clean_account} → {company_name} (банк: {bank_name})")
        return {'success': True, 'id': new_id}

    except Exception as e:
        logger.error(f"Ошибка добавления маппинга счета: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def update_account_mapping(mapping_id, account_number=None, company_name=None,
                           bank_name=None, is_active=None):
    """
    Обновляет существующую запись маппинга.

    Returns:
        dict — {'success': True} или {'success': False, 'error': '...'}
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        updates = []
        params = []

        if account_number is not None:
            updates.append("account_number = %s")
            params.append(account_number.replace(' ', ''))
        if company_name is not None:
            updates.append("company_name = %s")
            params.append(company_name)
        if bank_name is not None:
            updates.append("bank_name = %s")
            params.append(bank_name)
        if is_active is not None:
            updates.append("is_active = %s")
            params.append(is_active)

        if not updates:
            return {'success': False, 'error': 'Нет полей для обновления'}

        updates.append("updated_at = NOW()")
        params.append(mapping_id)

        cur.execute(
            f"UPDATE account_mapping SET {', '.join(updates)} WHERE id = %s",
            params
        )
        conn.commit()
        cur.close()

        logger.info(f"Обновлён маппинг id={mapping_id}")
        return {'success': True}

    except Exception as e:
        logger.error(f"Ошибка обновления маппинга: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def delete_account_mapping(mapping_id):
    """
    Мягкое удаление записи маппинга (is_active = false).

    Returns:
        dict — {'success': True} или {'success': False, 'error': '...'}
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        cur.execute("""
            UPDATE account_mapping
            SET is_active = false, updated_at = NOW()
            WHERE id = %s
        """, (mapping_id,))
        conn.commit()
        cur.close()

        logger.info(f"Мягкое удаление маппинга id={mapping_id}")
        return {'success': True}

    except Exception as e:
        logger.error(f"Ошибка удаления маппинга: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def get_companies_list():
    """Возвращает список уникальных названий компаний из маппинга (для автодополнения)."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        cur.execute("""
            SELECT DISTINCT company_name
            FROM account_mapping
            WHERE is_active = true
            ORDER BY company_name ASC
        """)
        result = [fix_encoding(row[0]) for row in cur.fetchall()]
        cur.close()
        return result

    except Exception as e:
        logger.error(f"Ошибка чтения списка компаний: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def get_banks_list():
    """Возвращает список уникальных названий банков из маппинга (для автодополнения)."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        cur.execute("""
            SELECT DISTINCT bank_name
            FROM account_mapping
            WHERE is_active = true AND bank_name IS NOT NULL AND bank_name != ''
            ORDER BY bank_name ASC
        """)
        result = [fix_encoding(row[0]) for row in cur.fetchall()]
        cur.close()
        return result

    except Exception as e:
        logger.error(f"Ошибка чтения списка банков: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def sync_account_mapping_from_parser(mapping_list):
    """
    Синхронизирует маппинг из клиента (parser.js) с БД.

    Returns:
        dict — {'success': True, 'added': N, 'updated': N}
    """
    conn = None
    added = 0
    updated = 0
    try:
        conn = get_connection()
        cur = conn.cursor()

        for item in mapping_list:
            clean_account = item.get('account_number', '').replace(' ', '')
            company = item.get('company_name', '')
            bank = item.get('bank_name', '')

            if not clean_account or not company:
                continue

            cur.execute("""
                INSERT INTO account_mapping (account_number, company_name, bank_name, updated_at)
                VALUES (%s, %s, %s, NOW())
                ON CONFLICT (account_number) DO UPDATE SET
                    company_name = EXCLUDED.company_name,
                    bank_name = EXCLUDED.bank_name,
                    updated_at = NOW()
                RETURNING (xmax = 0) AS is_insert
            """, (clean_account, company, bank))

            row = cur.fetchone()
            if row and row[0]:
                added += 1
            else:
                updated += 1

        conn.commit()
        cur.close()

        logger.info(f"Синхронизация маппинга: добавлено {added}, обновлено {updated}")
        return {'success': True, 'added': added, 'updated': updated}

    except Exception as e:
        logger.error(f"Ошибка синхронизации маппинга: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


# ============================================================
# БИБЛИОТЕКА КОНТРАГЕНТОВ (contractors_library)
# ============================================================

def get_all_contractors():
    """Возвращает все записи из библиотеки контрагентов."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        cur.execute("""
            SELECT id, name, organization, explanation, created_at, updated_at
            FROM contractors_library
            ORDER BY name ASC
        """)
        rows = cur.fetchall()
        cur.close()

        result = []
        for row in rows:
            result.append({
                'id': row['id'],
                'name': fix_encoding(row['name']),
                'organization': fix_encoding(row['organization']) if row['organization'] else '',
                'explanation': fix_encoding(row['explanation']) if row['explanation'] else '',
                'created_at': row['created_at'].isoformat() if row['created_at'] else None,
                'updated_at': row['updated_at'].isoformat() if row['updated_at'] else None
            })

        logger.info(f"Загружено {len(result)} контрагентов из библиотеки")
        return result

    except Exception as e:
        logger.error(f"Ошибка чтения библиотеки контрагентов: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def find_contractor_by_name(name):
    """Ищет контрагента по точному совпадению имени (нормализованному)."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        cur.execute("""
            SELECT id, name, organization, explanation
            FROM contractors_library
            WHERE UPPER(TRIM(name)) = %s
            LIMIT 1
        """, (name.upper().strip(),))
        row = cur.fetchone()
        cur.close()

        if row:
            return {
                'id': row['id'],
                'name': fix_encoding(row['name']),
                'organization': fix_encoding(row['organization']) if row['organization'] else '',
                'explanation': fix_encoding(row['explanation']) if row['explanation'] else ''
            }
        return None

    except Exception as e:
        logger.error(f"Ошибка поиска контрагента: {e}")
        return None
    finally:
        if conn:
            return_connection(conn)


def insert_or_update_contractor(name, organization='', explanation=''):
    """
    Добавляет или обновляет контрагента в библиотеке.

    Returns:
        dict — {'success': True/False, 'id': N, 'action': 'inserted'/'updated'}
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        cur.execute("""
            INSERT INTO contractors_library (name, organization, explanation, updated_at)
            VALUES (%s, %s, %s, NOW())
            ON CONFLICT (name) DO UPDATE SET
                organization = COALESCE(NULLIF(EXCLUDED.organization, ''), contractors_library.organization),
                explanation = COALESCE(NULLIF(EXCLUDED.explanation, ''), contractors_library.explanation),
                updated_at = NOW()
            RETURNING id, (xmax = 0) AS is_insert
        """, (name, organization, explanation))

        row = cur.fetchone()
        conn.commit()
        cur.close()

        if row:
            is_insert = row[1]
            action = 'inserted' if is_insert else 'updated'
            logger.info(f"Контрагент '{name}': {action} (id={row[0]})")
            return {'success': True, 'id': row[0], 'action': action}
        return {'success': False, 'error': 'Не удалось сохранить'}

    except Exception as e:
        logger.error(f"Ошибка сохранения контрагента '{name}': {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def update_contractor(contractor_id, name=None, organization=None, explanation=None):
    """
    Обновляет поля контрагента по ID.

    Returns:
        dict — {'success': True/False}
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        updates = []
        params = []

        if name is not None:
            updates.append("name = %s")
            params.append(name)
        if organization is not None:
            updates.append("organization = %s")
            params.append(organization)
        if explanation is not None:
            updates.append("explanation = %s")
            params.append(explanation)

        if not updates:
            return {'success': False, 'error': 'Нет полей для обновления'}

        updates.append("updated_at = NOW()")
        params.append(contractor_id)

        cur.execute(
            f"UPDATE contractors_library SET {', '.join(updates)} WHERE id = %s",
            params
        )
        conn.commit()
        cur.close()

        logger.info(f"Обновлён контрагент id={contractor_id}")
        return {'success': True}

    except Exception as e:
        logger.error(f"Ошибка обновления контрагента id={contractor_id}: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def delete_contractor(contractor_id):
    """Удаляет контрагента по ID."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        cur.execute("DELETE FROM contractors_library WHERE id = %s RETURNING id", (contractor_id,))
        deleted = cur.fetchone()
        conn.commit()
        cur.close()

        if deleted:
            logger.info(f"Удалён контрагент id={contractor_id}")
            return {'success': True, 'message': 'Контрагент удалён'}
        else:
            return {'success': False, 'message': 'Контрагент не найден'}

    except Exception as e:
        logger.error(f"Ошибка удаления контрагента id={contractor_id}: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def import_contractors_batch(contractors_list):
    """
    Массовый импорт контрагентов (перезапись при совпадении имени).

    Returns:
        dict — {'success': True, 'added': N, 'updated': N}
    """
    conn = None
    added = 0
    updated = 0
    try:
        conn = get_connection()
        cur = conn.cursor()

        for item in contractors_list:
            name = item.get('name', '').strip()
            if not name:
                continue

            organization = item.get('organization', '')
            explanation = item.get('explanation', '')

            cur.execute("""
                INSERT INTO contractors_library (name, organization, explanation, updated_at)
                VALUES (%s, %s, %s, NOW())
                ON CONFLICT (name) DO UPDATE SET
                    organization = COALESCE(NULLIF(EXCLUDED.organization, ''), contractors_library.organization),
                    explanation = COALESCE(NULLIF(EXCLUDED.explanation, ''), contractors_library.explanation),
                    updated_at = NOW()
                RETURNING (xmax = 0) AS is_insert
            """, (name, organization, explanation))

            row = cur.fetchone()
            if row and row[0]:
                added += 1
            else:
                updated += 1

        conn.commit()
        cur.close()

        logger.info(f"Импорт библиотеки: добавлено {added}, обновлено {updated}")
        return {'success': True, 'added': added, 'updated': updated}

    except Exception as e:
        logger.error(f"Ошибка импорта библиотеки: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def delete_all_contractors():
    """Удаляет ВСЕ записи из библиотеки контрагентов."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()

        cur.execute("DELETE FROM contractors_library")
        count = cur.rowcount
        conn.commit()
        cur.close()

        logger.info(f"Удалено {count} записей из библиотеки контрагентов")
        return {'success': True, 'deleted': count}

    except Exception as e:
        logger.error(f"Ошибка очистки библиотеки: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def get_contractors_stats():
    """Возвращает статистику библиотеки контрагентов."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

        cur.execute("""
            SELECT
                COUNT(*) as total,
                COUNT(*) FILTER (WHERE explanation IS NOT NULL AND explanation != '') as with_explanation,
                COUNT(*) FILTER (WHERE organization IS NOT NULL AND organization != '') as with_organization
            FROM contractors_library
        """)
        row = cur.fetchone()
        cur.close()

        return {
            'total': row['total'] if row else 0,
            'withExplanation': row['with_explanation'] if row else 0,
            'withOrganization': row['with_organization'] if row else 0
        }

    except Exception as e:
        logger.error(f"Ошибка статистики библиотеки: {e}")
        return {'total': 0, 'withExplanation': 0, 'withOrganization': 0}
    finally:
        if conn:
            return_connection(conn)


# ============================================================
# ПРАВИЛА ИСКЛЮЧЕНИЙ (exclusion_rules)
# ============================================================

def get_exclusion_rules():
    """Возвращает все правила исключения."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT id, rule_type, pattern, is_regex, created_at, updated_at FROM exclusion_rules ORDER BY id ASC")
        rows = cur.fetchall()
        cur.close()
        result = []
        for row in rows:
            result.append({
                'id': row['id'],
                'type': row['rule_type'],
                'pattern': fix_encoding(row['pattern']),
                'is_regex': row['is_regex']
            })
        return result
    except Exception as e:
        logger.error(f"Ошибка чтения правил исключения: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def save_exclusion_rule(rule_type, pattern, is_regex=False):
    """Добавляет правило исключения."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO exclusion_rules (rule_type, pattern, is_regex, updated_at) VALUES (%s, %s, %s, NOW()) RETURNING id",
            (rule_type, pattern, is_regex))
        row = cur.fetchone()
        conn.commit()
        cur.close()
        if row:
            return {'success': True, 'id': row[0]}
        return {'success': False, 'error': 'Не удалось добавить правило'}
    except Exception as e:
        logger.error(f"Ошибка добавления правила: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def delete_exclusion_rule(rule_id):
    """Удаляет правило исключения по ID."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("DELETE FROM exclusion_rules WHERE id = %s RETURNING id", (rule_id,))
        deleted = cur.fetchone()
        conn.commit()
        cur.close()
        if deleted:
            return {'success': True}
        return {'success': False, 'error': 'Правило не найдено'}
    except Exception as e:
        logger.error(f"Ошибка удаления правила: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def update_exclusion_rule(rule_id, rule_type=None, pattern=None, is_regex=None):
    """Обновляет правило исключения по ID."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        updates = []
        params = []
        if rule_type is not None:
            updates.append("rule_type = %s")
            params.append(rule_type)
        if pattern is not None:
            updates.append("pattern = %s")
            params.append(pattern)
        if is_regex is not None:
            updates.append("is_regex = %s")
            params.append(is_regex)
        if not updates:
            return {'success': False, 'error': 'Нет полей для обновления'}
        updates.append("updated_at = NOW()")
        params.append(rule_id)
        cur.execute(f"UPDATE exclusion_rules SET {', '.join(updates)} WHERE id = %s", params)
        conn.commit()
        cur.close()
        return {'success': True}
    except Exception as e:
        logger.error(f"Ошибка обновления правила исключения: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


# ============================================================
# ПРАВИЛА КАТЕГОРИЗАЦИИ (categorization_rules)
# ============================================================

def get_categorization_rules():
    """Возвращает все правила категоризации."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT id, field, pattern, display_name, created_at, updated_at FROM categorization_rules ORDER BY id ASC")
        rows = cur.fetchall()
        cur.close()
        result = []
        for row in rows:
            result.append({
                'id': row['id'],
                'field': row['field'],
                'pattern': fix_encoding(row['pattern']),
                'display_name': fix_encoding(row['display_name'])
            })
        return result
    except Exception as e:
        logger.error(f"Ошибка чтения правил категоризации: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def save_categorization_rule(field, pattern, display_name):
    """Добавляет правило категоризации."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO categorization_rules (field, pattern, display_name, updated_at) VALUES (%s, %s, %s, NOW()) RETURNING id",
            (field, pattern, display_name))
        row = cur.fetchone()
        conn.commit()
        cur.close()
        if row:
            return {'success': True, 'id': row[0]}
        return {'success': False, 'error': 'Не удалось добавить правило'}
    except Exception as e:
        logger.error(f"Ошибка добавления правила: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def delete_categorization_rule(rule_id):
    """Удаляет правило категоризации по ID."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("DELETE FROM categorization_rules WHERE id = %s RETURNING id", (rule_id,))
        deleted = cur.fetchone()
        conn.commit()
        cur.close()
        if deleted:
            return {'success': True}
        return {'success': False, 'error': 'Правило не найдено'}
    except Exception as e:
        logger.error(f"Ошибка удаления правила: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def update_categorization_rule(rule_id, field=None, pattern=None, display_name=None):
    """Обновляет правило категоризации по ID."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        updates = []
        params = []
        if field is not None:
            updates.append("field = %s")
            params.append(field)
        if pattern is not None:
            updates.append("pattern = %s")
            params.append(pattern)
        if display_name is not None:
            updates.append("display_name = %s")
            params.append(display_name)
        if not updates:
            return {'success': False, 'error': 'Нет полей для обновления'}
        updates.append("updated_at = NOW()")
        params.append(rule_id)
        cur.execute(f"UPDATE categorization_rules SET {', '.join(updates)} WHERE id = %s", params)
        conn.commit()
        cur.close()
        return {'success': True}
    except Exception as e:
        logger.error(f"Ошибка обновления правила категоризации: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


# ============================================================
# СИНОНИМЫ КОМПАНИЙ (company_aliases)
# ============================================================

def get_company_aliases():
    """Возвращает все синонимы компаний."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT id, pattern, canonical, match_type, created_at, updated_at FROM company_aliases ORDER BY id ASC")
        rows = cur.fetchall()
        cur.close()
        result = []
        for row in rows:
            result.append({
                'id': row['id'],
                'pattern': fix_encoding(row['pattern']),
                'canonical': fix_encoding(row['canonical']),
                'match_type': row['match_type']
            })
        return result
    except Exception as e:
        logger.error(f"Ошибка чтения синонимов: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def save_company_alias(pattern, canonical, match_type='contains'):
    """Добавляет синоним компании."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO company_aliases (pattern, canonical, match_type, updated_at) VALUES (%s, %s, %s, NOW()) RETURNING id",
            (pattern, canonical, match_type))
        row = cur.fetchone()
        conn.commit()
        cur.close()
        if row:
            return {'success': True, 'id': row[0]}
        return {'success': False, 'error': 'Не удалось добавить синоним'}
    except Exception as e:
        logger.error(f"Ошибка добавления синонима: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def delete_company_alias(alias_id):
    """Удаляет синоним по ID."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("DELETE FROM company_aliases WHERE id = %s RETURNING id", (alias_id,))
        deleted = cur.fetchone()
        conn.commit()
        cur.close()
        if deleted:
            return {'success': True}
        return {'success': False, 'error': 'Синоним не найден'}
    except Exception as e:
        logger.error(f"Ошибка удаления синонима: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def update_company_alias(alias_id, pattern=None, canonical=None, match_type=None):
    """Обновляет синоним компании по ID."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        updates = []
        params = []
        if pattern is not None:
            updates.append("pattern = %s")
            params.append(pattern)
        if canonical is not None:
            updates.append("canonical = %s")
            params.append(canonical)
        if match_type is not None:
            updates.append("match_type = %s")
            params.append(match_type)
        if not updates:
            return {'success': False, 'error': 'Нет полей для обновления'}
        updates.append("updated_at = NOW()")
        params.append(alias_id)
        cur.execute(f"UPDATE company_aliases SET {', '.join(updates)} WHERE id = %s", params)
        conn.commit()
        cur.close()
        return {'success': True}
    except Exception as e:
        logger.error(f"Ошибка обновления синонима: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


# ============================================================
# СБОРНАЯ ФУНКЦИЯ — ВСЯ КОНФИГУРАЦИЯ РАЗОМ
# ============================================================

def get_all_config():
    """Возвращает полную конфигурацию: маппинг счетов, правила исключения, категоризации, синонимы."""
    raw = _raw_query(
        "SELECT account_number, company_name, bank_name, is_active FROM account_mapping WHERE is_active = true")
    account_mapping = {}
    for row in raw:
        account_mapping[row['account_number']] = {
            'company': fix_encoding(row['company_name']),
            'bank': fix_encoding(row['bank_name']) if row['bank_name'] else '',
            'is_active': row['is_active']
        }
    return {
        'account_mapping': account_mapping,
        'exclusion_rules': get_exclusion_rules(),
        'categorization_rules': get_categorization_rules(),
        'company_aliases': get_company_aliases()
    }


def _raw_query(query, params=None):
    """Выполняет сырой запрос и возвращает список RealDictRow."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(query, params or [])
        rows = cur.fetchall()
        cur.close()
        return rows
    except Exception as e:
        logger.error(f"Ошибка raw query: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


# ============================================================
# HEALTH CHECK — ПРОВЕРКА ДОСТУПНОСТИ БД
# ============================================================# ============================================================
# УПРАВЛЕНИЕ СРОЧНЫМИ ДЕПОЗИТАМИ (term_deposits)
# ============================================================

def get_term_deposits():
    """Возвращает все активные срочные депозиты (с end_date > сегодня)."""
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
        cur.execute("""
            SELECT id, account_number, amount, rate, start_date, end_date
            FROM term_deposits
            WHERE end_date >= CURRENT_DATE
            ORDER BY account_number, end_date
        """)
        rows = cur.fetchall()
        cur.close()
        return [dict(r) for r in rows]
    except Exception as e:
        logger.error(f"Ошибка получения срочных депозитов: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)

def sync_term_deposits(deposits):
    """
    Слияние срочных депозитов с БД (merge, не полная замена).
    Удаляет записи только для переданных счетов, затем вставляет новые.
    deposits: { account_number: [{amount, rate, startDate, endDate}, ...] }
    """
    conn = None
    synced = 0
    try:
        conn = get_connection()
        cur = conn.cursor()
        # Удаляем записи только для тех счетов, которые переданы
        accounts = list(deposits.keys())
        if accounts:
            cur.execute("DELETE FROM term_deposits WHERE account_number = ANY(%s)", (accounts,))
        for account, deps in deposits.items():
            for dep in deps:
                if not dep.get('endDate'):
                    continue  # пропускаем обычные депозиты
                cur.execute("""
                    INSERT INTO term_deposits (account_number, amount, rate, start_date, end_date)
                    VALUES (%s, %s, %s, %s, %s)
                """, (
                    account,
                    dep.get('amount', 0),
                    dep.get('rate', 0),
                    dep.get('startDate'),
                    dep.get('endDate')
                ))
                synced += 1
        conn.commit()
        cur.close()
        logger.info(f"Синхронизировано {synced} срочных депозитов на {len(accounts)} счетах")
        return synced
    except Exception as e:
        logger.error(f"Ошибка синхронизации срочных депозитов: {e}")
        if conn:
            conn.rollback()
        return 0
    finally:
        if conn:
            return_connection(conn)

# ============================================================
# ФУНКЦИИ ДЛЯ КОНТРОЛЯ ЛИМИТОВ ПЛАТЕЖЕЙ ПОСТАВЩИКАМ
# ============================================================

_PLP_COLUMNS = [
    'id', 'limit_month', 'period', 'stage', 'registry_id', 'confirmed_at', 'article', 'detail_limit',
    'amount', 'recipient', 'request_no', 'request_date', 'status',
    'organization', 'subdivision', 'priority', 'purpose', 'payment_date',
    'applicant', 'payment_type'
]

# Стадии подтверждения: черновик (первичное сохранение) и финальный вариант
PLP_STAGE_DRAFT = 'draft'
PLP_STAGE_FINAL = 'final'

_PLP_SCHEMA_APPLIED = False

_PLP_CREATE_SQL = """
CREATE TABLE IF NOT EXISTS payment_limits_payments (
    id SERIAL PRIMARY KEY,
    limit_month DATE NOT NULL,
    period SMALLINT NOT NULL CHECK (period BETWEEN 1 AND 4),
    stage VARCHAR(10) NOT NULL DEFAULT 'final',
    registry_id INTEGER,
    confirmed_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    article VARCHAR(255) NOT NULL,
    detail_limit VARCHAR(255) DEFAULT '',
    amount NUMERIC(20, 2) NOT NULL DEFAULT 0,
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
CREATE INDEX IF NOT EXISTS idx_plp_registry ON payment_limits_payments(registry_id);
"""

# Реестр загрузки: исходный файл + снятые чекбоксами строки для каждой стадии.
# Нужен, чтобы финальный файл можно было выгрузить с любого компьютера и после перезапуска.
_PLP_REGISTRY_CREATE_SQL = """
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
    final_exported_at TIMESTAMP,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_plp_registry_period ON payment_limits_registry(limit_month, period);
"""

# Настройки лимитов хранятся одной строкой (id = 1) в виде JSONB
_PLP_CONFIG_CREATE_SQL = """
CREATE TABLE IF NOT EXISTS payment_limits_config (
    id SMALLINT PRIMARY KEY DEFAULT 1 CHECK (id = 1),
    config JSONB NOT NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
"""

# Значения лимитов по умолчанию (рубли) — из технического задания
DEFAULT_PAYMENT_LIMITS_CONFIG = {
    "globalLimitP12": 31000000,
    "categories": [
        {
            "id": "cat_aks",
            "name": "Аренда, Коммуналка, Связь",
            "limits": {"1": 6500000, "2": 1000000, "3": 7000000, "4": 0},
            "details": [
                {"id": "det_internet", "name": "Связь и интернет", "limits": {"1": 3500000, "2": 0, "3": 0, "4": 0}},
                {"id": "det_utility", "name": "Коммунальные услуги", "limits": {"1": 1500000, "2": 1000000, "3": 0, "4": 0}},
                {"id": "det_rent_flat", "name": "Аренда квартир", "limits": {"1": 1500000, "2": 0, "3": 0, "4": 0}},
                {"id": "det_rent_office", "name": "Аренда баз и офисов", "limits": {"1": 0, "2": 0, "3": 7000000, "4": 0}},
            ],
        },
        {
            "id": "cat_life",
            "name": "Жизнеобеспечение",
            "limits": {"1": 6500000, "2": 5000000, "3": 7000000, "4": 7000000},
            "details": [],
        },
        {
            "id": "cat_fleet",
            "name": "Автопарк",
            "limits": {"1": 4500000, "2": 2500000, "3": 20000000, "4": 14000000},
            "details": [],
        },
        {
            "id": "cat_other",
            "name": "Прочая операционка",
            "limits": {"1": 1500000, "2": 1500000, "3": 10000000, "4": 10500000},
            "details": [],
        },
        {
            "id": "cat_overdue",
            "name": "Графики просрочки",
            "limits": {"1": 500000, "2": 500000, "3": 5000000, "4": 19000000},
            "details": [],
        },
    ],
}


def _copy_default_plp_config():
    """Возвращает глубокую копию конфигурации лимитов по умолчанию."""
    return json.loads(json.dumps(DEFAULT_PAYMENT_LIMITS_CONFIG))


def _ensure_plp_schema(conn):
    """Одноразово (на процесс) создаёт таблицы контроля лимитов и доводит схему до актуальной."""
    global _PLP_SCHEMA_APPLIED
    if _PLP_SCHEMA_APPLIED:
        return
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT to_regclass('public.payment_limits_payments'), "
            "to_regclass('public.payment_limits_config'), "
            "to_regclass('public.payment_limits_registry')"
        )
        payments_exists, config_exists, registry_exists = cur.fetchone()
        if payments_exists is None:
            cur.execute(_PLP_CREATE_SQL)
            logger.info("Создана таблица payment_limits_payments (контроль лимитов платежей)")
        else:
            # Стадия подтверждения: черновик / финальный вариант
            cur.execute(
                "ALTER TABLE payment_limits_payments "
                "ADD COLUMN IF NOT EXISTS stage VARCHAR(10) NOT NULL DEFAULT 'final'"
            )
            cur.execute(
                "CREATE INDEX IF NOT EXISTS idx_plp_stage "
                "ON payment_limits_payments(limit_month, period, stage)"
            )
            # Партия загрузки: к какому реестру (загрузке) относятся строки
            cur.execute(
                "ALTER TABLE payment_limits_payments "
                "ADD COLUMN IF NOT EXISTS registry_id INTEGER"
            )
            cur.execute(
                "CREATE INDEX IF NOT EXISTS idx_plp_registry "
                "ON payment_limits_payments(registry_id)"
            )
        if config_exists is None:
            cur.execute(_PLP_CONFIG_CREATE_SQL)
            logger.info("Создана таблица payment_limits_config (настройки лимитов платежей)")
        if registry_exists is None:
            cur.execute(_PLP_REGISTRY_CREATE_SQL)
            logger.info("Создана таблица payment_limits_registry (реестр загрузок лимитов)")
        else:
            # Отметка о выгрузке финального файла (для уведомления на кнопке экспорта)
            cur.execute(
                "ALTER TABLE payment_limits_registry "
                "ADD COLUMN IF NOT EXISTS final_exported_at TIMESTAMP"
            )
            # В периоде может быть несколько сохранённых загрузок (партий):
            # снимаем уникальность (limit_month, period) для накопления сумм
            cur.execute(
                "ALTER TABLE payment_limits_registry "
                "DROP CONSTRAINT IF EXISTS payment_limits_registry_limit_month_period_key"
            )
            cur.execute(
                "CREATE INDEX IF NOT EXISTS idx_plp_registry_period "
                "ON payment_limits_registry(limit_month, period)"
            )
            # Старые строки (до появления партий) привязываем к реестру своего месяца/периода
            cur.execute(
                """
                UPDATE payment_limits_payments p
                   SET registry_id = r.id
                  FROM payment_limits_registry r
                 WHERE p.registry_id IS NULL
                   AND p.limit_month = r.limit_month
                   AND p.period = r.period
                """
            )
        conn.commit()
        cur.close()
    except Exception as e:
        logger.error(f"Не удалось проверить/создать таблицы контроля лимитов: {e}")
        return
    _PLP_SCHEMA_APPLIED = True


def ensure_payment_limits_schema():
    """Публичная обёртка: гарантирует наличие таблицы контроля лимитов платежей."""
    conn = None
    try:
        conn = get_connection()
        _ensure_plp_schema(conn)
        return True
    except Exception as e:
        logger.error(f"Ошибка ensure_payment_limits_schema: {e}")
        return False
    finally:
        if conn:
            return_connection(conn)


def get_payment_limits_payments(month, stage=None):
    """
    Возвращает платежи лимитов за указанный месяц учёта.

    Args:
        month: str/date — 'YYYY-MM' или 'YYYY-MM-01' (первый день месяца)
        stage: 'draft' | 'final' | None — фильтр по стадии подтверждения

    Returns:
        list[dict] — строки таблицы payment_limits_payments за месяц (поле stage включено)
    """
    conn = None
    try:
        conn = get_connection()
        _ensure_plp_schema(conn)
        cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
        columns_sql = ', '.join(_PLP_COLUMNS)
        if stage:
            cur.execute(
                f"SELECT {columns_sql} FROM payment_limits_payments "
                f"WHERE limit_month = %s AND stage = %s ORDER BY period, confirmed_at, id",
                (_PLP_MONTH(month), stage)
            )
        else:
            cur.execute(
                f"SELECT {columns_sql} FROM payment_limits_payments "
                f"WHERE limit_month = %s ORDER BY period, confirmed_at, id",
                (_PLP_MONTH(month),)
            )
        rows = cur.fetchall()
        cur.close()
        return [dict(r) for r in rows]
    except Exception as e:
        logger.error(f"Ошибка получения платежей лимитов за {month}: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def get_payment_limits_effective(month):
    """
    Эффективные сохранённые строки месяца для итогов контроля лимитов.

    По каждой загрузке (партии) берётся финальный вариант, а если его нет — черновик.
    Строки всех сохранённых загрузок периода складываются (накапливаются).

    Returns:
        list[dict] — строки (с полем registry_id)
    """
    rows = get_payment_limits_payments(month)
    groups = {}
    for r in rows:
        key = (r.get('period'), r.get('registry_id'))
        group = groups.setdefault(key, {PLP_STAGE_FINAL: [], PLP_STAGE_DRAFT: []})
        stage = r.get('stage') or PLP_STAGE_FINAL
        group.setdefault(stage, []).append(r)

    result = []
    for key in sorted(groups.keys(), key=lambda k: (k[0] or 0, k[1] or 0)):
        group = groups[key]
        result.extend(group[PLP_STAGE_FINAL] or group[PLP_STAGE_DRAFT])
    return result


def _PLP_MONTH(month):
    """Нормализует месяц ('YYYY-MM' или 'YYYY-MM-01') к DATE 'YYYY-MM-01'."""
    s = str(month)
    if len(s) == 7:
        s = s + '-01'
    return s


def insert_payment_limits_payments(payments):
    """
    Вставляет подтверждённые платежи (пачка) в таблицу payment_limits_payments.

    Args:
        payments: list[dict] — подготовленные сервером строки:
            {limit_month, period, article, detail_limit, amount, recipient,
             request_no, request_date, status, organization, subdivision,
             priority, purpose, payment_date, applicant, payment_type}

    Returns:
        dict — {'success': True, 'count': N} или {'success': False, 'error': '...'}
    """
    conn = None
    try:
        conn = get_connection()
        _ensure_plp_schema(conn)
        cur = conn.cursor()

        rows = []
        for p in payments:
            rows.append((
                _PLP_MONTH(p.get('limit_month')),
                int(p.get('period', 0)),
                p.get('stage') or PLP_STAGE_FINAL,
                p.get('article') or '',
                p.get('detail_limit') or '',
                p.get('amount') or 0,
                p.get('recipient') or '',
                p.get('request_no') or '',
                p.get('request_date'),
                p.get('status') or '',
                p.get('organization') or '',
                p.get('subdivision') or '',
                p.get('priority') or '',
                p.get('purpose') or '',
                p.get('payment_date'),
                p.get('applicant') or '',
                p.get('payment_type') or ''
            ))

        if rows:
            psycopg2.extras.execute_values(
                cur,
                """
                INSERT INTO payment_limits_payments
                    (limit_month, period, stage, article, detail_limit, amount,
                     recipient, request_no, request_date, status,
                     organization, subdivision, priority, purpose,
                     payment_date, applicant, payment_type)
                VALUES %s
                """,
                rows
            )
        conn.commit()
        cur.close()

        logger.info(f"Подтверждено платежей по лимитам: {len(rows)}")
        return {'success': True, 'count': len(rows)}
    except Exception as e:
        logger.error(f"Ошибка сохранения платежей лимитов: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def save_payment_limits_stage(month, period, stage, payments, registry_id=None):
    """
    Сохраняет вариант (черновик/финал) за месяц + период.

    Если задан registry_id — заменяет прежний вариант ТОЛЬКО этой загрузки (партии),
    не затрагивая ранее сохранённые загрузки периода (они накапливаются).
    Если registry_id не задан — прежнее поведение: заменяет вариант всей стадии периода.

    Returns:
        dict — {'success': True, 'count': N, 'replaced': M} или {'success': False, 'error': '...'}
    """
    conn = None
    try:
        conn = get_connection()
        _ensure_plp_schema(conn)
        cur = conn.cursor()
        month_val = _PLP_MONTH(month)
        period_val = int(period)
        registry_val = int(registry_id) if registry_id else None

        if registry_val is not None:
            cur.execute(
                "DELETE FROM payment_limits_payments "
                "WHERE limit_month = %s AND period = %s AND stage = %s AND registry_id = %s",
                (month_val, period_val, stage, registry_val)
            )
        else:
            cur.execute(
                "DELETE FROM payment_limits_payments "
                "WHERE limit_month = %s AND period = %s AND stage = %s",
                (month_val, period_val, stage)
            )
        replaced = cur.rowcount

        rows = []
        for p in payments:
            rows.append((
                month_val, period_val, stage, registry_val,
                p.get('article') or '',
                p.get('detail_limit') or '',
                p.get('amount') or 0,
                p.get('recipient') or '',
                p.get('request_no') or '',
                p.get('request_date'),
                p.get('status') or '',
                p.get('organization') or '',
                p.get('subdivision') or '',
                p.get('priority') or '',
                p.get('purpose') or '',
                p.get('payment_date'),
                p.get('applicant') or '',
                p.get('payment_type') or ''
            ))

        if rows:
            psycopg2.extras.execute_values(
                cur,
                """
                INSERT INTO payment_limits_payments
                    (limit_month, period, stage, registry_id,
                     article, detail_limit, amount,
                     recipient, request_no, request_date, status,
                     organization, subdivision, priority, purpose,
                     payment_date, applicant, payment_type)
                VALUES %s
                """,
                rows
            )
        conn.commit()
        cur.close()
        logger.info(f"Сохранён вариант «{stage}» лимитов: месяц {month_val}, P{period_val}, "
                    f"загрузка #{registry_val}, {len(rows)} строк (заменено {replaced})")
        return {'success': True, 'count': len(rows), 'replaced': replaced}
    except Exception as e:
        logger.error(f"Ошибка сохранения варианта «{stage}» лимитов: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


# ---------------- Реестр загрузок (исходный файл + снятые строки) ----------------

_PLP_REGISTRY_META_SQL = """
SELECT id, limit_month, period, file_name, sheet_name, row_count,
       draft_removed_rows, draft_saved_at, draft_author,
       final_removed_rows, final_saved_at, final_author, final_exported_at,
       created_at, updated_at,
       (file_bytes IS NOT NULL) AS has_file,
       COALESCE(OCTET_LENGTH(file_bytes), 0) AS file_size
  FROM payment_limits_registry
"""


def save_payment_limits_registry(month, period, stage, file_name=None, sheet_name=None,
                                 file_bytes=None, row_count=None, removed_rows=None, author='',
                                 registry_id=None):
    """
    Сохраняет/обновляет реестр загрузки (партию) месяца + периода:
    исходный файл (один раз на загрузку) и набор снятых чекбоксами строк для указанной стадии.

    Если registry_id не задан — создаётся НОВАЯ загрузка (партия): так внутри одного периода
    накапливаются все сохранённые реестры.

    Returns:
        dict — {'success': True, 'registry_id': N, 'created': bool} или {'success': False, 'error': '...'}
    """
    conn = None
    try:
        conn = get_connection()
        _ensure_plp_schema(conn)
        cur = conn.cursor()
        month_val = _PLP_MONTH(month)
        period_val = int(period)
        removed_json = json.dumps(sorted({int(r) for r in (removed_rows or [])}))
        created = False

        reg_id = int(registry_id) if registry_id else None
        if reg_id is None:
            # Новая загрузка периода — отдельная партия
            cur.execute(
                """
                INSERT INTO payment_limits_registry
                    (limit_month, period, file_name, sheet_name, file_bytes, row_count)
                VALUES (%s, %s, COALESCE(%s, ''), COALESCE(%s, ''), %s, COALESCE(%s, 0))
                RETURNING id
                """,
                (month_val, period_val, file_name or '', sheet_name or '',
                 psycopg2.Binary(file_bytes) if file_bytes else None, row_count)
            )
            reg_id = cur.fetchone()[0]
            created = True

        # Исходный файл сохраняем/обновляем в рамках этой загрузки
        if file_bytes:
            cur.execute(
                """
                UPDATE payment_limits_registry
                   SET file_bytes = %s,
                       file_name  = COALESCE(NULLIF(%s, ''), file_name),
                       sheet_name = COALESCE(NULLIF(%s, ''), sheet_name),
                       row_count  = COALESCE(%s, row_count),
                       updated_at = CURRENT_TIMESTAMP
                 WHERE id = %s
                """,
                (psycopg2.Binary(file_bytes), file_name or '', sheet_name or '', row_count, reg_id)
            )

        if stage == PLP_STAGE_DRAFT:
            cur.execute(
                """
                UPDATE payment_limits_registry
                   SET draft_removed_rows = %s::jsonb,
                       draft_saved_at = CURRENT_TIMESTAMP,
                       draft_author = %s,
                       updated_at = CURRENT_TIMESTAMP
                 WHERE id = %s
                """,
                (removed_json, author or '', reg_id)
            )
        else:
            cur.execute(
                """
                UPDATE payment_limits_registry
                   SET final_removed_rows = %s::jsonb,
                       final_saved_at = CURRENT_TIMESTAMP,
                       final_author = %s,
                       updated_at = CURRENT_TIMESTAMP
                 WHERE id = %s
                """,
                (removed_json, author or '', reg_id)
            )

        conn.commit()
        cur.close()
        logger.info(f"Реестр лимитов сохранён: месяц {month_val}, P{period_val}, загрузка #{reg_id}"
                    f"{' (новая)' if created else ''}, стадия «{stage}», "
                    f"снято строк {len(json.loads(removed_json))}, файл {'обновлён' if file_bytes else 'не менялся'}")
        return {'success': True, 'registry_id': reg_id, 'created': created}
    except Exception as e:
        logger.error(f"Ошибка сохранения реестра лимитов: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def get_payment_limits_registry(month, period=None):
    """Метаданные реестра(ов) за месяц (без байтов файла)."""
    conn = None
    try:
        conn = get_connection()
        _ensure_plp_schema(conn)
        cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
        if period:
            cur.execute(_PLP_REGISTRY_META_SQL + " WHERE limit_month = %s AND period = %s ORDER BY period, id",
                        (_PLP_MONTH(month), int(period)))
        else:
            cur.execute(_PLP_REGISTRY_META_SQL + " WHERE limit_month = %s ORDER BY period, id",
                        (_PLP_MONTH(month),))
        rows = cur.fetchall()
        cur.close()
        return [dict(r) for r in rows]
    except Exception as e:
        logger.error(f"Ошибка получения реестра лимитов за {month}: {e}")
        return []
    finally:
        if conn:
            return_connection(conn)


def get_payment_limits_registry_by_id(registry_id):
    """Метаданные одного реестра по id (без байтов файла)."""
    conn = None
    try:
        conn = get_connection()
        _ensure_plp_schema(conn)
        cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
        cur.execute(_PLP_REGISTRY_META_SQL + " WHERE id = %s", (int(registry_id),))
        row = cur.fetchone()
        cur.close()
        return dict(row) if row else None
    except Exception as e:
        logger.error(f"Ошибка получения реестра лимитов #{registry_id}: {e}")
        return None
    finally:
        if conn:
            return_connection(conn)


def mark_payment_limits_registry_exported(registry_id):
    """
    Отмечает, что финальный файл сохранённого реестра выгружен.
    Используется для уведомления (красный кружок) на кнопке «Экспорт в Excel».
    """
    conn = None
    try:
        conn = get_connection()
        _ensure_plp_schema(conn)
        cur = conn.cursor()
        cur.execute(
            "UPDATE payment_limits_registry SET final_exported_at = CURRENT_TIMESTAMP "
            "WHERE id = %s RETURNING final_exported_at",
            (int(registry_id),)
        )
        row = cur.fetchone()
        conn.commit()
        cur.close()
        if not row:
            return {'success': False, 'error': 'Реестр не найден'}
        logger.info(f"Реестр лимитов #{registry_id}: финальный файл выгружен")
        return {'success': True,
                'final_exported_at': row[0].isoformat() if row[0] else None}
    except Exception as e:
        logger.error(f"Ошибка отметки выгрузки реестра {registry_id}: {e}")
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def get_payment_limits_registry_file(registry_id):
    """Исходный файл реестра: (file_name, bytes) либо (None, None)."""
    conn = None
    try:
        conn = get_connection()
        _ensure_plp_schema(conn)
        cur = conn.cursor()
        cur.execute("SELECT file_name, file_bytes FROM payment_limits_registry WHERE id = %s",
                    (int(registry_id),))
        row = cur.fetchone()
        cur.close()
        if not row or row[1] is None:
            return None, None
        return row[0], bytes(row[1])
    except Exception as e:
        logger.error(f"Ошибка получения файла реестра {registry_id}: {e}")
        return None, None
    finally:
        if conn:
            return_connection(conn)


def reset_payment_limits(month, period=None, stage=None):
    """
    Обнуляет сохранённые данные контроля лимитов за месяц: удаляет платежи
    (при необходимости — только выбранной стадии) и реестр вместе с исходным файлом.

    Args:
        month: 'YYYY-MM' или 'YYYY-MM-01'
        period: 1..4 — ограничить обнуление одним периодом (необязательно)
        stage: 'draft' | 'final' — удалить только эту стадию (необязательно)

    Returns:
        dict — {'success': True, 'payments_deleted': N, 'registries_deleted': M}
    """
    conn = None
    try:
        conn = get_connection()
        _ensure_plp_schema(conn)
        cur = conn.cursor()
        month_val = _PLP_MONTH(month)

        sql = "DELETE FROM payment_limits_payments WHERE limit_month = %s"
        params = [month_val]
        if period:
            sql += " AND period = %s"
            params.append(int(period))
        if stage:
            sql += " AND stage = %s"
            params.append(stage)
        cur.execute(sql, tuple(params))
        payments_deleted = cur.rowcount

        sql = "DELETE FROM payment_limits_registry WHERE limit_month = %s"
        params = [month_val]
        if period:
            sql += " AND period = %s"
            params.append(int(period))
        if stage:
            # Обнуляем только выбранную стадию: исходный файл нужен другой стадии, поэтому строку реестра сохраняем
            upd = (f"UPDATE payment_limits_registry "
                   f"SET {stage}_removed_rows = '[]'::jsonb, {stage}_saved_at = NULL, "
                   f"{stage}_author = '', updated_at = CURRENT_TIMESTAMP WHERE limit_month = %s")
            upd_params = [month_val]
            if period:
                upd += " AND period = %s"
                upd_params.append(int(period))
            cur.execute(upd, tuple(upd_params))
            registries_deleted = 0
        else:
            cur.execute(sql, tuple(params))
            registries_deleted = cur.rowcount

        conn.commit()
        cur.close()
        logger.info(f"Обнуление лимитов платежей: месяц {month_val}"
                    + (f", P{period}" if period else "")
                    + (f", стадия «{stage}»" if stage else "")
                    + f" — удалено платежей {payments_deleted}, реестров {registries_deleted}")
        return {'success': True, 'payments_deleted': payments_deleted,
                'registries_deleted': registries_deleted}
    except Exception as e:
        logger.error(f"Ошибка обнуления лимитов платежей: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def cleanup_payment_limits_before(month):
    """
    Автоочистка: удаляет подтверждённые платежи за ВСЕ месяцы строго раньше
    указанного (начало нового месяца — учёт начинается заново).

    Args:
        month: str/date — 'YYYY-MM' или 'YYYY-MM-01'

    Returns:
        int — количество удалённых строк
    """
    conn = None
    try:
        conn = get_connection()
        _ensure_plp_schema(conn)
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM payment_limits_payments WHERE limit_month < %s",
            (_PLP_MONTH(month),)
        )
        deleted = cur.rowcount
        # Реестры (исходные файлы) прошлых месяцев тоже обнуляем
        cur.execute(
            "DELETE FROM payment_limits_registry WHERE limit_month < %s",
            (_PLP_MONTH(month),)
        )
        deleted_registries = cur.rowcount
        conn.commit()
        cur.close()
        if deleted or deleted_registries:
            logger.info(f"Автоочистка лимитов платежей: удалено {deleted} строк и "
                        f"{deleted_registries} реестр(ов) до месяца {_PLP_MONTH(month)}")
        return deleted
    except Exception as e:
        logger.error(f"Ошибка автоочистки лимитов платежей: {e}")
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return 0
    finally:
        if conn:
            return_connection(conn)


def get_payment_limits_config():
    """
    Возвращает конфигурацию лимитов платежей (JSONB).
    Если настройки ещё не сохранялись — значения по умолчанию.
    """
    conn = None
    try:
        conn = get_connection()
        _ensure_plp_schema(conn)
        cur = conn.cursor()
        cur.execute("SELECT config FROM payment_limits_config WHERE id = 1")
        row = cur.fetchone()
        cur.close()
        if row and row[0]:
            return row[0]
        return _copy_default_plp_config()
    except Exception as e:
        logger.error(f"Ошибка получения настроек лимитов: {e}")
        return _copy_default_plp_config()
    finally:
        if conn:
            return_connection(conn)


def _rename_payment_limits_data(cur, old_config, new_config):
    """
    Обновляет подтверждённые платежи при переименовании категорий/подстатей
    (сопоставление старых и новых названий по id).
    """
    renames = []
    if not old_config:
        return renames
    old_by_id = {c.get('id'): c for c in (old_config.get('categories') or [])}
    for cat in (new_config.get('categories') or []):
        old_cat = old_by_id.get(cat.get('id'))
        if not old_cat:
            continue
        if old_cat.get('name') and cat.get('name') and old_cat['name'] != cat['name']:
            cur.execute(
                "UPDATE payment_limits_payments SET article = %s WHERE article = %s",
                (cat['name'], old_cat['name'])
            )
            renames.append({'kind': 'category', 'from': old_cat['name'], 'to': cat['name'], 'rows': cur.rowcount})
        old_details = {d.get('id'): d for d in (old_cat.get('details') or [])}
        for det in (cat.get('details') or []):
            old_det = old_details.get(det.get('id'))
            if not old_det:
                continue
            if old_det.get('name') and det.get('name') and old_det['name'] != det['name']:
                cur.execute(
                    "UPDATE payment_limits_payments SET detail_limit = %s WHERE detail_limit = %s",
                    (det['name'], old_det['name'])
                )
                renames.append({'kind': 'detail', 'from': old_det['name'], 'to': det['name'], 'rows': cur.rowcount})
    return renames


def save_payment_limits_config(config):
    """
    Сохраняет конфигурацию лимитов (одна строка id=1).
    При переименовании категорий/подстатей обновляет уже подтверждённые платежи.

    Returns:
        dict — {'success': True, 'config': ..., 'renames': [...]}
    """
    conn = None
    try:
        if not isinstance(config, dict) or not isinstance(config.get('categories'), list):
            return {'success': False, 'error': 'Некорректная структура настроек лимитов'}

        conn = get_connection()
        _ensure_plp_schema(conn)
        cur = conn.cursor()

        cur.execute("SELECT config FROM payment_limits_config WHERE id = 1")
        row = cur.fetchone()
        old_config = row[0] if row else None

        cur.execute(
            """
            INSERT INTO payment_limits_config (id, config, updated_at)
            VALUES (1, %s, NOW())
            ON CONFLICT (id) DO UPDATE SET config = EXCLUDED.config, updated_at = NOW()
            """,
            (psycopg2.extras.Json(config),)
        )

        renames = _rename_payment_limits_data(cur, old_config, config)

        conn.commit()
        cur.close()
        logger.info(f"Настройки лимитов платежей сохранены (переименований: {len(renames)})")
        return {'success': True, 'config': config, 'renames': renames}
    except Exception as e:
        logger.error(f"Ошибка сохранения настроек лимитов: {e}", exc_info=True)
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        return {'success': False, 'error': str(e)}
    finally:
        if conn:
            return_connection(conn)


def check_health():
    """
    Проверяет доступность БД.
    Возвращает True если БД доступна, иначе False.
    """
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("SELECT 1")
        cur.close()
        return True
    except Exception as e:
        logger.error(f"Health check не пройден: {e}")
        return False
    finally:
        if conn:
            return_connection(conn)