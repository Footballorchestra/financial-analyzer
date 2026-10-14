from flask import Flask, request, send_file, jsonify, send_from_directory
from flask_cors import CORS
import openpyxl
import openpyxl.utils
from openpyxl.styles import numbers, Alignment, Font, PatternFill, Border, Side
from openpyxl.styles.cell_style import StyleArray
from openpyxl.worksheet.views import Pane
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.merge import MergedCellRange
import io
import base64
import re
import zipfile
from datetime import datetime
import json
import traceback
import os
import logging
import gc
from logging.handlers import RotatingFileHandler
import urllib.parse
import uuid
import db
import allocator
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

# ============================================================
# НАСТРОЙКА ЛОГГИРОВАНИЯ В ФАЙЛЫ
# ============================================================
LOG_DIR = '/var/log/financial-analyzer'
os.makedirs(LOG_DIR, exist_ok=True)

# Корневой логгер
root_logger = logging.getLogger()
root_logger.setLevel(logging.INFO)

# Формат логов
log_format = logging.Formatter(
    '[%(asctime)s] %(levelname)-8s %(name)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Ротируемый файл (10 МБ, 5 файлов)
file_handler = RotatingFileHandler(
    os.path.join(LOG_DIR, 'app.log'),
    maxBytes=10 * 1024 * 1024,
    backupCount=5,
    encoding='utf-8'
)
file_handler.setFormatter(log_format)
root_logger.addHandler(file_handler)

# Также пишем в stdout (для gunicorn/systemd)
console_handler = logging.StreamHandler()
console_handler.setFormatter(log_format)
root_logger.addHandler(console_handler)

logger = logging.getLogger('financial_analyzer')

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB max file size

# ============================================================
# RATE LIMITING
# ============================================================
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["60 per minute"],
    storage_uri="memory://"
)

# ============================================================
# ИСПРАВЛЕНИЕ CORS - разрешаем запросы с любых источников
# ВАЖНО: expose_headers должен включать 'X-Filial-Data' для передачи данных филиалов
# ============================================================
CORS(app,
     expose_headers=['X-Filial-Data', 'X-Allocation-Summary', 'X-Email-Data',
                     'X-Export-Removed', 'X-Export-Warning', 'X-Export-Tabs'],
     resources={r'/*': {'origins': '*'}})

# ============================================================
# НАСТРОЙКИ КОДИРОВКИ ДЛЯ КИРИЛЛИЦЫ
# ============================================================
app.config['JSON_AS_ASCII'] = False
app.config['JSON_SORT_KEYS'] = False

# ============================================================
# МАРШРУТЫ ДЛЯ РАЗДАЧИ СТАТИКИ (фронтенд через HTTP)
# ============================================================
BASE_DIR = '/opt/financial-analyzer'

@app.route('/')
def serve_index():
    return send_from_directory(BASE_DIR, 'index.html')

@app.route('/lib/<path:filename>')
def serve_lib(filename):
    return send_from_directory(os.path.join(BASE_DIR, 'lib'), filename)

@app.route('/<path:filename>')
def serve_static(filename):
    if '..' in filename or filename.startswith('/'):
        return jsonify({'error': 'Invalid path'}), 400
    return send_from_directory(BASE_DIR, filename)

# ============================================================
# ГЛОБАЛЬНЫЙ ОБРАБОТЧИК ДЛЯ УСТАНОВКИ КОДИРОВКИ
# ============================================================
@app.after_request
def add_charset_header(response):
    if response.content_type and response.content_type.startswith('application/json'):
        response.headers['Content-Type'] = 'application/json; charset=utf-8'
    return response

# ============================================================
# HEALTH CHECK ЭНДПОИНТ
# ============================================================
@app.route('/api/health', methods=['GET'])
def api_health():
    """Проверка работоспособности сервера и доступности БД."""
    db_ok = db.check_health()
    status_code = 200 if db_ok else 503
    return jsonify({
        'status': 'ok' if db_ok else 'degraded',
        'database': 'connected' if db_ok else 'unavailable',
        'timestamp': datetime.now().isoformat()
    }), status_code

# Список целевых контрагентов (используется только для фильтрации документов)
TARGET_CONTRAGENTS = ['ВАНКОРНЕФТЬ АО', 'РН-Ванкор ООО']

# Колонки (1‑индексация Excel)
COLUMNS = {
    'DOCUMENT_NAME': 1,      # A
    'DEBT_AMOUNT': 12,       # L
    'UNSIGNED_DEBT': 13,     # M
    'OVERDUE': 15,           # O
    'UNSIGNED_OVERDUE': 16,  # P
    'DAYS': 18,              # R
    'OUR_DEBT': 19,          # S
    'NOT_OVERDUE': 20,       # T
    'INTERVAL_1_15': 21,     # U
    'INTERVAL_16_29': 22,    # V
    'INTERVAL_30_89': 23,    # W
    'INTERVAL_90_179': 24,   # X
    'INTERVAL_180_PLUS': 25, # Y
}

SUM_COLUMNS = [
    COLUMNS['DEBT_AMOUNT'], COLUMNS['UNSIGNED_DEBT'],
    COLUMNS['OVERDUE'], COLUMNS['UNSIGNED_OVERDUE'],
    COLUMNS['NOT_OVERDUE'], COLUMNS['INTERVAL_1_15'],
    COLUMNS['INTERVAL_16_29'], COLUMNS['INTERVAL_30_89'],
    COLUMNS['INTERVAL_90_179'], COLUMNS['INTERVAL_180_PLUS'],
]

NUMERIC_COLUMNS = SUM_COLUMNS + [COLUMNS['DAYS']]

# ============================================================
# КЭШ ОБЪЕДИНЁННЫХ ЯЧЕЕК
# is_cell_merged()/get_cell_to_write() вызываются десятки тысяч раз за один
# запрос /save-excel. Линейный обход ws.merged_cells.ranges на каждый вызов
# делал пересчёт итогов и сортировку подразделений квадратичными
# (O(строки × колонки × объединения)) — на реальном реестре это давало
# десятки секунд. Кэшируем карту «координата → якорь объединения» один раз
# на лист.
# ============================================================
_MERGED_CELLS_CACHE_ATTR = '_fa_merged_cells_map'


def _invalidate_merged_cells_cache(ws):
    """Сбрасывает кэш объединённых ячеек листа (вызывать после merge/unmerge)."""
    try:
        delattr(ws, _MERGED_CELLS_CACHE_ATTR)
    except AttributeError:
        pass


def get_merged_cells_map(ws):
    """Карта {(row, col): (anchor_row, anchor_col)} для объединённых ячеек.

    Необъединённые ячейки в карте отсутствуют. Строится один раз на лист и
    переиспользуется всеми операциями чтения/записи.
    """
    cached = getattr(ws, _MERGED_CELLS_CACHE_ATTR, None)
    if cached is not None:
        return cached
    merged = {}
    for merged_range in ws.merged_cells.ranges:
        anchor = (merged_range.min_row, merged_range.min_col)
        for r in range(merged_range.min_row, merged_range.max_row + 1):
            for c in range(merged_range.min_col, merged_range.max_col + 1):
                merged[(r, c)] = anchor
    setattr(ws, _MERGED_CELLS_CACHE_ATTR, merged)
    return merged


def is_cell_merged(ws, row, col):
    return (row, col) in get_merged_cells_map(ws)

def get_cell_to_write(ws, row, col):
    anchor = get_merged_cells_map(ws).get((row, col))
    if anchor is None:
        return ws.cell(row=row, column=col)
    return ws.cell(row=anchor[0], column=anchor[1])


# Компоненты стиля ячейки: (поле StyleArray, коллекция в книге, свойство ячейки)
_STYLE_COLLECTIONS = (
    ('fontId', '_fonts', 'font'),
    ('fillId', '_fills', 'fill'),
    ('borderId', '_borders', 'border'),
    ('protectionId', '_protections', 'protection'),
    ('alignmentId', '_alignments', 'alignment'),
)


def copy_cell_style(src_cell, dst_cell):
    """Переносит стиль ячейки в ячейку другой книги без XML-копирования.

    openpyxl реализует __copy__ стилей через to_tree/from_tree (~80 мкс на
    объект, 6 объектов на ячейку). Просто скопировать StyleArray нельзя:
    он хранит индексы в таблицах стилей книги (у другой книги это другие
    объекты). Поэтому берём сами объекты стилей из таблиц исходной книги по
    индексам ячейки — целевая книга зарегистрирует их у себя, а IndexedList
    дедуплицирует по значению. Объекты стилей неизменяемы, поэтому их
    совместное использование между книгами безопасно.
    """
    style = src_cell._style
    if style is None:
        return
    src_wb = src_cell.parent.parent
    for key, collection, attr in _STYLE_COLLECTIONS:
        setattr(dst_cell, attr, getattr(src_wb, collection)[getattr(style, key)])
    dst_cell.number_format = src_cell.number_format


def add_merged_range_fast(ws, coord, existing_ranges=None):
    """Добавляет объединение без полного обхода списка объединений (O(M)).

    Повторяет ws.merge_cells(), но вместо O(M)-поиска пересечений кладёт
    диапазон в множество напрямую. Применимо только к заведомо
    непересекающимся диапазонам (перенос уже существующих объединений,
    копирование листа).

    existing_ranges — «оставшиеся на месте» объединения, с которыми нужно
    свериться на вложенность (их обычно мало: шапка, строки за пределами
    сортировки). Для копирования листа не нужен — диапазоны исходного листа
    взаимно непересекающиеся.
    """
    mcr = MergedCellRange(ws, coord)
    if existing_ranges and any(mcr.issubset(other) for other in existing_ranges):
        # как ws.merge_cells(): диапазон уже входит в существующее объединение
        pass
    else:
        ws.merged_cells.ranges.add(mcr)
    ws._clean_merge_range(mcr)
    return mcr

def safe_set_number_format(ws, row, col, value):
    if row <= 13:
        return
    cell = get_cell_to_write(ws, row, col)
    cell.value = value
    cell.number_format = '#,##0.00'
    cell.alignment = Alignment(horizontal='right')

def safe_set_value(ws, row, col, value):
    if row <= 13:
        return
    cell = get_cell_to_write(ws, row, col)
    cell.value = value
    cell.alignment = Alignment(horizontal='right')

def align_numeric_cells(ws):
    logger.info("Выравнивание числовых ячеек по правому краю...")
    # Общий кэш объединённых ячеек: O(1) проверка вместо O(merges) на ячейку
    merged_map = get_merged_cells_map(ws)
    for row in range(14, ws.max_row + 1):
        for col in NUMERIC_COLUMNS:
            if (row, col) in merged_map:
                continue
            cell = ws.cell(row=row, column=col)
            if cell.value is not None:
                cell.alignment = Alignment(horizontal='right')

def _row_to_float(value):
    """Приводит значение ячейки к числу (0, если не число)."""
    if value is None or isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        cleaned = str(value).replace('\xa0', '').replace(' ', '').replace(',', '.')
        return float(cleaned)
    except (TypeError, ValueError):
        return 0.0


def sort_filial_blocks_by_overdue(ws):
    """Сортирует блоки подразделений (строки «ДТ …» с отступом 0) листа реестра
    по убыванию просроченной задолженности (колонка O / 15).

    Блок перемещается целиком: строка подразделения + все вложенные строки
    (контрагенты/договоры/документы), сохраняя их порядок, форматирование,
    объединённые ячейки и высоты строк. Сортировка стабильная: при равных
    суммах блоки остаются в исходном порядке. Строка «Итого», шапка и строки
    после «Итого» не сортируются.
    """
    logger.info("=== СОРТИРОВКА ПОДРАЗДЕЛЕНИЙ ПО УБЫВАНИЮ ПРОСРОЧКИ ===")
    start = 14
    nodes, total_row, filial_rows = build_hierarchy(ws, start_row=start)
    if not filial_rows:
        logger.warning("Подразделения не найдены — сортировка пропущена")
        return

    if total_row is not None:
        data_end = total_row
    else:
        data_end = start
        # ws.max_row/ws.max_column — «дорогие» свойства (сканируют все ячейки),
        # поэтому вычисляем их один раз до циклов.
        max_row_limit = ws.max_row
        max_col_limit = min(ws.max_column, 30)
        for r in range(start, max_row_limit + 1):
            for c in range(1, max_col_limit + 1):
                if ws.cell(row=r, column=c).value is not None:
                    data_end = r + 1

    # Собираем блоки: все строки от первого подразделения до data_end принадлежат
    # ровно одному блоку. Новая строка «ДТ …» (отступ 0) открывает следующий блок.
    blocks = []
    current = None
    for r in range(start, data_end):
        cell_val = ws.cell(row=r, column=1).value
        str_val = str(cell_val).strip() if cell_val is not None else ''
        if str_val:
            indent = get_cell_indent(ws, r)
            if indent == 0 and str_val.startswith('ДТ '):
                if current is not None:
                    current['end'] = r
                    blocks.append(current)
                current = {'start': r, 'end': r + 1, 'filial_row': r}
                continue
        # вложенные строки и пустые строки прикрепляются к текущему блоку
        if current is not None:
            current['end'] = r + 1

    if current is not None:
        blocks.append(current)

    if len(blocks) < 2:
        logger.info("Меньше двух блоков подразделений — сортировка не требуется")
        return

    sort_start = blocks[0]['start']
    sort_end = max(b['end'] for b in blocks)
    blocks[-1]['end'] = sort_end

    # Ключ сортировки — просроченная задолженность в строке подразделения после пересчёта
    for b in blocks:
        b['key'] = _row_to_float(get_cell_value(ws, b['filial_row'], COLUMNS['OVERDUE']))
    sorted_blocks = sorted(blocks, key=lambda b: b['key'], reverse=True)

    # Новый порядок физических строк (перемещаются целые блоки)
    old_rows = [r for b in blocks for r in range(b['start'], b['end'])]
    new_rows = [r for b in sorted_blocks for r in range(b['start'], b['end'])]
    if sorted(old_rows) != list(range(sort_start, sort_end)):
        logger.warning("Блоки не покрывают диапазон строк — сортировка пропущена")
        return
    # ВАЖНО: соответствие «старая строка -> новая строка» нельзя строить через
    # zip(old_rows, new_rows): это ставит в соответствие старой строке «старый номер»
    # другой строки, а не номер новой позиции (ломается при блоках разного размера).
    # Правильно: содержимое строки row (которая в новом порядке стоит на позиции p)
    # должно быть записано на строку (sort_start + p).
    row_map = {row: sort_start + p for p, row in enumerate(new_rows)}

    # Снимок содержимого строк и высот ДО очистки.
    # ws.max_column вычисляется один раз: внутри цикла по строкам это свойство
    # каждый раз сканирует все ячейки листа (O(rows × cols²) на больших файлах).
    cells_snapshot = []
    max_col = ws.max_column
    for r in old_rows:
        row_cells = []
        for c in range(1, max_col + 1):
            if (r, c) in ws._cells:
                row_cells.append((c, ws._cells[(r, c)]))
        cells_snapshot.append((r, row_cells))

    row_props = {}
    for r in old_rows:
        if r in ws.row_dimensions:
            rd = ws.row_dimensions[r]
            row_props[r] = (rd.height, rd.hidden, rd.outlineLevel, rd.collapsed)

    # Объединённые ячейки внутри сортируемого диапазона (переносятся вместе со строками).
    # Переносим только те, что целиком лежат в одном блоке (одинаковый сдвиг обеих границ).
    merge_map = []
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= sort_start and rng.max_row < sort_end:
            d_min = row_map[rng.min_row] - rng.min_row
            d_max = row_map[rng.max_row] - rng.max_row
            if d_min == d_max:
                merge_map.append((rng, d_min))
            else:
                logger.warning(f"Объединение {rng} пересекает границы блоков — не переносится")
    for rng, _ in merge_map:
        ws.merged_cells.remove(rng)
    # Объединения, оставшиеся на месте (шапка, строки за пределами сортировки).
    # Их немного, и по ним проверяется вложенность переносимых диапазонов —
    # чтобы поведение совпадало с ws.merge_cells().
    kept_merge_ranges = list(ws.merged_cells.ranges)
    # Объединения изменились — сбрасываем кэш карты ячеек
    _invalidate_merged_cells_cache(ws)

    # Очищаем диапазон: удаляем старые ячейки и свойства строк
    for key in [k for k in list(ws._cells) if sort_start <= k[0] < sort_end]:
        del ws._cells[key]
    for r in range(sort_start, sort_end):
        if r in ws.row_dimensions:
            del ws.row_dimensions[r]

    # Записываем строки в новом порядке (значения, стили, гиперссылки, высоты)
    for old_row, row_cells in cells_snapshot:
        new_row = row_map[old_row]
        for col, src_cell in row_cells:
            dst_cell = ws.cell(row=new_row, column=col)
            dst_cell.value = src_cell.value
            if src_cell.has_style:
                # Копируем весь стиль одним массивом индексов. Прежний вариант
                # (copy() на font/border/fill/alignment/protection) для каждого
                # стиля делал XML-сериализацию и обратный разбор (~80 мкс),
                # т.е. ~0.5 мс на ячейку. StyleArray живёт в той же книге,
                # поэтому индексы стилей валидны.
                dst_cell._style = StyleArray(src_cell._style)
            if src_cell.hyperlink:
                dst_cell.hyperlink = src_cell.hyperlink
        if old_row in row_props:
            height, hidden, outline, collapsed = row_props[old_row]
            rd = ws.row_dimensions[new_row]
            rd.height = height
            rd.hidden = hidden
            rd.outlineLevel = outline
            rd.collapsed = collapsed

    # Восстанавливаем объединённые ячейки на новых местах.
    # ws.merge_cells() для каждого диапазона выполняет проверку пересечений O(M),
    # что даёт O(M²) на лист. Диапазоны заведомо не пересекаются (они были
    # корректными объединениями и переносятся целиком внутри своих блоков),
    # поэтому используем add_merged_range_fast без квадратичной проверки.
    for rng, delta in merge_map:
        new_coord = CellRange(
            min_row=rng.min_row + delta,
            min_col=rng.min_col,
            max_row=rng.max_row + delta,
            max_col=rng.max_col,
        ).coord
        add_merged_range_fast(ws, new_coord, existing_ranges=kept_merge_ranges)
    # Позиции объединений изменились — кэш карты ячеек больше не актуален
    _invalidate_merged_cells_cache(ws)

    # Диагностика: фактический порядок строк подразделений после перестановки
    order_lines = []
    for r in range(sort_start, sort_end):
        cell_val = ws.cell(row=r, column=1).value
        if cell_val and str(cell_val).strip().startswith('ДТ '):
            order_lines.append(
                f"{str(cell_val).strip()}={_row_to_float(get_cell_value(ws, r, COLUMNS['OVERDUE'])):,.2f}"
            )
    if order_lines:
        logger.info("Порядок в файле после сортировки: " + " | ".join(order_lines))

    logger.info(
        f"Отсортировано блоков подразделений: {len(blocks)} "
        f"(строки {sort_start}–{sort_end - 1}) по убыванию просрочки"
    )


def get_cell_value(ws, row, col):
    if is_cell_merged(ws, row, col):
        cell = get_cell_to_write(ws, row, col)
        return cell.value
    return ws.cell(row=row, column=col).value

def safe_set_top_table_value(ws, row, col, value):
    cell = get_cell_to_write(ws, row, col)
    cell.value = value
    cell.number_format = '#,##0.00' if col == 5 else '0.00%'
    cell.alignment = Alignment(horizontal='right')

def get_interval_col(days):
    if days <= 0:
        return COLUMNS['NOT_OVERDUE']
    elif 1 <= days <= 15:
        return COLUMNS['INTERVAL_1_15']
    elif 16 <= days <= 29:
        return COLUMNS['INTERVAL_16_29']
    elif 30 <= days <= 89:
        return COLUMNS['INTERVAL_30_89']
    elif 90 <= days <= 179:
        return COLUMNS['INTERVAL_90_179']
    else:
        return COLUMNS['INTERVAL_180_PLUS']

def clear_all_intervals(ws, row):
    if row <= 13:
        return
    interval_cols = [
        COLUMNS['NOT_OVERDUE'], COLUMNS['INTERVAL_1_15'],
        COLUMNS['INTERVAL_16_29'], COLUMNS['INTERVAL_30_89'],
        COLUMNS['INTERVAL_90_179'], COLUMNS['INTERVAL_180_PLUS'],
    ]
    for col in interval_cols:
        safe_set_number_format(ws, row, col, 0)

DOCUMENT_KEYWORDS = [
    'Акт', 'Реализация', 'Корректировка', 'Поступление',
    'Взаимозачет', 'Взаимозачёт', 'Списание', 'УПД', 'Счет-фактура',
    'Товарная накладная', 'ТОРГ-12', 'Универсальный передаточный'
]

def find_structure(ws):
    filials = []
    kontragents = []
    dogovors = []
    documents = []
    total_row = None
    for row in range(14, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=1).value
        if not cell_value:
            continue
        str_val = str(cell_value).strip()
        if str_val.startswith('ДТ '):
            filials.append(row)
        elif 'Итого' in str_val or 'ИТОГО' in str_val:
            total_row = row
        elif any(keyword in str_val for keyword in DOCUMENT_KEYWORDS):
            documents.append(row)
        elif str_val.startswith('Договор') or (str_val.startswith('договор') and not any(kw in str_val for kw in DOCUMENT_KEYWORDS)):
            dogovors.append(row)
        else:
            if len(str_val) > 2 and not str_val[0].isdigit():
                kontragents.append(row)
    return filials, kontragents, dogovors, documents, total_row

def get_cell_indent(ws, row, col=1):
    cell = ws.cell(row=row, column=col)
    if cell.alignment and cell.alignment.indent is not None:
        return int(cell.alignment.indent)
    return 0

def build_hierarchy(ws, start_row=14):
    nodes = {}
    stack = []
    total_row = None
    filial_rows = []
    for row in range(start_row, ws.max_row + 1):
        cell = ws.cell(row=row, column=1)
        value = cell.value
        if value is None:
            continue
        str_val = str(value).strip()
        if not str_val:
            continue
        if 'Итого' in str_val or 'ИТОГО' in str_val:
            total_row = row
            continue
        indent = get_cell_indent(ws, row, 1)
        nodes[row] = {'indent': indent, 'children': [], 'parent': None}
        if indent == 0 and str_val.startswith('ДТ '):
            filial_rows.append(row)
        while stack and stack[-1][0] >= indent:
            stack.pop()
        if stack:
            parent_row = stack[-1][1]
            nodes[row]['parent'] = parent_row
            nodes[parent_row]['children'].append(row)
        stack.append((indent, row))
    return nodes, total_row, filial_rows

def recalc_totals(ws):
    logger.info("=== ПЕРЕСЧЁТ ИТОГОВ (по отступам, только листья) ===")
    nodes, total_row, filial_rows = build_hierarchy(ws)
    if not nodes:
        logger.warning("Не найдено строк с данными для пересчёта")
        return
    logger.info(f"Найдено узлов: {len(nodes)}, филиалов: {len(filial_rows)}, итого: строка {total_row}")
    leaves = set()
    for row, node in nodes.items():
        if not node['children']:
            leaves.add(row)
    logger.info(f"Листьев: {len(leaves)}")

    # Агрегаты считаются один раз — снизу вверх. Листья читаются из файла, а
    # родительские строки суммируют уже готовые суммы детей.
    # Прежняя реализация для КАЖДОГО узла заново обходила всё его поддерево
    # листьев (и на каждое чтение ячейки линейно искала объединения), что на
    # реестре ~3500 строк занимало десятки секунд.
    sums = {}         # row -> {col: сумма}
    max_days = {}     # row -> максимум по колонке «дней»
    leaf_counts = {}  # row -> количество листьев в поддереве

    for row in leaves:
        sums[row] = {col: _row_to_float(get_cell_value(ws, row, col)) for col in SUM_COLUMNS}
        max_days[row] = _row_to_float(get_cell_value(ws, row, COLUMNS['DAYS']))
        leaf_counts[row] = 1

    # Обход в порядке убывания строк: дети всегда имеют номер больше родителя,
    # поэтому к моменту обработки узла агрегаты всех его детей уже посчитаны.
    sorted_rows = sorted(nodes.keys(), reverse=True)
    recalculated_count = 0
    for row in sorted_rows:
        node = nodes[row]
        if not node['children']:
            continue
        recalculated_count += 1
        agg = {col: 0 for col in SUM_COLUMNS}
        max_day = 0
        leaf_count = 0
        for child in node['children']:
            child_sums = sums.get(child)
            if child_sums:
                for col in SUM_COLUMNS:
                    agg[col] += child_sums[col]
            child_days = max_days.get(child, 0)
            if child_days > max_day:
                max_day = child_days
            leaf_count += leaf_counts.get(child, 0)
        sums[row] = agg
        max_days[row] = max_day
        leaf_counts[row] = leaf_count
        # Округляем до копеек: суммы — денежные величины, формат '#,##0.00'
        for col in SUM_COLUMNS:
            safe_set_number_format(ws, row, col, round(agg[col], 2))
        safe_set_value(ws, row, COLUMNS['DAYS'], max_day)
    logger.info(f"Пересчитано узлов: {recalculated_count}")

    if total_row and filial_rows:
        agg = {col: 0 for col in SUM_COLUMNS}
        max_day = 0
        leaf_count = 0
        for fil_row in filial_rows:
            fil_sums = sums.get(fil_row)
            if fil_sums:
                for col in SUM_COLUMNS:
                    agg[col] += fil_sums[col]
            fil_days = max_days.get(fil_row, 0)
            if fil_days > max_day:
                max_day = fil_days
            leaf_count += leaf_counts.get(fil_row, 0)
        logger.info(f"Общий итог стр.{total_row}: {leaf_count} листьев под филиалами")
        for col in SUM_COLUMNS:
            safe_set_number_format(ws, total_row, col, round(agg[col], 2))
        safe_set_value(ws, total_row, COLUMNS['DAYS'], max_day)
    logger.info("=== ПЕРЕСЧЁТ ИТОГОВ ЗАВЕРШЁН ===")

def update_top_table(ws, total_row):
    logger.info("=== ОБНОВЛЕНИЕ ВЕРХНЕЙ ТАБЛИЦЫ ===")
    if not total_row:
        logger.warning("Не найдена итоговая строка")
        return
    t_value = get_cell_value(ws, total_row, COLUMNS['NOT_OVERDUE']) or 0
    u_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_1_15']) or 0
    v_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_16_29']) or 0
    w_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_30_89']) or 0
    x_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_90_179']) or 0
    y_value = get_cell_value(ws, total_row, COLUMNS['INTERVAL_180_PLUS']) or 0
    l_value = get_cell_value(ws, total_row, COLUMNS['DEBT_AMOUNT']) or 0
    safe_set_top_table_value(ws, 2, 5, t_value)
    safe_set_top_table_value(ws, 2, 6, t_value / l_value if l_value else 0)
    safe_set_top_table_value(ws, 3, 5, u_value)
    safe_set_top_table_value(ws, 3, 6, u_value / l_value if l_value else 0)
    safe_set_top_table_value(ws, 4, 5, v_value)
    safe_set_top_table_value(ws, 4, 6, v_value / l_value if l_value else 0)
    safe_set_top_table_value(ws, 5, 5, w_value)
    safe_set_top_table_value(ws, 5, 6, w_value / l_value if l_value else 0)
    safe_set_top_table_value(ws, 6, 5, x_value)
    safe_set_top_table_value(ws, 6, 6, x_value / l_value if l_value else 0)
    safe_set_top_table_value(ws, 7, 5, y_value)
    safe_set_top_table_value(ws, 7, 6, y_value / l_value if l_value else 0)
    safe_set_top_table_value(ws, 8, 5, l_value)
    safe_set_top_table_value(ws, 8, 6, 1.0)
    logger.info("Верхняя таблица обновлена")

def copy_worksheet_full(ws, wb):
    logger.info(f"Копирование ячеек листа '{ws.title}'...")
    new_ws = wb.create_sheet(ws.title)
    # Копируем только реально существующие ячейки. ws.iter_rows() создаёт
    # Cell-объекты на всей прямоугольной области (включая пустые) — для
    # широких листов это лишняя работа и мусор в выходном файле.
    for (row, col), cell in ws._cells.items():
        new_cell = new_ws.cell(row=row, column=col, value=cell.value)
        if cell.has_style:
            # Без copy(): openpyxl реализует __copy__ стилей через XML
            # (to_tree + from_tree) — это ~80 мкс на объект и 6 объектов на
            # ячейку.
            copy_cell_style(cell, new_cell)
        if cell.data_type == 'f':
            new_cell.data_type = 'f'
    # Объединения исходного листа непересекающиеся — добавляем их без
    # квадратичной проверки пересечений
    for merged_range in list(ws.merged_cells.ranges):
        add_merged_range_fast(new_ws, str(merged_range))
    for col_letter, col_dim in ws.column_dimensions.items():
        new_ws.column_dimensions[col_letter].width = col_dim.width
        if col_dim.outline_level:
            new_ws.column_dimensions[col_letter].outline_level = col_dim.outline_level
        if col_dim.hidden:
            new_ws.column_dimensions[col_letter].hidden = col_dim.hidden
    for row_num, row_dim in ws.row_dimensions.items():
        new_ws.row_dimensions[row_num].height = row_dim.height
        if row_dim.outline_level:
            new_ws.row_dimensions[row_num].outline_level = row_dim.outline_level
        if row_dim.hidden:
            new_ws.row_dimensions[row_num].hidden = row_dim.hidden
    if hasattr(ws.sheet_properties, 'outlinePr') and ws.sheet_properties.outlinePr:
        new_ws.sheet_properties.outlinePr.summaryBelow = ws.sheet_properties.outlinePr.summaryBelow
        new_ws.sheet_properties.outlinePr.summaryRight = ws.sheet_properties.outlinePr.summaryRight
    if ws.print_options:
        for attr in ['grid_lines', 'grid_lines_set', 'horizontal_centered', 'vertical_centered']:
            if hasattr(ws.print_options, attr):
                setattr(new_ws.print_options, attr, getattr(ws.print_options, attr))
    if ws.page_setup:
        for attr in ['orientation', 'paperSize', 'scale', 'fitToHeight', 'fitToWidth',
                     'pageOrder', 'blackAndWhite', 'draft', 'cellComments', 'errors']:
            if hasattr(ws.page_setup, attr) and getattr(ws.page_setup, attr) is not None:
                setattr(new_ws.page_setup, attr, getattr(ws.page_setup, attr))
    if ws.page_margins:
        for attr in ['left', 'right', 'top', 'bottom', 'header', 'footer']:
            if hasattr(ws.page_margins, attr):
                setattr(ws.page_margins, attr, getattr(ws.page_margins, attr))
    if ws.sheet_view and ws.sheet_view.pane:
        pane = ws.sheet_view.pane
        try:
            new_ws.sheet_view.pane = Pane(
                xSplit=pane.xSplit, ySplit=pane.ySplit,
                topLeftCell=pane.topLeftCell, activePane=pane.activePane, state=pane.state
            )
        except (TypeError, AttributeError):
            logger.warning(f"Пропущена настройка freeze panes (несовместимость версий)")
    if ws.auto_filter.ref:
        new_ws.auto_filter.ref = ws.auto_filter.ref
    for cf_rule in ws.conditional_formatting._cf_rules:
        new_ws.conditional_formatting._cf_rules.append(cf_rule)
    if hasattr(ws, 'tables') and ws.tables:
        for table in ws.tables:
            new_ws.tables.add(table)
    logger.info(f"Лист '{ws.title}' скопирован с полным форматированием")
    return new_ws

# ============================================================
# API: СВЕРКА ДЕБИТОРСКОЙ ЗАДОЛЖЕННОСТИ
# ============================================================
@app.route('/save-excel', methods=['POST'])
@limiter.limit("5 per minute")  # Ограничение для тяжёлой операции
def save_excel():
    try:
        file = request.files['file']
        data = json.loads(request.form['data'])
        mode = request.form.get('mode', '')

        logger.info(f"=== ПОЛУЧЕН ЗАПРОС ===")
        logger.info(f"Режим: {mode or 'обычный'}")
        logger.info(f"Файл: {file.filename}")
        logger.info(f"Документов для обновления: {len(data['updatedDocuments'])}")

        wb = openpyxl.load_workbook(io.BytesIO(file.read()))

        logger.info("=== ШАГ 0: ПЕРЕИМЕНОВАНИЕ ЛИСТОВ ===")
        all_sheets = list(wb.sheetnames)
        logger.info(f"Исходные листы: {all_sheets}")

        siuat_sheets = []
        summary_sheets = []
        main_sheets = []

        for sheet_name in all_sheets:
            if sheet_name.startswith('Свод ДЗ СИ УАТ') or sheet_name.startswith('Свод ДЗ СИ'):
                siuat_sheets.append(sheet_name)
            elif 'Сводные' in sheet_name or 'Сводная' in sheet_name:
                summary_sheets.append(sheet_name)
            else:
                test_ws = wb[sheet_name]
                is_main = False
                for r in range(1, min(30, test_ws.max_row + 1)):
                    cell_val = test_ws.cell(row=r, column=1).value
                    if cell_val and str(cell_val).strip().startswith('ДТ '):
                        is_main = True
                        break
                if is_main:
                    main_sheets.append(sheet_name)

        logger.info(f"main_sheets={main_sheets}, siuat_sheets={siuat_sheets}, summary_sheets={summary_sheets}")

        temp_names = {}
        temp_counter = 0

        for sheet_name in all_sheets:
            if sheet_name.startswith('Свод ДЗ СИ УАТ') or sheet_name.startswith('Свод ДЗ СИ'):
                temp_counter += 1
                temp_names[sheet_name] = f'__temp_siuat_{temp_counter}__'
            elif 'Сводные' in sheet_name or 'Сводная' in sheet_name:
                temp_counter += 1
                temp_names[sheet_name] = f'__temp_summary_{temp_counter}__'

        for idx, sheet_name in enumerate(main_sheets):
            if idx == 0:
                temp_names[sheet_name] = '__temp_main__'
            elif idx == 1:
                temp_names[sheet_name] = '__temp_siuat_from_main__'
            else:
                temp_names[sheet_name] = f'__temp_extra_{idx}__'

        for old_name, temp_name in temp_names.items():
            if old_name != temp_name and old_name in wb.sheetnames:
                wb[old_name].title = temp_name
                logger.info(f"Шаг 1: '{old_name}' → '{temp_name}'")

        for sn in list(wb.sheetnames):
            if sn == '__temp_main__':
                wb[sn].title = 'Свод ДЗ'
                logger.info("Шаг 2: '__temp_main__' → 'Свод ДЗ'")
            elif sn == '__temp_siuat_from_main__':
                wb[sn].title = 'Свод ДЗ СИ УАТ'
                logger.info("Шаг 2: '__temp_siuat_from_main__' → 'Свод ДЗ СИ УАТ'")
            elif sn.startswith('__temp_siuat_'):
                wb[sn].title = 'Свод ДЗ СИ УАТ'
                logger.info(f"Шаг 2: '{sn}' → 'Свод ДЗ СИ УАТ'")
            elif sn.startswith('__temp_summary_'):
                wb[sn].title = 'Сводные таблицы'
                logger.info(f"Шаг 2: '{sn}' → 'Сводные таблицы'")
            elif sn.startswith('__temp_extra_'):
                del wb[sn]
                logger.info(f"Удалён лишний лист: {sn}")

        logger.info(f"Листы после переименования: {wb.sheetnames}")
        logger.info("=== ПЕРЕИМЕНОВАНИЕ ЗАВЕРШЕНО ===")

        ws = wb['Свод ДЗ']
        wb._active_sheet_index = wb.sheetnames.index('Свод ДЗ')

        today = datetime.now().date()
        logger.info(f"Текущая дата: {today}")

        _, _, _, _, total_row = find_structure(ws)

        updated_rows = set()
        docs_with_date = 0
        docs_without_date = 0
        docs_overdue = 0
        docs_not_overdue = 0

        for idx, item in enumerate(data['updatedDocuments']):
            row_number = item['rowNumber']
            debt_amount = float(item['amount'])
            expected_date_str = item['date']
            doc_name = item.get('documentName', 'неизвестно')
            action = item.get('action', '')

            expected_date = None
            if expected_date_str and expected_date_str not in ('null', 'None', ''):
                try:
                    clean_str = expected_date_str.strip().strip('"').strip("'")
                    if clean_str:
                        expected_date = datetime.fromisoformat(clean_str.replace('Z', '+00:00')).date()
                except Exception as e:
                    logger.warning(f"Ошибка парсинга даты: {e}")
                    expected_date = None

            if idx < 10 or expected_date is None:
                status = "БЕЗ ДАТЫ" if expected_date is None else f"дата: {expected_date}"
                logger.info(f"[{idx+1}] Строка {row_number}: '{doc_name[:50]}...' — {status}, сумма: {debt_amount:,.2f}")

            clear_all_intervals(ws, row_number)

            if expected_date is None or expected_date >= today:
                reason = 'нет даты в файле поступлений' if expected_date is None else 'дата в будущем или сегодня'
                logger.info(f"  → НЕ ПРОСРОЧЕНО (причина: {reason})")
                safe_set_number_format(ws, row_number, COLUMNS['OVERDUE'], 0)
                safe_set_value(ws, row_number, COLUMNS['DAYS'], 0)
                safe_set_number_format(ws, row_number, COLUMNS['NOT_OVERDUE'], debt_amount)
                updated_rows.add(row_number)
                docs_not_overdue += 1
                if expected_date is None:
                    docs_without_date += 1
                else:
                    docs_with_date += 1
            elif expected_date < today:
                days_overdue = (today - expected_date).days
                interval_col = get_interval_col(days_overdue)
                logger.info(f"  → ПРОСРОЧЕНО на {days_overdue} дн. (интервал: {interval_col})")
                safe_set_number_format(ws, row_number, COLUMNS['OVERDUE'], debt_amount)
                safe_set_value(ws, row_number, COLUMNS['DAYS'], days_overdue)
                safe_set_number_format(ws, row_number, interval_col, debt_amount)
                updated_rows.add(row_number)
                docs_overdue += 1
                docs_with_date += 1

        if updated_rows:
            logger.info(f"=== СТАТИСТИКА ОБРАБОТКИ ===")
            logger.info(f"Обновлено строк: {len(updated_rows)}")
            logger.info(f"  Документов с датой: {docs_with_date}")
            logger.info(f"  Документов БЕЗ даты: {docs_without_date}")
            logger.info(f"  Просрочено: {docs_overdue}")
            logger.info(f"  Не просрочено: {docs_not_overdue}")
            logger.info("============================")

            recalc_totals(ws)
            if total_row:
                update_top_table(ws, total_row)
        else:
            logger.info("Нет обновлённых строк")

        align_numeric_cells(ws)

        # Сортировка блоков подразделений по убыванию просроченной задолженности.
        # Выполняется после пересчёта итогов, чтобы учитывать актуальные суммы
        # (в т.ч. подразделения, где после сверки просрочка стала 0 — уходят вниз).
        try:
            sort_filial_blocks_by_overdue(ws)
        except Exception as e:
            logger.error(f"Ошибка сортировки подразделений: {e}")
            traceback.print_exc()

        current_day_data_from_file = extract_filial_overdue(ws)
        data['currentDayData'] = current_day_data_from_file
        total_debt_data = extract_total_row_debt(ws, total_row)
        logger.info(f"Из итоговой строки ДТ: общая ДЗ={total_debt_data['totalDebt']}, ПДЗ={total_debt_data['totalOverdue']}")

        siuat_file = request.files.get('siUatFile')
        siuat_total_debt = 0
        siuat_total_overdue = 0
        siuat_sheet_created = False
        siuat_sheet1_table = None

        logger.info(f"=== ОТЛАДКА СИ УАТ ===")
        logger.info(f"siuat_file: {siuat_file}")
        logger.info(f"siuat_file.filename: {siuat_file.filename if siuat_file else 'None'}")

        summary_siuat = data.get('summarySIUAT', {})
        if summary_siuat:
            siuat_total_debt = summary_siuat.get('totalDebt', 0) or 0
            siuat_total_overdue = summary_siuat.get('totalOverdue', 0) or 0
            logger.info(f"Получены данные СИ УАТ с фронтенда: общая ДЗ={siuat_total_debt}, ПДЗ={siuat_total_overdue}")

        if siuat_file and siuat_file.filename:
            logger.info(f"=== ДОБАВЛЯЕМ ЛИСТ 'Свод ДЗ СИ УАТ' из файла {siuat_file.filename} ===")
            try:
                sheets_to_delete = [sn for sn in wb.sheetnames if sn.startswith('Свод ДЗ СИ УАТ')]
                for sheet_name in sheets_to_delete:
                    del wb[sheet_name]
                    logger.info(f"Удалён лист: {sheet_name}")

                siuat_file_content = siuat_file.read()
                logger.info(f"Размер файла СИ УАТ: {len(siuat_file_content)} байт")
                siuat_wb = openpyxl.load_workbook(io.BytesIO(siuat_file_content))
                logger.info(f"Листы в файле СИ УАТ: {siuat_wb.sheetnames}")

                siuat_ws = siuat_wb.worksheets[0]
                logger.info(f"Используем лист: '{siuat_ws.title}', размер: {siuat_ws.max_row} строк")

                new_siuat_ws = copy_worksheet_full(siuat_ws, wb)
                new_siuat_ws.title = 'Свод ДЗ СИ УАТ'
                siuat_sheet_created = True
                logger.info(f"Лист скопирован и переименован в 'Свод ДЗ СИ УАТ'")

                # Извлекаем таблицу из листа 'Лист1' (диапазон A5:D8) для письма
                try:
                    siuat_sheet1_table = extract_siuat_sheet1_table(siuat_file_content)
                    logger.info(f"Таблица из 'Лист1' A5:D8: {len(siuat_sheet1_table)} строк")
                except Exception as e:
                    logger.warning(f"Не удалось извлечь таблицу из 'Лист1': {e}")
                    siuat_sheet1_table = None

                # Закрываем промежуточный workbook для освобождения памяти
                siuat_wb.close()
                del siuat_wb
                gc.collect()

                if siuat_total_debt == 0:
                    logger.info("Вызываем extract_siuat_totals_by_max...")
                    file_debt, file_overdue = extract_siuat_totals_by_max(new_siuat_ws)
                    if file_debt > 0:
                        siuat_total_debt = file_debt
                        siuat_total_overdue = file_overdue
                        logger.info(f"Используем данные из файла: общая ДЗ={siuat_total_debt}, ПДЗ={siuat_total_overdue}")
                    else:
                        logger.warning("ВНИМАНИЕ: extract_siuat_totals_by_max вернул 0!")
                else:
                    logger.info("Используем данные с фронтенда, файл не сканируем")
            except Exception as e:
                logger.error(f"!!! Ошибка при добавлении листа СИ УАТ: {e}")
                traceback.print_exc()
        else:
            logger.info("Файл СИ УАТ не загружен (siuat_file=None или filename пустой)")

        logger.info("=== СОЗДАЁМ ЛИСТ 'Сводные таблицы' ===")
        try:
            if 'Сводные таблицы' in wb.sheetnames:
                del wb['Сводные таблицы']
                logger.info("Удалён существующий лист 'Сводные таблицы'")
            summary_ws = wb.create_sheet('Сводные таблицы')
            create_summary_sheet(
                summary_ws, data,
                total_debt=total_debt_data['totalDebt'],
                total_overdue=total_debt_data['totalOverdue'],
                siuat_total_debt=siuat_total_debt,
                siuat_total_overdue=siuat_total_overdue,
            )
            logger.info("Лист 'Сводные таблицы' создан")
        except Exception as e:
            logger.error(f"!!! Ошибка создания сводных таблиц: {e}")
            traceback.print_exc()

        email_data_encoded = None
        if mode == 'email':
            # Скрываем вкладку 'Сводные таблицы' для почтового файла
            try:
                if 'Сводные таблицы' in wb.sheetnames:
                    wb['Сводные таблицы'].sheet_state = 'hidden'
                    logger.info("Вкладка 'Сводные таблицы' скрыта (режим email)")
            except Exception as e:
                logger.warning(f"Не удалось скрыть вкладку 'Сводные таблицы': {e}")

            # Формируем данные для таблиц письма
            try:
                email_data = build_email_data(
                    data,
                    current_day_data_from_file,
                    total_debt_data['totalDebt'],
                    total_debt_data['totalOverdue'],
                    siuat_total_debt,
                    siuat_total_overdue,
                    siuat_sheet1_table,
                )
                email_data_encoded = base64.b64encode(
                    json.dumps(email_data, ensure_ascii=False).encode('utf-8')
                ).decode('ascii')
                logger.info("Данные письма сформированы")
            except Exception as e:
                logger.error(f"Ошибка формирования данных письма: {e}")
                traceback.print_exc()

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        # Закрываем workbook для освобождения памяти
        wb.close()
        del wb
        gc.collect()

        logger.info("=== ФАЙЛ УСПЕШНО ОБРАБОТАН, ОТПРАВЛЯЕМ ===")

        filial_data_json = json.dumps(current_day_data_from_file, ensure_ascii=False)
        filial_data_encoded = base64.b64encode(filial_data_json.encode("utf-8")).decode("ascii")

        response = send_file(
            output,
            as_attachment=True,
            download_name=f"ДЗ_обновленный_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        response.headers["X-Filial-Data"] = filial_data_encoded
        if email_data_encoded:
            response.headers["X-Email-Data"] = email_data_encoded

        return response

    except Exception as e:
        logger.error(f"!!! ОШИБКА в save_excel: {e}")
        traceback.print_exc()
        return {'error': str(e)}, 500

# ============================================================
# API: ДАННЫЕ ПРЕДЫДУЩЕГО ДНЯ
# ============================================================
@app.route('/api/previous-day-data', methods=['POST', 'GET'])
def api_previous_day_data():
    try:
        if request.method == 'POST':
            data = request.get_json()
            if not data:
                return jsonify({'error': 'Нет данных'}), 400

            swipe_date = data.get('date')
            filial_data = data.get('data', {})
            summary_dt = data.get('summaryDT', {})
            summary_siuat = data.get('summarySIUAT', {})

            if not swipe_date:
                return jsonify({'error': 'Не указана дата'}), 400

            logger.info(f"=== СОХРАНЕНИЕ ДАННЫХ ПРЕДЫДУЩЕГО ДНЯ В БД ===")
            logger.info(f"Дата: {swipe_date}, Филиалов: {len(filial_data)}")

            result = db.save_previous_day_data(swipe_date, filial_data, summary_dt, summary_siuat)
            return jsonify(result)

        else:  # GET
            requested_date = request.args.get('date')
            if not requested_date:
                return jsonify({'error': 'Не указана дата (параметр ?date=YYYY-MM-DD)'}), 400

            logger.info(f"=== ЗАГРУЗКА ДАННЫХ ПРЕДЫДУЩЕГО ДНЯ ИЗ БД ===")
            logger.info(f"Запрошенная дата: {requested_date}")

            filial_data = db.get_previous_day_data(requested_date)
            summary_dt = {}
            summary_siuat = {}
            actual_date = requested_date

            if filial_data:
                conn = None
                try:
                    conn = db.get_connection()
                    cur = conn.cursor()
                    cur.execute("""
                        SELECT DISTINCT swipe_date
                        FROM swipe_history
                        WHERE swipe_date = %s
                    """, (requested_date,))
                    result = cur.fetchone()

                    if not result:
                        last_date = db.get_last_available_swipe_date(requested_date)
                        if last_date:
                            actual_date = last_date
                            logger.info(f"🔄 Данные найдены за предыдущую дату: {actual_date}")

                    cur.close()
                except Exception as e:
                    logger.warning(f"⚠️ Не удалось определить фактическую дату: {e}")
                finally:
                    if conn:
                        db.return_connection(conn)

                summary_result = db.get_summary_data(actual_date)
                if summary_result:
                    summary_dt = summary_result.get('summaryDT', {})
                    summary_siuat = summary_result.get('summarySIUAT', {})
                    logger.info(f"✅ Загружены сводные данные за {actual_date}")

                logger.info(f"✅ Загружено {len(filial_data)} филиалов за {actual_date}")
                return jsonify({
                    'success': True,
                    'data': filial_data,
                    'date': actual_date,
                    'requested_date': requested_date,
                    'summaryDT': summary_dt,
                    'summarySIUAT': summary_siuat
                })
            else:
                logger.warning(f"⚠️ Данные не найдены ни за {requested_date}, ни за предыдущие даты")
                return jsonify({
                    'success': True,
                    'data': {},
                    'date': requested_date,
                    'requested_date': requested_date,
                    'message': 'Данные не найдены',
                    'summaryDT': {},
                    'summarySIUAT': {}
                })

    except Exception as e:
        logger.error(f"❌ Ошибка API previous-day-data: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

# ============================================================
# API: РАСПРЕДЕЛЕНИЕ НАЛОГОВ ПО ПОДРАЗДЕЛЕНИЯМ
# ============================================================
@app.route('/api/allocate-cashflow', methods=['POST'])
@limiter.limit("3 per minute")  # Тяжёлая операция
def api_allocate_cashflow():
    try:
        logger.info("=== ЗАПРОС НА РАСПРЕДЕЛЕНИЕ НАЛОГОВ ===")

        file = request.files.get('file')
        if not file:
            return jsonify({'error': 'Файл не загружен'}), 400

        settings_str = request.form.get('settings', '{}')
        settings = json.loads(settings_str)

        logger.info(f"Файл: {file.filename}, Настройки: {settings}")

        file_bytes = file.read()
        logger.info(f"Размер файла: {len(file_bytes)} байт")

        result = allocator.allocate_cashflow(
            file_bytes=file_bytes,
            strategy=settings.get('strategy', 'revenue'),
            rounding=settings.get('rounding', 'kopeks'),
            period_type=settings.get('periodType', 'month'),
            period_value=settings.get('periodValue', ''),
            date_from=settings.get('dateFrom', ''),
            date_to=settings.get('dateTo', ''),
            exclude_articles=settings.get('excludeArticles', [])
        )

        if not result.get('success'):
            return jsonify({'error': result.get('error', 'Ошибка распределения')}), 500

        summary = {
            'success': True,
            'filials_count': result.get('filials_count', 0),
            'total_revenue': result.get('total_revenue', 0),
            'total_taxes': result.get('total_taxes', 0),
            'shares': result.get('shares', {}),
            'period': settings.get('periodType', 'month')
        }

        summary_json = json.dumps(summary, ensure_ascii=False)
        summary_encoded = base64.b64encode(summary_json.encode("utf-8")).decode("ascii")

        logger.info(f"✅ Распределение завершено: {result.get('filials_count')} подразделений")

        response = send_file(
            result['file'],
            as_attachment=True,
            download_name=f"ДДС_налоги_распределены_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        response.headers["X-Allocation-Summary"] = summary_encoded

        return response

    except Exception as e:
        logger.error(f"!!! ОШИБКА РАСПРЕДЕЛЕНИЯ НАЛОГОВ: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ СВЕРКИ ДЗ
# ============================================================
def _cell_to_email_value(value):
    """Преобразует значение ячейки Excel в JSON-сериализуемое значение для письма."""
    if value is None:
        return ''
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return float(value)
    if hasattr(value, 'strftime'):
        return value.strftime('%d.%m.%Y')
    return str(value)


def extract_siuat_sheet1_table(siuat_file_content):
    """Извлекает таблицу из листа 'Лист1' файла СИ УАТ, диапазон A5:D8.

    Файл загружается с data_only=True, чтобы читать вычисленные значения,
    а не формулы.
    """
    wb = openpyxl.load_workbook(io.BytesIO(siuat_file_content), data_only=True)
    try:
        if 'Лист1' not in wb.sheetnames:
            logger.warning("Лист 'Лист1' не найден в файле СИ УАТ")
            return []
        ws = wb['Лист1']
        table = []
        for row_idx in range(5, 9):
            row = []
            for col_idx in range(1, 5):
                row.append(_cell_to_email_value(ws.cell(row=row_idx, column=col_idx).value))
            table.append(row)
        return table
    finally:
        wb.close()


def build_email_data(data, current_day_data, total_debt, total_overdue, siuat_total_debt, siuat_total_overdue, siuat_sheet1_table=None):
    """Формирует структурированные данные таблиц для тела email."""
    current_date = data.get('currentDate', datetime.now().strftime('%Y-%m-%d'))
    previous_date = data.get('previousDate', '')
    previous_day_data = data.get('previousDayData', {})
    summary_dt = data.get('summaryDT', {})

    def _fmt(v):
        try:
            return round(float(v or 0), 2)
        except (TypeError, ValueError):
            return 0.0

    all_filials = sorted(set(list(current_day_data.keys()) + list(previous_day_data.keys())))
    dynamics = []
    total_current = 0.0
    total_previous = 0.0
    for filial in all_filials:
        cv = _fmt(current_day_data.get(filial, 0))
        pv = _fmt(previous_day_data.get(filial, 0))
        total_current += cv
        total_previous += pv
        dynamics.append({
            'name': filial,
            'current': cv,
            'previous': pv,
            'delta': round(pv - cv, 2),
        })

    summary_dt_rows = [
        {'label': 'общая ДЗ', 'value': _fmt(total_debt)},
        {'label': 'из них ПДЗ', 'value': _fmt(total_overdue)},
        {'label': 'Судебная задолженность (Промстрой, АккаунтЭксперт, МКСталь, Кройл)', 'value': _fmt(summary_dt.get('legal', 0))},
    ]

    summary_siuat_rows = [
        {'label': 'Общая ДЗ', 'value': _fmt(siuat_total_debt)},
        {'label': 'Из них ПДЗ', 'value': _fmt(siuat_total_overdue)},
    ]

    return {
        'currentDate': current_date,
        'previousDate': previous_date,
        'dynamics': dynamics,
        'dynamicsTotal': {
            'current': round(total_current, 2),
            'previous': round(total_previous, 2),
            'delta': round(total_previous - total_current, 2),
        },
        'summaryDT': summary_dt_rows,
        'summarySIUAT': summary_siuat_rows,
        'siuatSheet1': siuat_sheet1_table or [],
    }


def create_summary_sheet(ws, data, total_debt=0, total_overdue=0, siuat_total_debt=0, siuat_total_overdue=0):
    logger.info("Создание листа 'Сводные таблицы'...")

    current_date = data.get('currentDate', datetime.now().strftime('%Y-%m-%d'))
    previous_date = data.get('previousDate', '')
    current_day_data = data.get('currentDayData', {})
    previous_day_data = data.get('previousDayData', {})

    title_font = Font(bold=True, size=14)
    header_font_white = Font(bold=True, size=11, color='FFFFFF')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    number_format = '#,##0.00'
    red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    total_font = Font(bold=True, size=11)
    total_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

    row = 1
    ws.cell(row=row, column=1, value='Динамика по подразделениям').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    row += 2

    def format_date_ddmmyyyy(date_str):
        if not date_str or date_str in ('предыдущий рабочий день',):
            return date_str
        try:
            dt = datetime.fromisoformat(date_str)
            return dt.strftime('%d.%m.%Y')
        except (ValueError, TypeError):
            return date_str

    current_date_formatted = format_date_ddmmyyyy(current_date)
    previous_date_formatted = format_date_ddmmyyyy(previous_date)

    headers = ['Подразделение', current_date_formatted, previous_date_formatted, 'Динамика']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    row += 1

    all_filials = sorted(set(list(current_day_data.keys()) + list(previous_day_data.keys())))
    total_current = total_previous = total_delta = 0

    for filial in all_filials:
        cv = current_day_data.get(filial, 0)
        pv = previous_day_data.get(filial, 0)
        d = pv - cv
        total_current += cv
        total_previous += pv
        total_delta += d
        ws.cell(row=row, column=1, value=filial).border = thin_border
        cc = ws.cell(row=row, column=2, value=cv)
        cc.number_format = number_format
        cc.border = thin_border
        cc.alignment = Alignment(horizontal='right')
        cp = ws.cell(row=row, column=3, value=pv)
        cp.number_format = number_format
        cp.border = thin_border
        cp.alignment = Alignment(horizontal='right')
        cd = ws.cell(row=row, column=4, value=d)
        cd.number_format = number_format
        cd.border = thin_border
        cd.alignment = Alignment(horizontal='right')
        if d > 0:
            cd.fill = green_fill
        elif d < 0:
            cd.fill = red_fill
        row += 1

    ws.cell(row=row, column=1, value='Общий итог').font = total_font
    ws.cell(row=row, column=1).fill = total_fill
    ws.cell(row=row, column=1).border = thin_border
    for col, val in enumerate([total_current, total_previous, total_delta], 2):
        c = ws.cell(row=row, column=col, value=val)
        c.number_format = number_format
        c.font = total_font
        c.fill = total_fill
        c.border = thin_border
        c.alignment = Alignment(horizontal='right')
    row += 3

    ws.cell(row=row, column=1, value='Свод задолженности ДТ').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    row += 2
    summary_dt = data.get('summaryDT', {})
    rows_dt = [
        ('общая ДЗ', total_debt),
        ('из них ПДЗ', total_overdue),
        ('Судебная задолженность (Промстрой, АккаунтЭксперт, МКСталь, Кройл)', summary_dt.get('legal', 0)),
    ]
    for label, value in rows_dt:
        cl = ws.cell(row=row, column=1, value=label)
        cl.border = thin_border
        if 'ПДЗ' in label:
            cl.font = Font(bold=True, color='FF0000')
        elif 'Судебная' in label:
            cl.font = Font(bold=True, color='0000FF')
        else:
            cl.font = Font(bold=True)
        cv = ws.cell(row=row, column=2, value=value)
        cv.number_format = number_format
        cv.border = thin_border
        cv.alignment = Alignment(horizontal='right')
        if 'ПДЗ' in label:
            cv.font = Font(bold=True, color='FF0000')
        elif 'Судебная' in label:
            cv.font = Font(bold=True, color='0000FF')
        else:
            cv.font = Font(bold=True)
        row += 1
    row += 3

    ws.cell(row=row, column=1, value='Свод задолженности СИ УАТ').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    row += 2
    for label, value in [('Общая ДЗ', siuat_total_debt), ('Из них ПДЗ', siuat_total_overdue)]:
        cl = ws.cell(row=row, column=1, value=label)
        cl.border = thin_border
        cl.font = Font(bold=True)
        cv = ws.cell(row=row, column=2, value=value)
        cv.number_format = number_format
        cv.border = thin_border
        cv.alignment = Alignment(horizontal='right')
        if 'ПДЗ' in label:
            cv.font = Font(bold=True, color='FF0000')
        else:
            cv.font = Font(bold=True)
        row += 1
    logger.info("Лист 'Сводные таблицы' создан")


def extract_filial_overdue(ws):
    data = {}
    for r in range(14, ws.max_row + 1):
        v = get_cell_value(ws, r, 1)
        if v and str(v).strip().startswith('ДТ '):
            cell = ws.cell(row=r, column=1)
            indent = int(cell.alignment.indent) if cell.alignment and cell.alignment.indent is not None else 0
            if indent != 0:
                continue
            ov = get_cell_value(ws, r, COLUMNS['OVERDUE'])
            data[str(v).strip()] = ov if isinstance(ov, (int, float)) else 0
    for k in data:
        data[k] = round(data[k], 2)
    logger.info(f"=== ДАННЫЕ ФИЛИАЛОВ ===")
    for f, a in sorted(data.items()):
        logger.info(f"  {f}: {a:,.2f}")
    return data


def find_siuat_columns(ws):
    total_col = None
    overdue_col = None
    total_candidates = []
    overdue_candidates = []

    for r in range(1, min(21, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=r, column=c).value
            if not cell_val:
                continue
            cell_str = str(cell_val).lower().strip()
            cell_str = ' '.join(cell_str.split())

            if 'всего' in cell_str:
                total_candidates.append((c, 0, cell_str))
            elif 'общая дз' in cell_str:
                total_candidates.append((c, 1, cell_str))
            elif 'общая задолженность' in cell_str:
                total_candidates.append((c, 2, cell_str))
            elif 'общая' in cell_str or 'total' in cell_str:
                total_candidates.append((c, 3, cell_str))

            if 'просроченно' in cell_str:
                overdue_candidates.append((c, 0, cell_str))
            elif 'просрочка' in cell_str:
                overdue_candidates.append((c, 1, cell_str))
            elif 'пдз' in cell_str:
                overdue_candidates.append((c, 2, cell_str))
            elif 'просроченная дз' in cell_str or 'просроченная задолженность' in cell_str:
                overdue_candidates.append((c, 3, cell_str))
            elif 'overdue' in cell_str:
                overdue_candidates.append((c, 3, cell_str))

    if total_candidates:
        total_candidates.sort(key=lambda x: (x[1], x[0]))
        best = total_candidates[0]
        total_col = best[0]
        logger.info(f"Найдена колонка 'всего': колонка {total_col} (приоритет {best[1]}, '{best[2]}')")

    if overdue_candidates:
        overdue_candidates.sort(key=lambda x: (x[1], x[0]))
        best = overdue_candidates[0]
        overdue_col = best[0]
        logger.info(f"Найдена колонка 'просроченно': колонка {overdue_col} (приоритет {best[1]}, '{best[2]}')")

    if total_col is None:
        total_col = 12
        logger.info(f"Fallback: колонка 'всего' = {total_col}")
    if overdue_col is None:
        overdue_col = 15
        logger.info(f"Fallback: колонка 'просроченно' = {overdue_col}")

    return total_col, overdue_col


def _parse_cell_number(value):
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        cleaned = value.replace(' ', '').replace(',', '.')
        try:
            return float(cleaned)
        except ValueError:
            return 0
    return 0


def extract_siuat_totals_by_max(ws):
    total_col = 12
    overdue_col = 15
    logger.info(f"=== extract_siuat_totals_by_max START ===")
    logger.info(f"Лист: '{ws.title}', строк: {ws.max_row}, колонок: {ws.max_column}")

    total_debt = 0
    total_overdue = 0
    dt_count = 0

    for r in range(1, ws.max_row + 1):
        cell_val = get_cell_value(ws, r, 1)
        if not cell_val:
            continue
        str_val = str(cell_val).strip()
        if str_val.startswith('ДТ '):
            v_total = _parse_cell_number(get_cell_value(ws, r, total_col))
            v_overdue = _parse_cell_number(get_cell_value(ws, r, overdue_col))
            total_debt += v_total
            total_overdue += v_overdue
            dt_count += 1

    total_debt = round(total_debt, 2)
    total_overdue = round(total_overdue, 2)
    logger.info(f"Найдено подразделений ДТ: {dt_count}")
    logger.info(f"Результат: общая ДЗ={total_debt}, ПДЗ={total_overdue}")
    return total_debt, total_overdue


def extract_siuat_totals(ws):
    total_col, overdue_col = find_siuat_columns(ws)
    total_debt = 0
    total_overdue = 0

    for r in range(1, ws.max_row + 1):
        cell_val = get_cell_value(ws, r, 1)
        if cell_val:
            str_val = str(cell_val).strip().lower()
            if 'итог' in str_val or 'total' in str_val or 'всего' in str_val:
                v_total = get_cell_value(ws, r, total_col)
                v_overdue = get_cell_value(ws, r, overdue_col)
                total_debt = round(v_total, 2) if isinstance(v_total, (int, float)) else 0
                total_overdue = round(v_overdue, 2) if isinstance(v_overdue, (int, float)) else 0
                logger.info(f"СИ УАТ из строки '{cell_val}' (row {r}): общая ДЗ={total_debt}, ПДЗ={total_overdue}")
                break

    if total_debt == 0:
        logger.info("Строка 'ИТОГО' не найдена, используем fallback...")
        for r in range(ws.max_row, 0, -1):
            val = get_cell_value(ws, r, total_col)
            if isinstance(val, (int, float)) and val > 0:
                total_debt = round(val, 2)
                v_overdue = get_cell_value(ws, r, overdue_col)
                total_overdue = round(v_overdue, 2) if isinstance(v_overdue, (int, float)) else 0
                logger.info(f"СИ УАТ fallback строка {r}: общая ДЗ={total_debt}, ПДЗ={total_overdue}")
                break

    return total_debt, total_overdue


def extract_total_row_debt(ws, total_row):
    result = {'totalDebt': 0, 'totalOverdue': 0}
    if not total_row:
        return result
    td = get_cell_value(ws, total_row, COLUMNS['DEBT_AMOUNT'])
    to = get_cell_value(ws, total_row, COLUMNS['OVERDUE'])
    result['totalDebt'] = round(td, 2) if isinstance(td, (int, float)) else 0
    result['totalOverdue'] = round(to, 2) if isinstance(to, (int, float)) else 0
    return result


@app.route('/save-suppliers', methods=['POST'])
@limiter.limit("5 per minute")
def save_suppliers():
    try:
        file = request.files['file']
        data = json.loads(request.form['data'])

        logger.info(f"=== ПОЛУЧЕН ЗАПРОС НА ОБРАБОТКУ ОПЛАТ ПОСТАВЩИКАМ ===")
        logger.info(f"Файл: {file.filename}")
        logger.info(f"Сводных таблиц: {len(data.get('pivotTables', []))}")

        wb = openpyxl.load_workbook(io.BytesIO(file.read()))

        for pivot_table in data.get('pivotTables', []):
            sheet_name = pivot_table['sheetName']
            headers = pivot_table['headers']
            rows_data = pivot_table['data']

            logger.info(f"Обработка сводной таблицы: {sheet_name}")
            logger.info(f"  Подразделений: {len(headers)}")
            logger.info(f"  Контрагентов: {len(rows_data)}")

            if sheet_name not in wb.sheetnames:
                logger.warning(f"  Предупреждение: лист '{sheet_name}' не найден, создаём новый")
                ws = wb.create_sheet(sheet_name)
                last_row = 0
            else:
                ws = wb[sheet_name]
                last_row = ws.max_row
                logger.info(f"  Реестр заканчивается на строке: {last_row}")

            pivot_start_row = last_row + 4
            create_pivot_sheet_at_row(ws, headers, rows_data, 'Сводная таблица', pivot_start_row)
            logger.info(f"  Сводная таблица добавлена со строки: {pivot_start_row}")

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        # Очистка памяти
        wb.close()
        del wb
        gc.collect()

        logger.info("=== ФАЙЛ ОПЛАТ УСПЕШНО ОБРАБОТАН, ОТПРАВЛЯЕМ ===")

        return send_file(
            output,
            as_attachment=True,
            download_name=f'Оплаты_поставщикам_{datetime.now().strftime("%Y-%m-%d")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        logger.error(f"!!! ОШИБКА ПРИ ОБРАБОТКЕ ОПЛАТ: {e}")
        traceback.print_exc()
        return {'error': str(e)}, 500


def create_pivot_sheet_at_row(ws, headers, rows_data, title, start_row):
    logger.info(f"Создание сводной таблицы '{title}', начиная со строки {start_row}...")

    title_font = Font(bold=True, size=14)
    header_font = Font(bold=True, size=11, color='FFFFFF')
    header_fill = PatternFill(start_color='1F3864', end_color='1F3864', fill_type='solid')
    explanation_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    total_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    number_format = '#,##0.00'
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    row = start_row
    ws.cell(row=row, column=1, value='Сводная таблица оплат по подразделениям').font = title_font
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers) + 3)
    row += 2

    header_cells = ['Контрагент'] + headers + ['Итого', 'Пояснение']
    for col, header in enumerate(header_cells, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    row += 1

    total_all = 0
    for item in rows_data:
        ws.cell(row=row, column=1, value=item['contractor']).border = thin_border
        total_sum = 0
        for col_idx, h in enumerate(headers, 2):
            value = item.get(h, 0)
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.number_format = number_format
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right')
            total_sum += value
        total_all += total_sum
        cell_total = ws.cell(row=row, column=len(headers) + 2, value=total_sum)
        cell_total.number_format = number_format
        cell_total.border = thin_border
        cell_total.alignment = Alignment(horizontal='right')
        cell_explanation = ws.cell(row=row, column=len(headers) + 3, value=item.get('explanation', ''))
        cell_explanation.border = thin_border
        if item.get('explanation'):
            for col_idx in range(1, len(headers) + 4):
                ws.cell(row=row, column=col_idx).fill = explanation_fill
        row += 1

    ws.cell(row=row, column=1, value='ИТОГО').font = Font(bold=True, size=11)
    ws.cell(row=row, column=1).fill = total_fill
    ws.cell(row=row, column=1).border = thin_border

    for col_idx, h in enumerate(headers, 2):
        subtotal = sum(item.get(h, 0) for item in rows_data)
        cell = ws.cell(row=row, column=col_idx, value=subtotal)
        cell.number_format = number_format
        cell.font = Font(bold=True, size=11)
        cell.fill = total_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='right')

    cell_grand_total = ws.cell(row=row, column=len(headers) + 2, value=total_all)
    cell_grand_total.number_format = number_format
    cell_grand_total.font = Font(bold=True, size=11)
    cell_grand_total.fill = total_fill
    cell_grand_total.border = thin_border
    cell_grand_total.alignment = Alignment(horizontal='right')

    ws.cell(row=row, column=len(headers) + 3).fill = total_fill
    ws.cell(row=row, column=len(headers) + 3).border = thin_border

    logger.info(f"Сводная таблица '{title}' создана, строк: {row - start_row + 1}")


# ============================================================
# API для отчётов (Dashboard)
# ============================================================

@app.route('/api/save-swipe-data', methods=['POST'])
def api_save_swipe_data():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Нет данных'}), 400

        swipe_date = data.get('swipeDate')
        filial_data = data.get('filialData', {})
        counterparty_data = data.get('counterpartyData', {})
        total_debt = data.get('totalDebt', 0)
        total_overdue = data.get('totalOverdue', 0)
        summary_dt = data.get('summaryDT', {})
        summary_siuat = data.get('summarySIUAT', {})

        if not swipe_date:
            return jsonify({'error': 'Не указана дата сверки'}), 400

        logger.info(f"=== СОХРАНЕНИЕ ДАННЫХ СВЕРКИ В БД ===")
        logger.info(f"Дата: {swipe_date}, Филиалов: {len(filial_data)}, Контрагентов: {len(counterparty_data)}")

        result = db.save_swipe_data(
            swipe_date, filial_data, counterparty_data,
            total_debt, total_overdue, summary_dt, summary_siuat
        )
        return jsonify(result)

    except Exception as e:
        logger.error(f"❌ Ошибка API save-swipe-data: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/swipe-dates', methods=['GET'])
def api_swipe_dates():
    try:
        from_date = request.args.get('from')
        to_date = request.args.get('to')
        dates = db.get_swipe_dates(from_date, to_date)
        return jsonify({'success': True, 'data': dates})
    except Exception as e:
        logger.error(f"❌ Ошибка API swipe-dates: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/swipe-raw', methods=['GET'])
def api_swipe_raw():
    try:
        from_date = request.args.get('from')
        to_date = request.args.get('to')
        filial = request.args.get('filial')
        if not from_date or not to_date:
            return jsonify({'error': 'Укажите период (from и to)'}), 400
        result = db.get_swipe_raw_data(from_date, to_date, filial)
        return jsonify({'success': True, 'data': result})
    except Exception as e:
        logger.error(f"❌ Ошибка API swipe-raw: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/filial-trend', methods=['GET'])
def api_filial_trend():
    try:
        from_date = request.args.get('from')
        to_date = request.args.get('to')
        filial = request.args.get('filial')
        if not from_date or not to_date:
            return jsonify({'error': 'Укажите период (from и to)'}), 400
        result = db.get_filial_trend(from_date, to_date, filial)
        return jsonify({'success': True, 'data': result})
    except Exception as e:
        logger.error(f"❌ Ошибка API filial-trend: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/counterparty-trend', methods=['GET'])
def api_counterparty_trend():
    try:
        from_date = request.args.get('from')
        to_date = request.args.get('to')
        filial = request.args.get('filial')
        counterparty = request.args.get('counterparty')
        if not from_date or not to_date:
            return jsonify({'error': 'Укажите период (from и to)'}), 400
        result = db.get_counterparty_trend(from_date, to_date, filial, counterparty)
        return jsonify({'success': True, 'data': result})
    except Exception as e:
        logger.error(f"❌ Ошибка API counterparty-trend: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/filial-list', methods=['GET'])
def api_filial_list():
    try:
        from_date = request.args.get('from')
        to_date = request.args.get('to')
        result = db.get_filial_list(from_date, to_date)
        return jsonify({'success': True, 'data': result})
    except Exception as e:
        logger.error(f"❌ Ошибка API filial-list: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/counterparty-list', methods=['GET'])
def api_counterparty_list():
    try:
        filial = request.args.get('filial')
        result = db.get_counterparty_list(filial)
        return jsonify({'success': True, 'data': result})
    except Exception as e:
        logger.error(f"❌ Ошибка API counterparty-list: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/summary', methods=['GET'])
def api_summary():
    try:
        from_date = request.args.get('from')
        to_date = request.args.get('to')
        if not from_date or not to_date:
            return jsonify({'error': 'Укажите период (from и to)'}), 400
        result = db.get_summary(from_date, to_date)
        if result:
            return jsonify({'success': True, 'data': result})
        else:
            return jsonify({'success': True, 'data': None, 'message': 'Нет данных за указанный период'})
    except Exception as e:
        logger.error(f"❌ Ошибка API summary: {e}")
        return jsonify({'error': str(e)}), 500


# ============================================================
# API: НАСТРОЙКИ ПАРСЕРА — МАППИНГ СЧЕТОВ
# ============================================================

@app.route('/api/account-mapping', methods=['GET'])
def api_get_account_mapping():
    try:
        all_records = request.args.get('all', 'false').lower() == 'true'
        result = db.get_account_mapping(active_only=not all_records)
        return jsonify({'success': True, 'data': result})
    except Exception as e:
        logger.error(f"❌ Ошибка API account-mapping GET: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/account-mapping', methods=['POST'])
def api_add_account_mapping():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Нет данных'}), 400
        account_number = data.get('account_number', '').strip()
        company_name = data.get('company_name', '').strip()
        bank_name = data.get('bank_name', '').strip()
        if not account_number or not company_name:
            return jsonify({'error': 'Номер счёта и название компании обязательны'}), 400
        result = db.add_account_mapping(account_number, company_name, bank_name)
        if result.get('success'):
            return jsonify(result)
        else:
            return jsonify({'error': result.get('error', 'Ошибка сохранения')}), 500
    except Exception as e:
        logger.error(f"❌ Ошибка API account-mapping POST: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/account-mapping/<int:mapping_id>', methods=['PUT'])
def api_update_account_mapping(mapping_id):
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Нет данных'}), 400
        result = db.update_account_mapping(
            mapping_id,
            account_number=data.get('account_number'),
            company_name=data.get('company_name'),
            bank_name=data.get('bank_name'),
            is_active=data.get('is_active')
        )
        if result.get('success'):
            return jsonify(result)
        else:
            return jsonify({'error': result.get('error', 'Ошибка обновления')}), 500
    except Exception as e:
        logger.error(f"❌ Ошибка API account-mapping PUT: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/account-mapping/<int:mapping_id>', methods=['DELETE'])
def api_delete_account_mapping(mapping_id):
    try:
        result = db.delete_account_mapping(mapping_id)
        if result.get('success'):
            return jsonify(result)
        else:
            return jsonify({'error': result.get('error', 'Ошибка удаления')}), 500
    except Exception as e:
        logger.error(f"❌ Ошибка API account-mapping DELETE: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/companies', methods=['GET'])
def api_companies_list():
    try:
        result = db.get_companies_list()
        return jsonify({'success': True, 'data': result})
    except Exception as e:
        logger.error(f"❌ Ошибка API companies: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/banks', methods=['GET'])
def api_banks_list():
    try:
        result = db.get_banks_list()
        return jsonify({'success': True, 'data': result})
    except Exception as e:
        logger.error(f"❌ Ошибка API banks: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/account-mapping/sync', methods=['POST'])
def api_sync_account_mapping():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Нет данных'}), 400
        mapping_list = data.get('mapping', [])
        if not mapping_list:
            return jsonify({'error': 'Пустой маппинг'}), 400
        result = db.sync_account_mapping_from_parser(mapping_list)
        return jsonify(result)
    except Exception as e:
        logger.error(f"❌ Ошибка API account-mapping/sync: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/delete-swipe', methods=['POST'])
def api_delete_swipe():
    try:
        data = request.get_json()
        swipe_date = data.get('date')
        if not swipe_date:
            return jsonify({'error': 'Не указана дата'}), 400
        result = db.delete_swipe_data(swipe_date)
        return jsonify(result)
    except Exception as e:
        logger.error(f"❌ Ошибка API delete-swipe: {e}")
        return jsonify({'error': str(e)}), 500


# ============================================================
# API: БИБЛИОТЕКА КОНТРАГЕНТОВ (CRUD)
# ============================================================

@app.route('/api/contractors', methods=['GET'])
def api_get_contractors():
    try:
        contractors = db.get_all_contractors()
        return jsonify({'success': True, 'data': contractors})
    except Exception as e:
        logger.error(f"❌ Ошибка API contractors GET: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/contractors', methods=['POST'])
def api_add_contractor():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Нет данных'}), 400
        name = data.get('name', '').strip()
        if not name:
            return jsonify({'error': 'Имя контрагента обязательно'}), 400
        organization = data.get('organization', '')
        explanation = data.get('explanation', '')
        result = db.insert_or_update_contractor(name, organization, explanation)
        return jsonify(result)
    except Exception as e:
        logger.error(f"❌ Ошибка API contractors POST: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/contractors/<int:contractor_id>', methods=['PUT'])
def api_update_contractor(contractor_id):
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Нет данных'}), 400
        result = db.update_contractor(
            contractor_id,
            name=data.get('name'),
            organization=data.get('organization'),
            explanation=data.get('explanation')
        )
        return jsonify(result)
    except Exception as e:
        logger.error(f"❌ Ошибка API contractors PUT: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/contractors/<int:contractor_id>', methods=['DELETE'])
def api_delete_contractor(contractor_id):
    try:
        result = db.delete_contractor(contractor_id)
        return jsonify(result)
    except Exception as e:
        logger.error(f"❌ Ошибка API contractors DELETE: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/contractors/find', methods=['GET'])
def api_find_contractor():
    try:
        name = request.args.get('name', '').strip()
        if not name:
            return jsonify({'error': 'Не указано имя для поиска'}), 400
        result = db.find_contractor_by_name(name)
        if result:
            return jsonify({'success': True, 'data': result})
        else:
            return jsonify({'success': True, 'data': None})
    except Exception as e:
        logger.error(f"❌ Ошибка API contractors/find: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/contractors/import', methods=['POST'])
def api_import_contractors():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Нет данных'}), 400
        contractors_list = data.get('contractors', [])
        if not contractors_list:
            return jsonify({'error': 'Пустой список контрагентов'}), 400
        result = db.import_contractors_batch(contractors_list)
        return jsonify(result)
    except Exception as e:
        logger.error(f"❌ Ошибка API contractors/import: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/contractors/clear', methods=['POST'])
def api_clear_contractors():
    try:
        result = db.delete_all_contractors()
        return jsonify(result)
    except Exception as e:
        logger.error(f"❌ Ошибка API contractors/clear: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/contractors/stats', methods=['GET'])
def api_contractors_stats():
    try:
        stats = db.get_contractors_stats()
        return jsonify({'success': True, 'data': stats})
    except Exception as e:
        logger.error(f"❌ Ошибка API contractors/stats: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ============================================================
# API: ГИБКИЕ НАСТРОЙКИ (правила исключений, категоризации, синонимы)
# ============================================================

@app.route('/api/config', methods=['GET'])
def api_get_config():
    try:
        config = db.get_all_config()
        return jsonify({'success': True, 'data': config})
    except Exception as e:
        logger.error(f"❌ Ошибка загрузки конфигурации: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/exclusion-rules', methods=['GET'])
def api_get_exclusion_rules():
    try:
        rules = db.get_exclusion_rules()
        return jsonify({'success': True, 'data': rules})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/exclusion-rules', methods=['POST'])
def api_add_exclusion_rule():
    try:
        data = request.get_json()
        result = db.save_exclusion_rule(
            data.get('rule_type', 'purpose'),
            data.get('pattern', ''),
            data.get('is_regex', False)
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/exclusion-rules/<int:rule_id>', methods=['DELETE'])
def api_delete_exclusion_rule(rule_id):
    try:
        result = db.delete_exclusion_rule(rule_id)
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/exclusion-rules/<int:rule_id>', methods=['PUT'])
def api_update_exclusion_rule(rule_id):
    try:
        data = request.get_json()
        result = db.update_exclusion_rule(
            rule_id,
            rule_type=data.get('rule_type'),
            pattern=data.get('pattern'),
            is_regex=data.get('is_regex')
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/categorization-rules', methods=['GET'])
def api_get_categorization_rules():
    try:
        rules = db.get_categorization_rules()
        return jsonify({'success': True, 'data': rules})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/categorization-rules', methods=['POST'])
def api_add_categorization_rule():
    try:
        data = request.get_json()
        result = db.save_categorization_rule(
            data.get('field', 'purpose'),
            data.get('pattern', ''),
            data.get('display_name', '')
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/categorization-rules/<int:rule_id>', methods=['DELETE'])
def api_delete_categorization_rule(rule_id):
    try:
        result = db.delete_categorization_rule(rule_id)
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/categorization-rules/<int:rule_id>', methods=['PUT'])
def api_update_categorization_rule(rule_id):
    try:
        data = request.get_json()
        result = db.update_categorization_rule(
            rule_id,
            field=data.get('field'),
            pattern=data.get('pattern'),
            display_name=data.get('display_name')
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/company-aliases', methods=['GET'])
def api_get_company_aliases():
    try:
        aliases = db.get_company_aliases()
        return jsonify({'success': True, 'data': aliases})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/company-aliases', methods=['POST'])
def api_add_company_alias():
    try:
        data = request.get_json()
        result = db.save_company_alias(
            data.get('pattern', ''),
            data.get('canonical_name', ''),
            data.get('match_type', 'contains')
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/company-aliases/<int:alias_id>', methods=['DELETE'])
def api_delete_company_alias(alias_id):
    try:
        result = db.delete_company_alias(alias_id)
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/company-aliases/<int:alias_id>', methods=['PUT'])
def api_update_company_alias(alias_id):
    try:
        data = request.get_json()
        result = db.update_company_alias(
            alias_id,
            pattern=data.get('pattern'),
            canonical=data.get('canonical_name'),
            match_type=data.get('match_type')
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/account-mapping/<int:mapping_id>/toggle', methods=['PUT'])
def api_toggle_account(mapping_id):
    try:
        accounts = db.get_account_mapping(active_only=False)
        target = next((a for a in accounts if a['id'] == mapping_id), None)
        if not target:
            return jsonify({'success': False, 'error': 'Счёт не найден'}), 404
        new_active = not target['is_active']
        result = db.update_account_mapping(mapping_id, is_active=new_active)
        if result['success']:
            result['is_active'] = new_active
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ============================================================
# API: СРОЧНЫЕ ДЕПОЗИТЫ (term_deposits)
# ============================================================

@app.route('/api/term-deposits', methods=['GET'])
def api_get_term_deposits():
    """Получить все срочные депозиты."""
    try:
        deposits = db.get_term_deposits()
        # Группируем по account_number для фронтенда
        result = {}
        for d in deposits:
            acc = d['account_number']
            if acc not in result:
                result[acc] = []
            result[acc].append({
                'amount': float(d['amount']),
                'rate': float(d['rate']),
                'startDate': d['start_date'].isoformat() if hasattr(d['start_date'], 'isoformat') else str(d['start_date']),
                'endDate': d['end_date'].isoformat() if hasattr(d['end_date'], 'isoformat') else str(d['end_date'])
            })
        return jsonify({'success': True, 'data': result})
    except Exception as e:
        logger.error(f"Ошибка API term-deposits GET: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/term-deposits/sync', methods=['POST'])
def api_sync_term_deposits():
    """Синхронизировать срочные депозиты (полная замена)."""
    try:
        data = request.get_json()
        deposits = data.get('deposits', {})
        result = db.sync_term_deposits(deposits)
        return jsonify({'success': True, 'synced': result})
    except Exception as e:
        logger.error(f"Ошибка API term-deposits/sync: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

# ============================================================
# API КОНТРОЛЯ ЛИМИТОВ ПЛАТЕЖЕЙ ПОСТАВЩИКАМ
# ============================================================

def payment_period_for_date(dt):
    """Период оплаты по дате загрузки: P1 (1-7), P2 (8-15), P3 (16-22), P4 (23-31)."""
    day = dt.day
    if day <= 7:
        return 1
    if day <= 15:
        return 2
    if day <= 22:
        return 3
    return 4


def _parse_plp_date(value):
    """Конвертирует значение даты (ISO-строка или None) в str YYYY-MM-DD или None."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    # Оставляем только дату из ISO 'YYYY-MM-DD...'
    s = s[:10]
    try:
        datetime.strptime(s, '%Y-%m-%d')
        return s
    except ValueError:
        return None


@app.route('/api/payment-limits/config', methods=['GET'])
def api_payment_limits_config_get():
    """Возвращает настройки лимитов платежей (категории, значения, подстатьи, глобальный лимит)."""
    try:
        config = db.get_payment_limits_config()
        return jsonify({'success': True, 'config': config})
    except Exception as e:
        logger.error(f"Ошибка API payment-limits/config GET: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/payment-limits/config', methods=['POST'])
@app.route('/api/payment-limits/config/save', methods=['POST'])
def api_payment_limits_config_save():
    """
    Сохраняет настройки лимитов (названия, значения по периодам, подстатьи, глобальный лимит).
    При переименовании категорий/подстатей обновляет уже подтверждённые платежи в БД.
    """
    try:
        data = request.get_json(force=True, silent=True) or {}
        config = data.get('config')
        result = db.save_payment_limits_config(config)
        if not result.get('success'):
            return jsonify(result), 400
        return jsonify(result)
    except Exception as e:
        logger.error(f"Ошибка API payment-limits/config POST: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/payment-limits/confirmed', methods=['GET'])
def api_payment_limits_confirmed():
    """
    Подтверждённые платежи за месяц (для накопительного итога в превью).

    Query param: month=YYYY-MM (по умолчанию — текущий месяц).
    При первом обращении в новом месяце автоматически удаляет данные
    предыдущих месяцев (начало нового месяца — учёт заново).
    """
    try:
        now = datetime.now()
        month = request.args.get('month')
        if not month:
            month = now.strftime('%Y-%m')

        month_first = month[:7] + '-01'
        db.cleanup_payment_limits_before(month_first)

        # stage: 'final' — только финал, 'draft' — черновик,
        # 'effective' — финал, если он сохранён, иначе черновик (для накопительных итогов),
        # 'all' — все стадии
        stage = (request.args.get('stage') or db.PLP_STAGE_FINAL).strip().lower()
        if stage == 'effective':
            # По каждой сохранённой загрузке (партии) берём финал, иначе черновик;
            # строки всех загрузок периода складываются.
            rows = db.get_payment_limits_effective(month_first)
        else:
            if stage not in (db.PLP_STAGE_DRAFT, db.PLP_STAGE_FINAL):
                stage = None
            rows = db.get_payment_limits_payments(month_first, stage)
        # Денормализуем значения дат/сумм в JSON-совместимый вид
        result = []
        for r in rows:
            item = dict(r)
            for key in ('request_date', 'payment_date'):
                val = item.get(key)
                if val is not None:
                    item[key] = val.isoformat() if hasattr(val, 'isoformat') else str(val)
            for key in ('limit_month',):
                val = item.get(key)
                if val is not None:
                    item[key] = val.isoformat() if hasattr(val, 'isoformat') else str(val)
            for key in ('confirmed_at',):
                val = item.get(key)
                if val is not None:
                    item[key] = val.isoformat() if hasattr(val, 'isoformat') else str(val)
            item['amount'] = float(item.get('amount') or 0)
            result.append(item)

        return jsonify({
            'success': True,
            'month': month_first,
            'period': payment_period_for_date(now),
            'rows': result
        })
    except Exception as e:
        logger.error(f"Ошибка API payment-limits/confirmed: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/payment-limits/confirm', methods=['POST'])
def api_payment_limits_confirm():
    """
    Подтверждение и сохранение загрузки платежей.

    Тело: JSON {'rows': [ {...}, ... ], 'period': 1..4 (необязательно)}.
    Период берётся из запроса (выбор пользователя в интерфейсе), если он валиден;
    иначе определяется по ДАТЕ ПОДТВЕРЖДЕНИЯ на сервере.
    Месяц учёта — по дате подтверждения. При смене месяца данные предыдущих
    месяцев автоматически очищаются.
    """
    try:
        data = request.get_json(force=True, silent=True) or {}
        rows = data.get('rows') or []
        if not rows:
            return jsonify({'success': False, 'error': 'Нет данных для сохранения'}), 400

        now = datetime.now()
        month_first = now.strftime('%Y-%m') + '-01'

        # Период: выбранный пользователем (1..4) или авто по дате подтверждения
        period = None
        try:
            if data.get('period') is not None:
                period = int(data.get('period'))
        except (TypeError, ValueError):
            period = None
        if period is None or period < 1 or period > 4:
            period = payment_period_for_date(now)

        # Автоочистка: начало нового месяца — начинаем учёт заново
        db.cleanup_payment_limits_before(month_first)

        payments = []
        for row in rows:
            payments.append({
                'limit_month': month_first,
                'period': period,
                'article': (row.get('article') or '').strip(),
                'detail_limit': (row.get('detail_limit') or '').strip(),
                'amount': float(row.get('amount') or 0),
                'recipient': (row.get('recipient') or '').strip(),
                'request_no': (row.get('request_no') or '').strip(),
                'request_date': _parse_plp_date(row.get('request_date')),
                'status': (row.get('status') or '').strip(),
                'organization': (row.get('organization') or '').strip(),
                'subdivision': (row.get('subdivision') or '').strip(),
                'priority': (row.get('priority') or '').strip(),
                'purpose': (row.get('purpose') or '').strip(),
                'payment_date': _parse_plp_date(row.get('payment_date')),
                'applicant': (row.get('applicant') or '').strip(),
                'payment_type': (row.get('payment_type') or '').strip()
            })

        result = db.insert_payment_limits_payments(payments)
        if not result.get('success'):
            return jsonify(result), 500

        total_amount = sum(p['amount'] for p in payments)
        logger.info(
            f"✅ Подтверждено платежей по лимитам: {result['count']} шт., "
            f"сумма {total_amount:.2f} ₽, период P{period}, месяц {month_first}"
        )
        return jsonify({
            'success': True,
            'month': month_first,
            'period': period,
            'count': result['count'],
            'total_amount': total_amount,
            'message': f'Подтверждено {result["count"]} платежей '
                       f'(период P{period}, {month_first[:7]})'
        })
    except Exception as e:
        logger.error(f"Ошибка API payment-limits/confirm: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ============================================================
# ЭКСПОРТ РЕЕСТРА «ЛИМИТЫ ПЛАТЕЖЕЙ» — исходный файл без снятых заявок
# ============================================================
# Файл НЕ пересобирается заново: копируется исходный .xlsx (тема, стили, шрифты,
# заливки, форматы, ширины/высоты, printerSettings, sharedStrings) и из XML листа
# удаляются строки снятых заявок — с перенумерацией строк/адресов ячеек и поправкой
# dimension, автофильтра и диапазонов условного форматирования (как это делает Excel).
# Поэтому порядок и наполнение столбцов, «Дата заявки» (время по Москве) и оформление
# остаются точно такими же, как в загруженном файле.
_PL_ROW_RE = re.compile(r'<row[^>]*\br="(\d+)"[^>]*>.*?</row>', re.S)
_PL_CF_RE = re.compile(r'<conditionalFormatting\b[^>]*sqref="([^"]+)"[^>]*>.*?</conditionalFormatting>', re.S)
_PL_AF_RE = re.compile(r'(<autoFilter\b[^>]*\bref=")([A-Z]+)(\d+):([A-Z]+)(\d+)(")')
_PL_DIM_RE = re.compile(r'(<dimension ref=")([A-Z]+)(\d+):([A-Z]+)(\d+)(")')
_PL_FORMULA_RE = re.compile(r'<f[ >]')
_PL_MERGE_RE = re.compile(r'<mergeCell\b')
_PL_CELL_REF_TMPL = r'(?<=r=")([A-Z]{1,3})%d(?=")'


def _pl_html_unescape(text):
    """Мини-раскодирование XML-сущностей (для имён листов)."""
    return (text.replace('&lt;', '<').replace('&gt;', '>')
                .replace('&quot;', '"').replace('&apos;', "'").replace('&amp;', '&'))


def _pl_parse_ref(ref):
    """'D66' → ('D', 66); при неудаче — (None, None)."""
    m = re.match(r'^([A-Z]+)(\d+)$', ref)
    return (m.group(1), int(m.group(2))) if m else (None, None)


def _pl_rel_map(rels_xml):
    """Карта связей Id → Target (порядок атрибутов в файле может быть любым)."""
    result = {}
    for tag in re.findall(r'<Relationship\b[^>]*/>', rels_xml or ''):
        rid = re.search(r'\bId="([^"]+)"', tag)
        target = re.search(r'\bTarget="([^"]+)"', tag)
        if rid and target:
            result[rid.group(1)] = target.group(1)
    return result


def _pl_resolve_sheet_path(zf, sheet_name=None):
    """Путь XML нужного листа внутри архива (по имени, иначе — первый лист)."""
    try:
        wb_xml = zf.read('xl/workbook.xml').decode('utf-8')
        rels_xml = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    except KeyError:
        return 'xl/worksheets/sheet1.xml'

    pairs = []
    for tag in re.findall(r'<sheet\b[^>]*?>', wb_xml):
        m_name = re.search(r'\bname="([^"]*)"', tag)
        m_rid = re.search(r'\br:id="([^"]*)"', tag)
        if m_name and m_rid:
            pairs.append((_pl_html_unescape(m_name.group(1)), m_rid.group(1)))
    if not pairs:
        return 'xl/worksheets/sheet1.xml'

    rel_map = _pl_rel_map(rels_xml)
    target_id = None
    if sheet_name:
        for name, rid in pairs:
            if name == sheet_name:
                target_id = rid
                break
    if target_id is None:
        target_id = pairs[0][1]

    target = rel_map.get(target_id, 'worksheets/sheet1.xml').lstrip('/')
    return target if target.startswith('xl/') else 'xl/' + target


_PL_MERGE_ELEM_RE = re.compile(r'<mergeCell\b[^>]*?/>')
_PL_MERGE_CELLS_RE = re.compile(r'<mergeCells\b[^>]*>.*?</mergeCells>', re.S)


def _pl_clean_remove_rows(remove_rows, protect_upto=1):
    """Убирает мусор из списка строк (1-based) и защищает строки 1..protect_upto (шапку)."""
    clean_rows = []
    for value in (remove_rows or []):
        try:
            num = int(str(value).strip())
        except (TypeError, ValueError):
            continue
        if num > protect_upto:
            clean_rows.append(num)
    return sorted(set(clean_rows))


def _pl_shift_row(num, remove_rows):
    """Номер строки после удаления: вычитаем все удалённые строки, стоящие выше."""
    return num - sum(1 for d in remove_rows if d < num)


def _pl_cell_xml(ref, style, value):
    """XML ячейки с текстом (стиль — номер оформления из книги)."""
    attrs = f' s="{style}"' if style else ''
    return (f'<c r="{ref}"{attrs} t="inlineStr">'
            f'<is><t xml:space="preserve">{_pl_xml_escape(value)}</t></is></c>')


def _pl_set_cell_value(body, column, row_num, value):
    """Меняет текст одной ячейки в теле строки, сохраняя её оформление.

    Если ячейки в строке нет — создаёт её в нужном месте (стиль берётся у соседней слева).
    """
    ref = f'{column}{row_num}'
    safe_col = re.escape(column)
    style = ''

    # 1) Пустая ячейка вида <c r="M5" s="7"/>
    empty_match = re.search(rf'<c r="{safe_col}{row_num}"[^>]*/>', body)
    # 2) Ячейка со значением вида <c r="M5" s="7" t="s"><v>3</v></c>
    filled_match = re.search(rf'<c r="{safe_col}{row_num}"[^>]*>(?:(?!<c ).)*?</c>', body, re.S)

    match = filled_match or empty_match
    if match:
        style_match = re.search(r'\bs="(\d+)"', match.group(0))
        if style_match:
            style = style_match.group(1)
        return body[:match.start()] + _pl_cell_xml(ref, style, value) + body[match.end():]

    # 3) Ячейки нет — создаём: ищем место по номеру колонки и берём стиль соседа слева
    own_index = openpyxl.utils.column_index_from_string(column)
    insert_at, neighbour_style = None, ''
    for cell_match in re.finditer(r'<c r="([A-Z]{1,3})\d+"[^>]*>', body):
        tag = cell_match.group(0)
        try:
            cell_index = openpyxl.utils.column_index_from_string(cell_match.group(1))
        except ValueError:
            continue
        if cell_index < own_index:
            style_match = re.search(r'\bs="(\d+)"', tag)
            if style_match:
                neighbour_style = style_match.group(1)
            continue
        insert_at = cell_match.start()
        break
    new_cell = _pl_cell_xml(ref, neighbour_style, value)
    if insert_at is None:
        return body + new_cell
    return body[:insert_at] + new_cell + body[insert_at:]


def _pl_set_sheet_cells(xml, edits):
    """Записывает новые значения в ячейки листа: edits = {номер строки: {колонка: текст}}."""
    if not edits:
        return xml

    def fix_row(match):
        row_num = int(match.group(1))
        row_edits = edits.get(row_num)
        if not row_edits:
            return match.group(0)
        block = match.group(0)
        tag_end = block.index('>')
        open_tag, body = block[:tag_end + 1], block[tag_end + 1:]
        for column, value in row_edits.items():
            body = _pl_set_cell_value(body, column, row_num, value)
        return open_tag + body

    return _PL_ROW_RE.sub(fix_row, xml)


def _pl_build_cell_edits(row_edits, columns):
    """Правки «Тип оплаты»/«Детальный лимит» из браузера → {строка: {колонка: текст}}."""
    if not isinstance(row_edits, dict) or not isinstance(columns, dict):
        return {}
    result = {}
    for row_key, edit in row_edits.items():
        try:
            row_num = int(str(row_key).strip())
        except (TypeError, ValueError):
            continue
        if not isinstance(edit, dict):
            continue
        cells = {}
        for field in ('category', 'detail'):
            column = str(columns.get(field) or '').strip().upper()
            if not re.fullmatch(r'[A-Z]{1,3}', column) or field not in edit:
                continue
            cells[column] = '' if edit.get(field) is None else str(edit.get(field))
        if cells:
            result[row_num] = cells
    return result


def _pl_shift_merges(xml, remove_rows):
    """Переносит объединённые ячейки вместе со строками; полностью удалённые убирает."""
    if '<mergeCell' not in xml:
        return xml

    def fix(mm):
        elem = mm.group(0)
        ref = re.search(r'\bref="([A-Z]+\d+):([A-Z]+\d+)"', elem)
        if not ref:
            return elem
        ca, ra = _pl_parse_ref(ref.group(1))
        cb, rb = _pl_parse_ref(ref.group(2))
        if ra is None or rb is None:
            return elem
        if not any(ra <= d <= rb for d in remove_rows):
            return elem
        ns = ra - sum(1 for d in remove_rows if d < ra)
        ne = rb - sum(1 for d in remove_rows if d <= rb)
        if ne < ns:
            return ''
        new_ref = f'{ca}{ns}' if ns == ne else f'{ca}{ns}:{cb}{ne}'
        return elem.replace(f'ref="{ref.group(1)}:{ref.group(2)}"', f'ref="{new_ref}"')

    xml = _PL_MERGE_ELEM_RE.sub(fix, xml)

    def fix_count(mm):
        body = mm.group(0)
        count = body.count('<mergeCell ')
        if not count:
            return ''
        return re.sub(r'<mergeCells\b[^>]*>', f'<mergeCells count="{count}">', body, count=1)

    return _PL_MERGE_CELLS_RE.sub(fix_count, xml)


def _pl_filter_rows_xml(xml, remove_rows):
    """Удаляет строки из XML листа с перенумерацией строк и адресов ячеек (как это делает Excel).

    Возвращает (новый XML, info). Если удалять нечего — XML отдаётся без изменений.
    """
    remove_rows = sorted(set(remove_rows or []))
    row_matches = list(_PL_ROW_RE.finditer(xml))
    warnings = []
    if _PL_FORMULA_RE.search(xml):
        warnings.append('в листе есть формулы — ссылки при удалении строк не пересчитываются')
    if _PL_MERGE_RE.search(xml):
        warnings.append('в листе есть объединённые ячейки — они перенесены вместе со строками')
    last_row = max((int(m.group(1)) for m in row_matches), default=1)

    if not remove_rows:
        return xml, {'removed': 0, 'rows_left': last_row,
                     'warning': '; '.join(warnings), 'changed': False}
    if not row_matches:
        raise ValueError('В листе не найдено строк для обработки')

    # 1. Удаление строк + перенумерация строк и адресов ячеек
    parts, kept_numbers = [], []
    for m in row_matches:
        num = int(m.group(1))
        if num in remove_rows:
            continue
        new_num = _pl_shift_row(num, remove_rows)
        kept_numbers.append(new_num)
        block = m.group(0)
        tag_end = block.index('>')
        open_tag, body = block[:tag_end + 1], block[tag_end + 1:]
        open_tag = re.sub(r'\br="\d+"', f'r="{new_num}"', open_tag, count=1)  # сам <row r="N">
        body = re.sub(_PL_CELL_REF_TMPL % num, lambda mm: mm.group(1) + str(new_num), body)  # ячейки
        parts.append(open_tag + body)

    new_last = max(kept_numbers) if kept_numbers else 1
    xml = xml[:row_matches[0].start()] + ''.join(parts) + xml[row_matches[-1].end():]

    # 2. dimension
    if _PL_DIM_RE.search(xml):
        xml = _PL_DIM_RE.sub(
            lambda mm: f'{mm.group(1)}{mm.group(2)}{mm.group(3)}:{mm.group(4)}{new_last}{mm.group(6)}',
            xml)

    # 3. Условное форматирование — диапазоны сдвигаются так же, как при удалении строк в Excel
    if _PL_CF_RE.search(xml):
        def fix_cf(mm):
            sqref = mm.group(1)
            new_ranges = []
            for rng in sqref.split(' '):
                a, _, b = rng.partition(':') if ':' in rng else (rng, ':', rng)
                ca, ra = _pl_parse_ref(a)
                cb, rb = _pl_parse_ref(b)
                if ra is None or rb is None:
                    new_ranges.append(rng)
                    continue
                ns = ra - sum(1 for d in remove_rows if d < ra)
                ne = rb - sum(1 for d in remove_rows if d <= rb)
                if ne < ns:
                    continue  # диапазон полностью удалён
                new_ranges.append(f'{ca}{ns}' if ns == ne else f'{ca}{ns}:{cb}{ne}')
            if not new_ranges:
                return ''
            return mm.group(0).replace(f'sqref="{sqref}"', 'sqref="' + ' '.join(new_ranges) + '"')
        xml = _PL_CF_RE.sub(fix_cf, xml)

    # 4. Автофильтр — тоже переносим на оставшиеся строки
    if _PL_AF_RE.search(xml):
        def fix_af(mm):
            start, end = int(mm.group(3)), int(mm.group(5))
            new_start = start - sum(1 for d in remove_rows if d < start)
            new_end = end - sum(1 for d in remove_rows if d <= end)
            if new_end < 1 or new_start < 1:
                return mm.group(0)
            return f'{mm.group(1)}{mm.group(2)}{new_start}:{mm.group(4)}{new_end}{mm.group(6)}'
        xml = _PL_AF_RE.sub(fix_af, xml)

    # 5. Объединённые ячейки — переносим вместе со строками
    xml = _pl_shift_merges(xml, remove_rows)

    return xml, {'removed': len(remove_rows), 'rows_left': new_last,
                 'warning': '; '.join(warnings), 'changed': True}


def _pl_delete_rows_from_xlsx(raw, remove_rows, sheet_name=None, protect_upto=1, cell_edits=None):
    """Удаляет строки (1-based, как в Excel) из листа xlsx.

    raw — байты исходного .xlsx; возвращает (BytesIO с готовым файлом, info).
    Строки 1..protect_upto считаются шапкой и не удаляются (по умолчанию — только строка 1).
    cell_edits — правки ячеек {номер строки: {колонка: текст}} («Тип оплаты», «Детальный лимит»).
    """
    remove_rows = _pl_clean_remove_rows(remove_rows, protect_upto)

    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        infos = z.infolist()
        data = {i.filename: z.read(i.filename) for i in infos}
        sheet_path = _pl_resolve_sheet_path(z, sheet_name)

    if not remove_rows and not cell_edits:
        return io.BytesIO(raw), {'removed': 0, 'sheet': sheet_path, 'rows_left': None,
                                 'byte_identical': True, 'warning': ''}

    if sheet_path not in data:
        raise ValueError('В файле не найден лист с данными')

    xml = _pl_set_sheet_cells(data[sheet_path].decode('utf-8'), cell_edits)
    xml, info = _pl_filter_rows_xml(xml, remove_rows)
    data[sheet_path] = xml.encode('utf-8')

    # Пересборка архива: тот же порядок записей и способ сжатия
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for entry in infos:
            zi = zipfile.ZipInfo(entry.filename, date_time=entry.date_time)
            zi.compress_type = entry.compress_type
            zi.external_attr = entry.external_attr
            zi.internal_attr = entry.internal_attr
            zi.create_system = entry.create_system
            zi.flag_bits = entry.flag_bits & ~0x08
            z.writestr(zi, data[entry.filename])
    out.seek(0)
    return out, {'removed': info['removed'], 'sheet': sheet_path, 'rows_left': info['rows_left'],
                 'byte_identical': False, 'warning': info['warning']}


# ============================================================
# ЭКСПОРТ РЕЕСТРА «ЛИМИТЫ ПЛАТЕЖЕЙ» ПО ВКЛАДКАМ-КАТЕГОРИЯМ
# ============================================================
# Первая вкладка «Общий свод» — весь список платежей без сводной таблицы.
# Далее один и тот же исходный лист копируется столько раз, сколько категорий
# встретилось в колонке «Группа» (в образце — колонка O). Строки каждой копии —
# только платежи своей категории (шапка сохраняется), под данными дописывается
# сводная таблица блока «Пояснения». Оформление исходника (шрифты, заливки, форматы,
# ширины, условное форматирование, настройки печати) не пересобирается, а копируется.
_PL_GROUP_CATEGORIES = [
    'Филиалы',
    'УК',
    'Питание, проживание, связь',
    'Оплаты в рамках 100 т.р.',
    'Непереносимые оплаты',
    'Графики',
    'Согласовано в оплату',
    'Кузьмина'
]
# Вкладка для платежей, у которых в колонке «Группа» нет знакомого названия
_PL_GROUP_OTHER = 'Прочее (не сопоставлено)'
# Вкладка со всем списком платежей (идёт первой, без сводной таблицы)
_PL_GROUP_ALL = 'Общий свод'
# Сколько пустых строк оставляем между данными и сводной таблицей (как в «Оплатах поставщикам»)
_PL_PIVOT_GAP = 4
_PL_SHEET_NAME_LIMIT = 31
_PL_SHEET_BAD_CHARS = re.compile(r'[\[\]\*\?/\\:]')
_PL_PIVOT_SUB_WIDTH = 18
_PL_PIVOT_EXP_WIDTH = 50


def _pl_xml_escape(text):
    """Текст → XML (как в файле Excel)."""
    return (str('' if text is None else text).replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;').replace("'", '&apos;'))


def _pl_letter(index):
    """Номер колонки → буква: 1 → A, 27 → AA."""
    return openpyxl.utils.get_column_letter(max(int(index), 1))


def _pl_num_text(value):
    """Число для ячейки Excel (без лишних знаков после запятой)."""
    try:
        num = float(value or 0)
    except (TypeError, ValueError):
        num = 0.0
    if num == int(num):
        return str(int(num))
    return ('%.2f' % num)


def _pl_fold(text):
    """Сравнение названий: без учёта регистра и «ё/е» — как в интерфейсе."""
    return str('' if text is None else text).strip().lower().replace('ё', 'е')


def _pl_norm_group(raw):
    """Название категории по значению колонки «Группа» (те же правила, что на экране)."""
    s = _pl_fold(raw)
    if not s:
        return ''
    for name in _PL_GROUP_CATEGORIES:
        if _pl_fold(name) == s:
            return name
    for name in _PL_GROUP_CATEGORIES:
        n = _pl_fold(name)
        if n and (n in s or s in n):
            return name
    return ''


def _pl_group_column_index(headers):
    """Номер колонки «Группа» (0-based): по названию, иначе — самая правая колонка."""
    for index, header in enumerate(headers):
        if _pl_fold(header) == 'группа':
            return index
    for index, header in enumerate(headers):
        if 'групп' in _pl_fold(header):
            return index
    return max(len(headers) - 1, 0)


def _pl_read_row_groups(raw, sheet_name, xml_last_row):
    """Читает лист: строка шапки, значения колонки «Группа» по строкам, непустые строки."""
    workbook = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    try:
        name = sheet_name if sheet_name in workbook.sheetnames else workbook.sheetnames[0]
        sheet = workbook[name]
        header_row, headers, group_col = None, [], None
        group_by_row, nonempty_rows = {}, set()
        for index, values in enumerate(
                sheet.iter_rows(min_row=1, max_row=max(int(xml_last_row), 1), values_only=True), start=1):
            has_value = any(v is not None and str(v).strip() != '' for v in values)
            if header_row is None:
                if not has_value:
                    continue
                header_row = index
                headers = ['' if v is None else str(v).strip() for v in values]
                group_col = _pl_group_column_index(headers)
                continue
            if not has_value:
                continue
            nonempty_rows.add(index)
            raw_group = values[group_col] if group_col is not None and group_col < len(values) else ''
            group_by_row[index] = _pl_norm_group(raw_group)
    finally:
        workbook.close()
    return {'header_row': header_row or 1, 'headers': headers, 'group_col': group_col,
            'group_by_row': group_by_row, 'nonempty_rows': nonempty_rows}


def _pl_insert_before_first(xml, element, anchors):
    """Вставляет кусок XML перед первым из указанных тегов (порядок тегов в листе важен)."""
    position = -1
    for anchor in anchors:
        found = xml.find(anchor)
        if found != -1 and (position == -1 or found < position):
            position = found
    if position == -1:
        return xml
    return xml[:position] + element + xml[position:]


# Теги, после которых в XML листа обязаны идти объединённые ячейки
_PL_AFTER_SHEETDATA_ANCHORS = (
    '<conditionalFormatting', '<dataValidations', '<hyperlinks', '<printOptions',
    '<pageMargins', '<pageSetup', '<headerFooter', '<drawing', '<legacyDrawing', '</worksheet>'
)


def _pl_insert_merge(xml, ref):
    """Добавляет объединённую ячейку (ref) в лист."""
    if '<mergeCells' in xml:
        match = _PL_MERGE_CELLS_RE.search(xml)
        if match:
            body = match.group(0)
            count = body.count('<mergeCell ')
            new_body = body.replace('</mergeCells>', f'<mergeCell ref="{ref}"/></mergeCells>')
            new_body = re.sub(r'<mergeCells\b[^>]*>', f'<mergeCells count="{count + 1}">', new_body, count=1)
            return xml[:match.start()] + new_body + xml[match.end():]
    element = f'<mergeCells count="1"><mergeCell ref="{ref}"/></mergeCells>'
    return _pl_insert_before_first(xml, element, _PL_AFTER_SHEETDATA_ANCHORS)


def _pl_add_cols(xml, col_from, col_to, width):
    """Задаёт ширину колонок, которых нет в исходном оформлении (исходные не трогаем)."""
    if col_to < col_from:
        return xml
    cols_match = re.search(r'<cols>.*?</cols>', xml, re.S)
    covered = set()
    if cols_match:
        for col_tag in re.findall(r'<col\b[^>]*/>', cols_match.group(0)):
            min_attr = re.search(r'\bmin="(\d+)"', col_tag)
            max_attr = re.search(r'\bmax="(\d+)"', col_tag)
            if min_attr and max_attr:
                covered.update(range(int(min_attr.group(1)), int(max_attr.group(1)) + 1))

    pieces, run_start = [], None
    for column in range(col_from, col_to + 1):
        if column in covered:
            if run_start is not None:
                pieces.append((run_start, column - 1))
                run_start = None
        elif run_start is None:
            run_start = column
    if run_start is not None:
        pieces.append((run_start, col_to))
    if not pieces:
        return xml

    fragment = ''.join(f'<col min="{a}" max="{b}" width="{width}" customWidth="1"/>' for a, b in pieces)
    if cols_match:
        insert_at = cols_match.end() - len('</cols>')
        return xml[:insert_at] + fragment + xml[insert_at:]
    return _pl_insert_before_first(xml, f'<cols>{fragment}</cols>', ('<sheetData',))


def _pl_safe_sheet_name(name, used):
    """Имя вкладки, допустимое для Excel: без служебных символов, не длиннее 31 знака, уникальное."""
    clean = _PL_SHEET_BAD_CHARS.sub(' ', str(name or '').strip())
    clean = re.sub(r'\s+', ' ', clean).strip(" '")
    if not clean:
        clean = 'Лист'
    if len(clean) > _PL_SHEET_NAME_LIMIT:
        clean = clean[:_PL_SHEET_NAME_LIMIT].strip()
    candidate, suffix = clean, 2
    while candidate.lower() in used:
        tail = f' ({suffix})'
        candidate = clean[:_PL_SHEET_NAME_LIMIT - len(tail)].strip() + tail
        suffix += 1
    used.add(candidate.lower())
    return candidate


def _pl_workbook_sheets(wb_xml, rels_xml):
    """Список листов книги: [(имя, sheetId, rId, путь XML)] в порядке книги."""
    rel_map = _pl_rel_map(rels_xml)
    result = []
    for match in re.finditer(r'<sheet\b[^>]*/>', wb_xml):
        tag = match.group(0)
        name_attr = re.search(r'\bname="([^"]*)"', tag)
        rid_attr = re.search(r'\br:id="([^"]*)"', tag)
        id_attr = re.search(r'\bsheetId="(\d+)"', tag)
        if not (name_attr and rid_attr):
            continue
        target = rel_map.get(rid_attr.group(1), '').lstrip('/')
        path = target if target.startswith('xl/') else 'xl/' + target
        sheet_id = int(id_attr.group(1)) if id_attr else len(result) + 1
        result.append((_pl_html_unescape(name_attr.group(1)), sheet_id, rid_attr.group(1), path))
    return result


def _pl_update_app_xml(app_xml, sheet_names):
    """Обновляет список листов в свойствах файла (docProps/app.xml)."""
    names_xml = ''.join(f'<vt:lpstr>{_pl_xml_escape(name)}</vt:lpstr>' for name in sheet_names)
    titles = (f'<TitlesOfParts><vt:vector size="{len(sheet_names)}" baseType="lpstr">'
              f'{names_xml}</vt:vector></TitlesOfParts>')
    app_xml = re.sub(r'<TitlesOfParts>.*?</TitlesOfParts>', titles, app_xml, flags=re.S)
    if '<HeadingPairs>' in app_xml:
        headings = ('<HeadingPairs><vt:vector size="2" baseType="variant">'
                    '<vt:variant><vt:lpstr>Листы</vt:lpstr></vt:variant>'
                    f'<vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant>'
                    '</vt:vector></HeadingPairs>')
        app_xml = re.sub(r'<HeadingPairs>.*?</HeadingPairs>', headings, app_xml, flags=re.S)
    return app_xml


def _pl_pivot_styles(styles_xml):
    """Добавляет в стили книги оформление сводного блока. Возвращает (XML, номера стилей)."""
    xml = styles_xml

    def append_items(text, tag, items):
        count_match = re.search(rf'<{tag}\b[^>]*\bcount="(\d+)"', text)
        end = text.find(f'</{tag}>')
        if end == -1:
            return text, None
        base = int(count_match.group(1)) if count_match else 0
        text = text[:end] + ''.join(items) + text[end:]
        if count_match:
            text = (text[:count_match.start()] + f'<{tag} count="{base + len(items)}"'
                    + text[count_match.end():])
        return text, base

    fonts = [
        # заголовок над таблицей
        '<font><b/><sz val="14"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>',
        # шапка таблицы — белый жирный
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>',
        # строка ИТОГО
        '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>'
    ]
    fills = [
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>',
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>',
        '<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill>'
    ]
    xml, font_base = append_items(xml, 'fonts', fonts)
    xml, fill_base = append_items(xml, 'fills', fills)
    if font_base is None or fill_base is None:
        return styles_xml, None
    font_title, font_header, font_bold = font_base, font_base + 1, font_base + 2
    fill_blue, fill_yellow, fill_green = fill_base, fill_base + 1, fill_base + 2

    # Тонкая рамка: берём готовую из книги, иначе добавляем свою
    thin_border = None
    border_count = 0
    borders_match = re.search(r'<borders\b[^>]*>.*?</borders>', xml, re.S)
    if borders_match:
        border_items = re.findall(r'<border\b[^>]*>.*?</border>|<border\b[^>]*/>',
                                  borders_match.group(0), re.S)
        border_count = len(border_items)
        for index, item in enumerate(border_items):
            if all(re.search(rf'<{side}\b[^>]*style="thin"', item)
                   for side in ('left', 'right', 'top', 'bottom')):
                thin_border = index
                break
    if thin_border is None:
        border = ('<border><left style="thin"><color indexed="64"/></left>'
                  '<right style="thin"><color indexed="64"/></right>'
                  '<top style="thin"><color indexed="64"/></top>'
                  '<bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>')
        xml, border_base = append_items(xml, 'borders', [border])
        if border_base is None:
            return styles_xml, None
        thin_border = border_count or border_base

    xfs = [
        # 0 — заголовок над таблицей
        f'<xf numFmtId="0" fontId="{font_title}" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
        # 1 — шапка таблицы
        f'<xf numFmtId="0" fontId="{font_header}" fillId="{fill_blue}" borderId="{thin_border}"'
        f' xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
        f'<alignment horizontal="center" vertical="center" wrapText="1"/></xf>',
        # 2 — название контрагента
        f'<xf numFmtId="0" fontId="0" fillId="0" borderId="{thin_border}" xfId="0" applyBorder="1"/>',
        # 3 — название контрагента в подсвеченной (жёлтой) строке
        f'<xf numFmtId="0" fontId="0" fillId="{fill_yellow}" borderId="{thin_border}" xfId="0"'
        f' applyFill="1" applyBorder="1"/>',
        # 4 — число
        f'<xf numFmtId="4" fontId="0" fillId="0" borderId="{thin_border}" xfId="0"'
        f' applyNumberFormat="1" applyBorder="1" applyAlignment="1">'
        f'<alignment horizontal="right" vertical="top"/></xf>',
        # 5 — число в подсвеченной (жёлтой) строке
        f'<xf numFmtId="4" fontId="0" fillId="{fill_yellow}" borderId="{thin_border}" xfId="0"'
        f' applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">'
        f'<alignment horizontal="right" vertical="top"/></xf>',
        # 6 — «ИТОГО» в первой колонке
        f'<xf numFmtId="0" fontId="{font_bold}" fillId="{fill_green}" borderId="{thin_border}"'
        f' xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>',
        # 7 — число в строке ИТОГО
        f'<xf numFmtId="4" fontId="{font_bold}" fillId="{fill_green}" borderId="{thin_border}"'
        f' xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
        f'<alignment horizontal="right" vertical="top"/></xf>',
        # 8 — пустая ячейка «Пояснение» в строке ИТОГО
        f'<xf numFmtId="0" fontId="0" fillId="{fill_green}" borderId="{thin_border}" xfId="0"'
        f' applyFill="1" applyBorder="1"/>'
    ]
    xml, xf_base = append_items(xml, 'cellXfs', xfs)
    if xf_base is None:
        return styles_xml, None

    styles = {'title': xf_base, 'header': xf_base + 1, 'name': xf_base + 2, 'name_y': xf_base + 3,
              'num': xf_base + 4, 'num_y': xf_base + 5, 'total_name': xf_base + 6,
              'num_total': xf_base + 7, 'text_total': xf_base + 8}
    return xml, styles


def _pl_trim_cf_to_data(xml, last_row):
    """Не даём условному форматированию исходника заходить на сводную таблицу."""
    if '<conditionalFormatting' not in xml:
        return xml

    def fix(mm):
        sqref = mm.group(1)
        new_ranges = []
        for rng in sqref.split(' '):
            a, _, b = rng.partition(':') if ':' in rng else (rng, ':', rng)
            ca, ra = _pl_parse_ref(a)
            cb, rb = _pl_parse_ref(b)
            if ra is None or rb is None:
                new_ranges.append(rng)
                continue
            if ra > last_row:
                continue                      # диапазон целиком ниже данных — он бы красил сводную
            ne = min(rb, last_row)
            new_ranges.append(f'{ca}{ra}' if ra == ne else f'{ca}{ra}:{cb}{ne}')
        if not new_ranges:
            return ''
        return mm.group(0).replace(f'sqref="{sqref}"', 'sqref="' + ' '.join(new_ranges) + '"')

    return _PL_CF_RE.sub(fix, xml)


def _pl_append_pivot_block(xml, pivot, last_row, styles):
    """Дописывает под данными листа сводную таблицу блока «Пояснения»."""
    headers = ['' if h is None else str(h) for h in (pivot.get('headers') or [])]
    items = pivot.get('rows') or []
    columns = 1 + len(headers) + 2                     # Контрагент + подразделения + Итого + Пояснение
    title_row = last_row + _PL_PIVOT_GAP
    header_row = title_row + 2
    first_data_row = header_row + 1
    total_row = first_data_row + len(items)

    def row_xml(row_num, cells):
        return f'<row r="{row_num}" spans="1:{columns}">' + ''.join(cells) + '</row>'

    def cell_text(row_num, col, style, value):
        text = _pl_xml_escape(value)
        return (f'<c r="{_pl_letter(col)}{row_num}" s="{style}" t="inlineStr">'
                f'<is><t xml:space="preserve">{text}</t></is></c>')

    def cell_num(row_num, col, style, value):
        return f'<c r="{_pl_letter(col)}{row_num}" s="{style}"><v>{_pl_num_text(value)}</v></c>'

    def cell_empty(row_num, col, style):
        return f'<c r="{_pl_letter(col)}{row_num}" s="{style}"/>'

    rows_xml = [row_xml(title_row, [
        cell_text(title_row, 1, styles['title'], 'Пояснения: сводная по контрагентам')])]

    header_cells = [cell_text(header_row, 1, styles['header'], 'Контрагент')]
    for index, header in enumerate(headers, start=2):
        header_cells.append(cell_text(header_row, index, styles['header'], header))
    header_cells.append(cell_text(header_row, columns - 1, styles['header'], 'Итого'))
    header_cells.append(cell_text(header_row, columns, styles['header'], 'Пояснение'))
    rows_xml.append(row_xml(header_row, header_cells))

    totals = [0.0] * len(headers)
    grand = 0.0
    for offset, item in enumerate(items):
        row_num = first_data_row + offset
        explanation = str(item.get('explanation') or '')
        name_style = styles['name_y'] if explanation else styles['name']
        num_style = styles['num_y'] if explanation else styles['num']
        values = item.get('values') or []
        cells = [cell_text(row_num, 1, name_style, item.get('contractor') or '')]
        row_total = 0.0
        for index in range(len(headers)):
            raw_value = values[index] if index < len(values) else 0
            try:
                value = float(raw_value or 0)
            except (TypeError, ValueError):
                value = 0.0
            totals[index] += value
            row_total += value
            cells.append(cell_num(row_num, index + 2, num_style, value))
        cells.append(cell_num(row_num, columns - 1, num_style, row_total))
        cells.append(cell_text(row_num, columns, name_style, explanation))
        grand += row_total
        rows_xml.append(row_xml(row_num, cells))

    total_cells = [cell_text(total_row, 1, styles['total_name'], 'ИТОГО')]
    for index, value in enumerate(totals):
        total_cells.append(cell_num(total_row, index + 2, styles['num_total'], value))
    total_cells.append(cell_num(total_row, columns - 1, styles['num_total'], grand))
    total_cells.append(cell_empty(total_row, columns, styles['text_total']))
    rows_xml.append(row_xml(total_row, total_cells))

    end = xml.rfind('</sheetData>')
    if end == -1:
        return xml
    xml = xml[:end] + ''.join(rows_xml) + xml[end:]

    # Заголовок таблицы — объединённые ячейки по всей её ширине
    xml = _pl_insert_merge(xml, f'A{title_row}:{_pl_letter(columns)}{title_row}')

    # Подсветка исходника остаётся только на данных — не красит сводную таблицу
    xml = _pl_trim_cf_to_data(xml, last_row)

    # Ширины колонок правее данных (внутри данных остаются исходные)
    xml = _pl_add_cols(xml, 2, columns - 1, _PL_PIVOT_SUB_WIDTH)
    xml = _pl_add_cols(xml, columns, columns, _PL_PIVOT_EXP_WIDTH)

    # dimension — расширяем, если сводная таблица вышла за пределы данных
    dimension = _PL_DIM_RE.search(xml)
    if dimension:
        end_col, end_row = dimension.group(4), int(dimension.group(5))
        new_end_col = _pl_letter(max(openpyxl.utils.column_index_from_string(end_col), columns))
        new_end_row = max(end_row, total_row)
        xml = _PL_DIM_RE.sub(
            lambda mm: f'{mm.group(1)}{mm.group(2)}{mm.group(3)}:{new_end_col}{new_end_row}{mm.group(6)}',
            xml)
    return xml


def _pl_tab_rows_by_group(row_info, removed, header_row):
    """Разбивает строки файла на вкладки: общий список, категории «Пояснений» и «Прочее»."""
    removed_set = set(removed)
    kept = [r for r in sorted(row_info['nonempty_rows']) if r > header_row and r not in removed_set]
    if not kept:
        return []
    # Первая вкладка — весь список платежей (сводная таблица на неё не добавляется)
    tabs = [{'name': _PL_GROUP_ALL, 'rows': kept, 'pivot': False}]
    categories = []
    for name in _PL_GROUP_CATEGORIES:
        rows = [r for r in kept if row_info['group_by_row'].get(r) == name]
        if rows:
            categories.append({'name': name, 'rows': rows})
    tabs.extend(categories)
    others = [r for r in kept if not row_info['group_by_row'].get(r)]
    # «Прочее» нужно только тогда, когда есть с чем сравнивать (иначе всё уже в «Общем своде»)
    if others and categories:
        tabs.append({'name': _PL_GROUP_OTHER, 'rows': others})
    return tabs


def _pl_pivots_by_name(pivots):
    """Сводные таблицы из браузера: {название вкладки: описание таблицы}."""
    result = {}
    for item in (pivots or []):
        if isinstance(item, dict) and item.get('name'):
            result[str(item['name']).strip()] = item
    return result


def _pl_set_sheet_uid(xml):
    """Делает идентификатор листа уникальным (Excel хранит его для истории правок)."""
    return re.sub(r'xr:uid="\{[^"]*\}"', f'xr:uid="{{{str(uuid.uuid4()).upper()}}}"', xml, count=1)


def _pl_split_sheets_by_group(raw, remove_rows, sheet_name=None, pivots=None, cell_edits=None):
    """Раскладывает реестр по вкладкам: «Общий свод», категории «Группы» со сводными таблицами.

    cell_edits — правки «Тип оплаты»/«Детальный лимит» (пишутся во все вкладки).
    Возвращает (BytesIO, info) либо (None, None), если категории определить не удалось.
    """
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        infos = z.infolist()
        data = {i.filename: z.read(i.filename) for i in infos}
        sheet_path = _pl_resolve_sheet_path(z, sheet_name)

    if sheet_path not in data:
        raise ValueError('В файле не найден лист с данными')

    sheet_xml = _pl_set_sheet_cells(data[sheet_path].decode('utf-8'), cell_edits)
    xml_rows = [int(m.group(1)) for m in _PL_ROW_RE.finditer(sheet_xml)]
    if not xml_rows:
        raise ValueError('В листе не найдено строк для обработки')

    row_info = _pl_read_row_groups(raw, sheet_name, max(xml_rows))
    header_row = row_info['header_row']
    removed = _pl_clean_remove_rows(remove_rows, header_row)
    tabs = _pl_tab_rows_by_group(row_info, removed, header_row)
    if not tabs:
        return None, None

    pivots_map = _pl_pivots_by_name(pivots)

    # Оформление сводного блока добавляем в стили книги один раз
    styles = None
    if any(tab.get('pivot', True) and tab['name'] in pivots_map for tab in tabs) and 'xl/styles.xml' in data:
        styles_xml, styles = _pl_pivot_styles(data['xl/styles.xml'].decode('utf-8'))
        if styles:
            data['xl/styles.xml'] = styles_xml.encode('utf-8')

    warnings = []
    for tab in tabs:
        keep_set = set(tab['rows']) | set(range(1, header_row + 1))
        extra = [r for r in xml_rows if r not in keep_set]
        tab_xml, filter_info = _pl_filter_rows_xml(sheet_xml, extra)
        if filter_info.get('warning'):
            warnings.append(filter_info['warning'])
        if styles and tab.get('pivot', True) and pivots_map.get(tab['name']):
            tab_xml = _pl_append_pivot_block(tab_xml, pivots_map[tab['name']],
                                             filter_info['rows_left'], styles)
        tab['xml'] = _pl_set_sheet_uid(tab_xml)

    return _pl_pack_group_workbook(data, infos, sheet_path, tabs, removed, xml_rows, warnings)


def _pl_pack_group_workbook(data, infos, sheet_path, tabs, removed, xml_rows, warnings):
    """Собирает книгу: исходные части + листы-вкладки со сводными таблицами."""
    wb_xml = data['xl/workbook.xml'].decode('utf-8')
    rels_path = 'xl/_rels/workbook.xml.rels'
    rels_xml = data.get(rels_path, '').decode('utf-8')
    sheet_meta = _pl_workbook_sheets(wb_xml, rels_xml)
    max_rid = max([int(m) for m in re.findall(r'\bId="rId(\d+)"', rels_xml)] or [0])

    other_sheets = [meta for meta in sheet_meta if meta[3] != sheet_path]
    used_parts = set(data.keys())
    used_names = {meta[0].lower() for meta in other_sheets}
    next_sheet_num = 2
    entries, new_rels, new_overrides = [], [], []

    for index, tab in enumerate(tabs):
        tab['sheet_name'] = _pl_safe_sheet_name(tab['name'], used_names)
        if index == 0:
            # Первая вкладка занимает XML исходного листа
            part = sheet_path
            rid = next((meta[2] for meta in sheet_meta if meta[3] == sheet_path), None)
            if rid is None:
                max_rid += 1
                rid = f'rId{max_rid}'
                new_rels.append((rid, part))
                new_overrides.append(part)
        else:
            while f'xl/worksheets/sheet{next_sheet_num}.xml' in used_parts:
                next_sheet_num += 1
            part = f'xl/worksheets/sheet{next_sheet_num}.xml'
            used_parts.add(part)
            next_sheet_num += 1
            max_rid += 1
            rid = f'rId{max_rid}'
            new_rels.append((rid, part))
            new_overrides.append(part)
        tab['part'] = part
        data[part] = tab['xml'].encode('utf-8')
        entries.append((tab['sheet_name'], rid, part))

    # Прочие листы исходной книги оставляем без изменений — после вкладок
    for name, _sheet_id, rid, part in other_sheets:
        entries.append((name, rid, part))

    # Настройки печати: связи исходного листа копируем новым вкладкам
    base_rels_name = f'xl/worksheets/_rels/{os.path.basename(sheet_path)}.rels'
    base_rels = data.get(base_rels_name)
    for tab in tabs[1:]:
        if base_rels and 'r:id="' in tab['xml']:
            data[f'xl/worksheets/_rels/{os.path.basename(tab["part"])}.rels'] = base_rels

    # workbook.xml — новый список листов
    sheets_xml = ''.join(
        f'<sheet name="{_pl_xml_escape(name)}" sheetId="{number}" r:id="{rid}"/>'
        for number, (name, rid, _part) in enumerate(entries, start=1))
    data['xl/workbook.xml'] = re.sub(
        r'<sheets>.*?</sheets>', f'<sheets>{sheets_xml}</sheets>', wb_xml, flags=re.S).encode('utf-8')

    # Связи книги и типы содержимого для новых листов
    if new_rels:
        additions = ''.join(
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            f'relationships/worksheet" Target="{part[len("xl/"):]}"/>' for rid, part in new_rels)
        data[rels_path] = rels_xml.replace('</Relationships>', additions + '</Relationships>').encode('utf-8')
    if new_overrides and '[Content_Types].xml' in data:
        content_types = data['[Content_Types].xml'].decode('utf-8')
        for part in new_overrides:
            content_types = content_types.replace(
                '</Types>',
                f'<Override PartName="/{part}" ContentType="application/vnd.openxmlformats-'
                f'officedocument.spreadsheetml.worksheet+xml"/></Types>')
        data['[Content_Types].xml'] = content_types.encode('utf-8')

    if 'docProps/app.xml' in data:
        data['docProps/app.xml'] = _pl_update_app_xml(
            data['docProps/app.xml'].decode('utf-8'),
            [name for name, _rid, _part in entries]).encode('utf-8')

    # Пересборка архива: исходный порядок частей + новые части
    out = io.BytesIO()
    written = set()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for entry in infos:
            zi = zipfile.ZipInfo(entry.filename, date_time=entry.date_time)
            zi.compress_type = entry.compress_type
            zi.external_attr = entry.external_attr
            zi.internal_attr = entry.internal_attr
            zi.create_system = entry.create_system
            zi.flag_bits = entry.flag_bits & ~0x08
            z.writestr(zi, data[entry.filename])
            written.add(entry.filename)
        for name, blob in data.items():
            if name not in written:
                z.writestr(name, blob)
    out.seek(0)

    info = {
        'removed': len(removed), 'sheet': sheet_path, 'rows_left': max(xml_rows),
        'byte_identical': False, 'warning': '; '.join(dict.fromkeys(warnings)),
        'tabs': [{'name': tab['sheet_name'], 'rows': len(tab['rows'])} for tab in tabs]
    }
    return out, info


def _plp_build_rows(rows, month_first, period, stage):
    """Преобразует строки заявок из запроса в записи для payment_limits_payments."""
    payments = []
    for row in rows or []:
        payments.append({
            'limit_month': month_first,
            'period': period,
            'stage': stage,
            'article': (row.get('article') or '').strip(),
            'detail_limit': (row.get('detail_limit') or '').strip(),
            'amount': float(row.get('amount') or 0),
            'recipient': (row.get('recipient') or '').strip(),
            'request_no': (row.get('request_no') or '').strip(),
            'request_date': _parse_plp_date(row.get('request_date')),
            'status': (row.get('status') or '').strip(),
            'organization': (row.get('organization') or '').strip(),
            'subdivision': (row.get('subdivision') or '').strip(),
            'priority': (row.get('priority') or '').strip(),
            'purpose': (row.get('purpose') or '').strip(),
            'payment_date': _parse_plp_date(row.get('payment_date')),
            'applicant': (row.get('applicant') or '').strip(),
            'payment_type': (row.get('payment_type') or '').strip()
        })
    return payments


@app.route('/api/payment-limits/save', methods=['POST'])
@limiter.limit("20 per minute")
def api_payment_limits_save():
    """
    Сохранение варианта загрузки (2 стадии подтверждения).

    multipart/form-data:
      file — исходный .xlsx (нужен хотя бы один раз: без него нельзя выгрузить финальный файл)
      data — JSON: {stage: 'draft'|'final', rows: [...], removedRows: [...],
                    rowEdits: {«номер строки»: {category, detail}},   ← правки «Тип оплаты»/«Детальный лимит»
                    period: 1..4, sheetName: '', fileName: '', author: '', registryId: N}

    Черновик сохраняет первичный вариант, финал — итоговый после проверки коллегой.
    Оба варианта и исходный файл хранятся в БД и обнуляются по истечении месяца.
    """
    try:
        try:
            data = json.loads(request.form.get('data') or '{}')
        except Exception:
            data = {}

        stage = (data.get('stage') or db.PLP_STAGE_FINAL).strip().lower()
        if stage not in (db.PLP_STAGE_DRAFT, db.PLP_STAGE_FINAL):
            return jsonify({'success': False, 'error': 'Неверная стадия сохранения'}), 400

        rows = data.get('rows') or []
        removed_rows = data.get('removedRows') or []
        row_edits = data.get('rowEdits')
        if not isinstance(row_edits, dict):
            row_edits = {}
        file = request.files.get('file')
        file_bytes = file.read() if file else None

        now = datetime.now()
        month_first = now.strftime('%Y-%m') + '-01'

        period = None
        try:
            if data.get('period') is not None:
                period = int(data.get('period'))
        except (TypeError, ValueError):
            period = None
        if period is None or period < 1 or period > 4:
            period = payment_period_for_date(now)

        # Автоочистка: начало нового месяца — начинаем учёт заново
        db.cleanup_payment_limits_before(month_first)

        # Партия (загрузка): если registryId не передан — это новая загрузка реестра,
        # её строки будут накапливаться вместе с ранее сохранёнными загрузками периода.
        registry_id = None
        try:
            if data.get('registryId') is not None:
                registry_id = int(data.get('registryId'))
        except (TypeError, ValueError):
            registry_id = None

        # Исходный файл: для новой загрузки он обязателен
        has_file = any(r.get('has_file') for r in db.get_payment_limits_registry(month_first, period))
        if not file_bytes and registry_id is None:
            return jsonify({'success': False,
                            'error': 'Нет исходного файла реестра: загрузите файл ещё раз, '
                                     'чтобы сохранить вариант'}), 400
        if not file_bytes and registry_id is not None:
            reg_meta = db.get_payment_limits_registry_by_id(registry_id)
            if not reg_meta or not reg_meta.get('has_file'):
                return jsonify({'success': False,
                                'error': 'Нет исходного файла реестра: загрузите файл ещё раз, '
                                         'чтобы сохранить вариант'}), 400

        reg_result = db.save_payment_limits_registry(
            month=month_first,
            period=period,
            stage=stage,
            file_name=(data.get('fileName') or '').strip(),
            sheet_name=(data.get('sheetName') or '').strip(),
            file_bytes=file_bytes,
            row_count=int(data.get('rowCount') or len(rows) or 0),
            removed_rows=removed_rows,
            author=(data.get('author') or '').strip(),
            registry_id=registry_id,
            row_edits=row_edits
        )
        if not reg_result.get('success'):
            return jsonify(reg_result), 500

        save_result = db.save_payment_limits_stage(
            month_first, period, stage, _plp_build_rows(rows, month_first, period, stage),
            registry_id=reg_result.get('registry_id'))
        if not save_result.get('success'):
            return jsonify(save_result), 500

        total_amount = sum(float(r.get('amount') or 0) for r in rows)
        stage_label = 'черновик' if stage == db.PLP_STAGE_DRAFT else 'финальный вариант'
        logger.info(f"✅ Сохранён {stage_label} лимитов: {save_result['count']} заявок, "
                    f"снято {len(removed_rows)}, правок категорий {len(row_edits)}, "
                    f"сумма {total_amount:.2f} ₽, P{period}, "
                    f"загрузка #{reg_result.get('registry_id')}, месяц {month_first}")

        return jsonify({
            'success': True,
            'stage': stage,
            'month': month_first,
            'period': period,
            'count': save_result['count'],
            'removed_count': len(removed_rows),
            'row_edits_count': len(row_edits),
            'contractors_learned': int(save_result.get('contractors_learned') or 0),
            'total_amount': total_amount,
            'registry_id': reg_result.get('registry_id'),
            'new_registry': bool(reg_result.get('created')),
            'saved_at': now.isoformat(timespec='seconds'),
            'message': f"Сохранён {stage_label}: {save_result['count']} заявок "
                       f"(период P{period}, {month_first[:7]})"
        })
    except Exception as e:
        logger.error(f"Ошибка API payment-limits/save: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/payment-limits/reset', methods=['POST'])
@limiter.limit("10 per minute")
def api_payment_limits_reset():
    """
    Обнуление ранее загруженных данных контроля лимитов за месяц.

    Тело: JSON {'month': 'YYYY-MM' (необязательно), 'period': 1..4 (необязательно),
                'stage': 'draft'|'final' (необязательно)}.
    Удаляются сохранённые варианты (черновик/финал), платежи и исходный файл реестра,
    чтобы можно было пройти полный цикл (2 стадии) заново.
    """
    try:
        data = request.get_json(force=True, silent=True) or {}
        now = datetime.now()
        month = (data.get('month') or now.strftime('%Y-%m'))
        month_first = str(month)[:7] + '-01'

        period = data.get('period')
        try:
            period = int(period) if period not in (None, '') else None
        except (TypeError, ValueError):
            period = None
        if period is not None and (period < 1 or period > 4):
            period = None

        stage = (data.get('stage') or '').strip().lower()
        if stage not in (db.PLP_STAGE_DRAFT, db.PLP_STAGE_FINAL):
            stage = None

        result = db.reset_payment_limits(month_first, period, stage)
        if not result.get('success'):
            return jsonify(result), 500

        logger.info(f"✅ Обнулены данные лимитов: месяц {month_first}"
                    + (f", P{period}" if period else "")
                    + (f", стадия «{stage}»" if stage else "")
                    + f" — платежей {result['payments_deleted']}, реестров {result['registries_deleted']}")

        return jsonify({
            'success': True,
            'month': month_first,
            'period': period,
            'stage': stage,
            'payments_deleted': result['payments_deleted'],
            'registries_deleted': result['registries_deleted'],
            'message': 'Данные обнулены: удалено платежей '
                       f"{result['payments_deleted']} и реестров {result['registries_deleted']}. "
                       'Можно загрузить файл и пройти обе стадии заново'
        })
    except Exception as e:
        logger.error(f"Ошибка API payment-limits/reset: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/payment-limits/contractor-categories', methods=['GET'])
def api_payment_limits_contractor_categories():
    """
    Справочник «контрагент → тип оплаты и детальный лимит» из прошлых загрузок.

    Нужен браузеру, чтобы при загрузке файла автоматически подставлять категории
    тем контрагентам, которых уже относили к лимитам ранее.
    """
    try:
        categories = db.get_contractor_categories()
        logger.info(f"Справочник контрагентов выдан: {len(categories)} записей")
        return jsonify({'success': True, 'count': len(categories), 'categories': categories})
    except Exception as e:
        logger.error(f"Ошибка API payment-limits/contractor-categories: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/payment-limits/registry', methods=['GET'])
def api_payment_limits_registry():
    """
    Состояние реестров за месяц: сохранённые варианты (черновик/финал) и исходный файл.

    Query: month=YYYY-MM (по умолчанию текущий), period=1..4 (необязательно).
    """
    try:
        now = datetime.now()
        month = request.args.get('month') or now.strftime('%Y-%m')
        month_first = month[:7] + '-01'
        period = request.args.get('period')

        # Начало нового месяца — учёт заново (платежи и реестры прошлых месяцев удаляются)
        db.cleanup_payment_limits_before(month_first)

        registries = db.get_payment_limits_registry(month_first, int(period) if period else None)
        result = []
        for r in registries:
            item = dict(r)
            for key in ('draft_saved_at', 'final_saved_at', 'final_exported_at', 'created_at', 'updated_at', 'limit_month'):
                val = item.get(key)
                if val is not None and hasattr(val, 'isoformat'):
                    item[key] = val.isoformat(timespec='seconds') if key.endswith('_at') else val.isoformat()
            item['file_size'] = int(item.get('file_size') or 0)
            item['draft_removed_rows'] = item.get('draft_removed_rows') or []
            item['final_removed_rows'] = item.get('final_removed_rows') or []
            # Правки «Тип оплаты»/«Детальный лимит» по строкам файла (ключ — номер строки Excel)
            item['draft_row_edits'] = item.get('draft_row_edits') or {}
            item['final_row_edits'] = item.get('final_row_edits') or {}
            result.append(item)

        return jsonify({
            'success': True,
            'month': month_first,
            'period': payment_period_for_date(now),
            'registries': result
        })
    except Exception as e:
        logger.error(f"Ошибка API payment-limits/registry: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/payment-limits/registry/file', methods=['GET'])
def api_payment_limits_registry_file():
    """Исходный файл сохранённого реестра (для просмотра/правки на другом компьютере)."""
    try:
        registry_id = request.args.get('registry_id')
        if not registry_id:
            return jsonify({'error': 'Не указан registry_id'}), 400
        file_name, file_bytes = db.get_payment_limits_registry_file(int(registry_id))
        if not file_bytes:
            return jsonify({'error': 'Исходный файл реестра не найден'}), 404
        return send_file(
            io.BytesIO(file_bytes),
            as_attachment=True,
            download_name=file_name or f'Реестр_{registry_id}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        logger.error(f"Ошибка API payment-limits/registry/file: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/payment-limits/export', methods=['POST'])
@limiter.limit("10 per minute")
def api_payment_limits_export():
    """
    Экспорт реестра заявок в исходном формате: исходный .xlsx без снятых чекбоксом строк.

    Источник файла:
      * загруженный в запросе файл (multipart 'file'); либо
      * сохранённый в БД реестр (data.registryId) — тогда файл из браузера не нужен.
    Снятые строки: data.removeRows (текущие отметки) либо сохранённый набор нужной стадии
    (data.stage = 'final' | 'draft'; по умолчанию финал, если он сохранён).
    Разбивка по вкладкам: data.splitByGroup = true + data.groupSheets — сводные таблицы
    блока «Пояснения» (браузер считает их так же, как показывает на экране).
    """
    try:
        try:
            data = json.loads(request.form.get('data') or '{}')
        except Exception:
            data = {}

        registry_id = data.get('registryId')
        stage = (data.get('stage') or '').strip().lower()
        sheet_name = (data.get('sheetName') or '').strip() or None
        remove_rows = data.get('removeRows')
        split_by_group = bool(data.get('splitByGroup'))
        group_sheets = data.get('groupSheets')
        if not isinstance(group_sheets, list):
            group_sheets = None
        # Правки «Тип оплаты»/«Детальный лимит» → изменения конкретных ячеек листа
        cell_edits = _pl_build_cell_edits(data.get('rowEdits'), data.get('editColumns'))
        file_name = None
        used_stage = None
        used_split = False

        if 'file' in request.files:
            file = request.files['file']
            raw = file.read()
            file_name = file.filename
        elif registry_id:
            meta = db.get_payment_limits_registry_by_id(int(registry_id))
            if not meta or not meta.get('has_file'):
                return jsonify({'error': 'Исходный файл реестра не найден в базе. '
                                         'Загрузите файл и сохраните вариант заново'}), 404
            file_name = meta.get('file_name') or f'Реестр_{registry_id}.xlsx'
            if not sheet_name:
                sheet_name = (meta.get('sheet_name') or '').strip() or None
            # Какая стадия фактически выгружается
            used_stage = stage if stage in (db.PLP_STAGE_DRAFT, db.PLP_STAGE_FINAL) \
                else (db.PLP_STAGE_FINAL if meta.get('final_saved_at') else db.PLP_STAGE_DRAFT)
            if remove_rows is None:
                # Берём сохранённый набор нужной стадии
                remove_rows = (meta.get('final_removed_rows') or []) if used_stage == db.PLP_STAGE_FINAL \
                    else (meta.get('draft_removed_rows') or [])
            _, raw = db.get_payment_limits_registry_file(int(registry_id))
        else:
            return jsonify({'error': 'Не передан исходный файл'}), 400

        if not raw:
            return jsonify({'error': 'Пустой файл'}), 400
        if remove_rows is None:
            remove_rows = []

        try:
            result, info = (None, None)
            if split_by_group:
                # Разбивка по вкладкам-категориям (колонка «Группа») со сводными таблицами
                result, info = _pl_split_sheets_by_group(raw, remove_rows, sheet_name, group_sheets,
                                                         cell_edits)
                used_split = result is not None
            if result is None:
                result, info = _pl_delete_rows_from_xlsx(raw, remove_rows, sheet_name,
                                                         cell_edits=cell_edits)
                if split_by_group:
                    note = 'категории в колонке «Группа» не найдены — файл выгружен одним листом'
                    info['warning'] = (info.get('warning') + '; ' if info.get('warning') else '') + note
        except zipfile.BadZipFile:
            return jsonify({'error': 'Экспорт доступен только для файлов .xlsx '
                                     '(файл не является книгой Excel 2007+). Сохраните реестр как .xlsx'}), 400
        except ValueError as e:
            return jsonify({'error': str(e)}), 400

        tabs = info.get('tabs') or []
        logger.info(f"✅ Экспорт реестра лимитов: файл «{file_name}»"
                    + (f" (реестр #{registry_id}, стадия «{stage or 'авто'}»)" if registry_id else '')
                    + f", снято строк {info['removed']}, осталось строк {info['rows_left']}, лист {info['sheet']}"
                    + (f", правок в ячейках {sum(len(c) for c in cell_edits.values())}" if cell_edits else '')
                    + (f", вкладок {len(tabs)}: "
                       + ', '.join(f"{tab['name']} ({tab['rows']})" for tab in tabs) if tabs else '')
                    + (f", предупреждение: {info['warning']}" if info.get('warning') else ''))

        response = send_file(
            result,
            as_attachment=True,
            download_name=(f"Заявки_лимиты_по_категориям_{datetime.now().strftime('%d.%m.%Y')}.xlsx"
                           if used_split else
                           f"Заявки_лимиты_{datetime.now().strftime('%d.%m.%Y')}.xlsx"),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response.headers['X-Export-Removed'] = str(info['removed'])
        if tabs:
            response.headers['X-Export-Tabs'] = str(len(tabs))
        if info.get('warning'):
            response.headers['X-Export-Warning'] = urllib.parse.quote(info['warning'])

        # Выгрузили финальный вариант — снимаем уведомление «можно выгружать»
        exported_final_at = None
        if registry_id and used_stage == db.PLP_STAGE_FINAL:
            marked = db.mark_payment_limits_registry_exported(int(registry_id))
            if marked.get('success'):
                exported_final_at = marked.get('final_exported_at')
                response.headers['X-Export-Final-At'] = str(exported_final_at or '')
        return response
    except Exception as e:
        logger.error(f"Ошибка API payment-limits/export: {e}")
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    logger.info("Сервер запущен на http://0.0.0.0:5000")
    try:
        app.run(debug=False, port=5000, host='0.0.0.0')
    finally:
        db.close_all_connections()
        logger.info("Сервер остановлен")