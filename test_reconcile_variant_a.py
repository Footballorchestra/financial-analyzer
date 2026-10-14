# -*- coding: utf-8 -*-
"""Регрессионный тест серверного движка сверки ДЗ (вариант A).

Запуск:  venv/bin/python test_reconcile_variant_a.py

Сценарий:
  1. Берёт файлы из Problema (исходник, поступления, СИ УАТ) как за 02.09.2026;
  2. вызывает POST /api/reconcile (режим email) с той же расчётной датой,
     сводными значениями и данными предыдущего дня из эталона;
  3. сравнивает контрольные точки с эталоном ДЗ_..._обработка.xlsx:
     верхняя таблица, строка «Итого», просрочка по всем филиалам (ДТ …),
     динамика по подразделениям и сводные блоки.

Работает без БД: функции БД подменяются заглушками.
"""
import base64
import json
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

import openpyxl  # noqa: E402


ETALON_PREV_DATE = '2026-09-02'
ETALON_SUMMARY = {
    'summaryDtLegal': '23012946.33',
    'summarySiuatTotalDebt': '390387353.44',
    'summarySiuatTotalOverdue': '150015075.11',
}


def _stub_db(server):
    server.db.get_target_contractors = lambda: []
    server.db.get_previous_day_data = lambda d: {}
    server.db.get_last_available_swipe_date = lambda d: None
    server.db.save_swipe_data = lambda *a, **k: {'success': True, 'count': 0}


def _etalon_previous(etalon_path):
    """Данные предыдущего дня из колонки C листа 'Сводные таблицы' эталона."""
    prev = {}
    wb = openpyxl.load_workbook(etalon_path, data_only=True, read_only=True)
    ws = wb['Сводные таблицы']
    for row in ws.iter_rows(min_row=4, values_only=True):
        if row and row[0] and str(row[0]).strip().startswith('ДТ ') and row[2] is not None:
            prev[str(row[0]).strip()] = row[2]
    wb.close()
    return prev


def run():
    import server
    _stub_db(server)
    app = server.app
    app.config['TESTING'] = True
    client = app.test_client()

    src = os.path.join(BASE, 'Problema', 'ДЗ_Транспорт_02.09.2026 исходник.xlsx')
    rec = os.path.join(BASE, 'Problema', 'Поступления 02.09.2026.xlsx')
    siu = os.path.join(BASE, 'Problema', 'ДЗ Транспорт 02.09.2026 СИ УАТ.xlsx')
    etalon = os.path.join(BASE, 'Problema', 'ДЗ_Транспорт_02.09.2026 обработка.xlsx')
    for p in (src, rec, siu, etalon):
        if not os.path.exists(p):
            raise SystemExit('Не найден файл: ' + p)

    prev = _etalon_previous(etalon)
    with open(src, 'rb') as f1, open(rec, 'rb') as f2, open(siu, 'rb') as f3:
        form = {
            'asOfDate': '2026-09-02',
            'targetContractors': json.dumps(['ВАНКОРНЕФТЬ АО', 'РН-Ванкор ООО']),
            'mode': 'email',
            'previousDate': ETALON_PREV_DATE,
            'previousDayData': json.dumps(prev),
            **ETALON_SUMMARY,
            'file': (f1, 'registry.xlsx'),
            'receiptsFile': (f2, 'receipts.xlsx'),
            'siUatFile': (f3, 'siuat.xlsx'),
        }
        resp = client.post('/api/reconcile', data=form)

    payload = resp.get_json()
    assert payload and payload.get('success'), 'Движок вернул ошибку: ' + str(payload)[:500]
    if not payload.get('file') or not payload.get('emailData'):
        raise AssertionError('В ответе нет file/emailData')

    engine_path = os.path.join('/tmp', 'reconcile_vA_engine.xlsx')
    with open(engine_path, 'wb') as out:
        out.write(base64.b64decode(payload['file']['data']))

    errors = _compare(engine_path, etalon)
    if errors:
        raise AssertionError('\n'.join(errors))
    print('OK — контрольные точки движка совпадают с эталоном')
    print('  stats:', {k: payload['stats'].get(k) for k in ('totalDocuments', 'foundDocuments', 'updatedDocuments')})
    print('  emailData keys:', sorted(payload['emailData'].keys()))


def _num(v):
    return float(v or 0)


def _compare(engine_path, etalon_path):
    errs = []
    E = openpyxl.load_workbook(engine_path, data_only=True)
    T = openpyxl.load_workbook(etalon_path, data_only=True)

    we, wt = E['Свод ДЗ'], T['Свод ДЗ']
    for r in list(range(2, 9)) + [2427]:
        if abs(_num(we.cell(row=r, column=5).value) - _num(wt.cell(row=r, column=5).value)) > 0.01:
            errs.append(f'Не совпала верхняя таблица/Итого: строка {r}, колонка E')
    for c in (12, 15, 18, 20, 21, 22, 23, 24, 25):
        if abs(_num(we.cell(row=2427, column=c).value) - _num(wt.cell(row=2427, column=c).value)) > 0.01:
            errs.append(f'Не совпала строка Итого: колонка {c}')

    def filial_map(ws):
        out = {}
        for i, row in enumerate(ws.iter_rows(min_row=14, values_only=True), start=14):
            if row and row[0] and str(row[0]).strip().startswith('ДТ '):
                out[str(row[0]).strip()] = row[14]
        return out

    me, mt = filial_map(we), filial_map(wt)
    if set(me) != set(mt):
        errs.append('Разные наборы филиалов в файлах')
    for name in me:
        if abs(_num(me[name]) - _num(mt.get(name))) > 0.01:
            errs.append(f'Не совпала просрочка филиала {name}: {me[name]} vs {mt.get(name)}')

    we2, wt2 = E['Сводные таблицы'], T['Сводные таблицы']
    for ws_a, ws_b, tag in [ (we2, wt2, 'динамика/сводки') ]:
        rows_a = {}
        for i, row in enumerate(ws_a.iter_rows(min_row=1, values_only=True), start=1):
            if row and row[0]:
                rows_a[str(row[0]).strip()] = row
        rows_b = {}
        for i, row in enumerate(ws_b.iter_rows(min_row=1, values_only=True), start=1):
            if row and row[0]:
                rows_b[str(row[0]).strip()] = row
        for name, row in rows_a.items():
            text = str(name)
            if text.startswith('ДТ '):
                other = rows_b.get(text)
                if other is None:
                    errs.append(f'Нет филиала {name} в эталоне ({tag})')
                    continue
                for col, colname in ((1, 'B'), (2, 'C'), (3, 'D')):
                    if abs(_num(row[col]) - _num(other[col])) > 0.01:
                        errs.append(f'Динамика {name} {colname}: {row[col]} vs {other[col]}')
            elif text in ('общая ДЗ', 'из них ПДЗ', 'Общая ДЗ', 'Из них ПДЗ'):
                other = rows_b.get(text)
                if other is None or abs(_num(row[1]) - _num(other[1])) > 0.01:
                    errs.append(f'Сводка {text}: {row[1]} vs {other and other[1]}')

    if 'Свод ДЗ СИ УАТ' not in E.sheetnames:
        errs.append('Нет листа Свод ДЗ СИ УАТ в результате')
    E.close()
    T.close()
    return errs


if __name__ == '__main__':
    run()
