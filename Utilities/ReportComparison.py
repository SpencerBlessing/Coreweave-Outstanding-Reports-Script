"""Compare CoreWeave case exports against the immediately preceding saved report run."""
from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


HISTORY_SCHEMA = '''
CREATE TABLE IF NOT EXISTS report_runs (
    run_id INTEGER PRIMARY KEY AUTOINCREMENT,
    captured_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS report_run_cases (
    run_id INTEGER NOT NULL,
    case_number TEXT NOT NULL,
    category TEXT NOT NULL,
    case_data TEXT NOT NULL,
    PRIMARY KEY (run_id, case_number),
    FOREIGN KEY (run_id) REFERENCES report_runs(run_id)
);
CREATE INDEX IF NOT EXISTS idx_run_case ON report_run_cases(case_number);
'''


def _text(value):
    if value is None:
        return ''
    try:
        if pd.isna(value):
            return ''
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def _serialized(value):
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.isoformat()
    return _text(value)


def _snapshot(cases, categories):
    if 'Case Number' not in cases:
        raise ValueError('Missing Case Number')
    if cases['Case Number'].isna().any() or cases['Case Number'].duplicated().any():
        raise ValueError('Report comparison requires populated, unique case numbers')
    categories = categories.reindex(cases.index)
    if categories.isna().any():
        raise ValueError('Missing worksheet category for some cases')
    return {
        _text(row['Case Number']): (
            _text(categories.loc[idx]),
            {str(key): _serialized(value) for key, value in row.items()},
        )
        for idx, row in cases.iterrows()
    }


def _previous_run(conn):
    previous = conn.execute('SELECT run_id, captured_at FROM report_runs ORDER BY run_id DESC LIMIT 1').fetchone()
    if previous is None:
        return None, {}
    rows = conn.execute(
        'SELECT case_number, category, case_data FROM report_run_cases WHERE run_id = ?',
        (previous[0],),
    )
    return previous, {number: (category, json.loads(payload)) for number, category, payload in rows}


def _age_bucket(value):
    import re
    match = re.search(r'\d+(?:\.\d+)?', _text(value))
    if not match:
        return 'Unknown'
    age = float(match.group())
    if age <= 9:
        return 'A: 0-9 days'
    if age <= 15:
        return 'B: 10-15 days'
    if age <= 30:
        return 'C: 16-30 days'
    return 'D: 31+ days'


def prepare_comparison(cases, categories, history_folder):
    """Read the last run without writing the current one; return comparison sections."""
    db_path = Path(history_folder) / 'outstanding_history.sqlite3'
    current = _snapshot(cases, categories)
    if db_path.exists():
        with sqlite3.connect(db_path) as conn:
            conn.executescript(HISTORY_SCHEMA)
            previous_run, previous = _previous_run(conn)
    else:
        previous_run, previous = None, {}

    now_ids, old_ids = set(current), set(previous)
    added = now_ids - old_ids if previous_run else set()
    removed = old_ids - now_ids if previous_run else set()
    common = now_ids & old_ids

    def details(ids, source):
        preferred = ['Case Number', 'Status', 'Case Age', 'Site', 'Asset Location',
                     'Product Name', 'Faulty Part Number', 'Faulty Part Description', 'RMA Number']
        if not ids:
            return pd.DataFrame(columns=preferred + ['Worksheet'])
        rows = []
        for number in sorted(ids):
            category, payload = source[number]
            rows.append({**{column: payload.get(column, '') for column in preferred},
                         'Case Number': number, 'Worksheet': category})
        return pd.DataFrame(rows)

    changed = []
    moved = []
    for number in sorted(common):
        old_category, old_data = previous[number]
        new_category, new_data = current[number]
        if _text(old_data.get('Status')) != _text(new_data.get('Status')):
            changed.append({'Case Number': number,
                            'Previous Status': old_data.get('Status', ''),
                            'Current Status': new_data.get('Status', ''),
                            'Site': new_data.get('Site', ''),
                            'Previous Worksheet': old_category,
                            'Current Worksheet': new_category})
        if old_category != new_category:
            moved.append({'Case Number': number, 'Previous Worksheet': old_category,
                          'Current Worksheet': new_category,
                          'Status': new_data.get('Status', ''),
                          'Site': new_data.get('Site', '')})

    status_changes = pd.DataFrame(changed, columns=[
        'Case Number', 'Previous Status', 'Current Status', 'Site',
        'Previous Worksheet', 'Current Worksheet'])
    movements = pd.DataFrame(moved, columns=[
        'Case Number', 'Previous Worksheet', 'Current Worksheet', 'Status', 'Site'])
    summary = pd.DataFrame([
        ('Previous report', previous_run[1] if previous_run else 'No previous snapshot'),
        ('Current report', datetime.now().isoformat(timespec='seconds')),
        ('Previous outstanding', len(old_ids) if previous_run else 'N/A'),
        ('Current outstanding', len(now_ids)),
        ('Newly listed cases', len(added) if previous_run else 'N/A'),
        ('Removed from outstanding', len(removed) if previous_run else 'N/A'),
        ('Status changes', len(status_changes) if previous_run else 'N/A'),
        ('Worksheet movements', len(movements) if previous_run else 'N/A'),
        ('Net backlog change', len(now_ids) - len(old_ids) if previous_run else 'N/A'),
    ], columns=['Metric', 'Value'])

    ages = pd.DataFrame([
        {'Age Bucket': _age_bucket(payload.get('Case Age')),
         'Status': _text(payload.get('Status')) or 'Unknown'}
        for _, payload in current.values()
    ])
    if ages.empty:
        age_status = pd.DataFrame(columns=['Age Bucket', 'Total'])
    else:
        age_status = pd.crosstab(ages['Age Bucket'], ages['Status'])
        age_status = age_status.reindex(['A: 0-9 days', 'B: 10-15 days',
                                         'C: 16-30 days', 'D: 31+ days', 'Unknown'], fill_value=0)
        age_status['Total'] = age_status.sum(axis=1)
        age_status = age_status.reset_index()
    return [
        ('REPORT SUMMARY', summary),
        ('NEWLY LISTED CASES', details(added, current)),
        ('REMOVED FROM OUTSTANDING', details(removed, previous)),
        ('STATUS CHANGES', status_changes),
        ('WORKSHEET MOVEMENTS', movements),
        ('CURRENT CASES BY AGE AND STATUS', age_status),
    ]


def write_comparison(writer, sections):
    """Write one formatted Comparison worksheet with stacked tables."""
    ws = writer.book.create_sheet('Comparison')
    writer.sheets['Comparison'] = ws
    heading_fill = PatternFill('solid', fgColor='17365D')
    header_fill = PatternFill('solid', fgColor='DCE8F3')
    row = 1
    for title, frame in sections:
        ws.cell(row, 1, title)
        ws.cell(row, 1).fill = heading_fill
        ws.cell(row, 1).font = Font(color='FFFFFF', bold=True, size=12)
        ws.merge_cells(start_row=row, start_column=1, end_row=row,
                       end_column=max(2, len(frame.columns)))
        row += 1
        for col, name in enumerate(frame.columns, 1):
            cell = ws.cell(row, col, name)
            cell.fill = header_fill
            cell.font = Font(bold=True)
        row += 1
        for values in frame.itertuples(index=False, name=None):
            for col, value in enumerate(values, 1):
                ws.cell(row, col, None if pd.isna(value) else value)
            row += 1
        row += 3
    ws.freeze_panes = 'A3'
    ws.sheet_view.showGridLines = False
    for col in range(1, max((len(frame.columns) for _, frame in sections), default=1) + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 24 if col < 3 else 22
    ws.column_dimensions['A'].width = 34
    ws.column_dimensions['E'].width = 40
    return ws


def save_report_run(cases, categories, history_folder):
    """Persist every generation as a separate run after the workbook was written."""
    folder = Path(history_folder)
    folder.mkdir(parents=True, exist_ok=True)
    snapshot = _snapshot(cases, categories)
    with sqlite3.connect(folder / 'outstanding_history.sqlite3') as conn:
        conn.executescript(HISTORY_SCHEMA)
        with conn:
            cursor = conn.execute('INSERT INTO report_runs (captured_at) VALUES (?)',
                                  (datetime.now().isoformat(timespec='microseconds'),))
            run_id = cursor.lastrowid
            conn.executemany('''INSERT INTO report_run_cases
                (run_id, case_number, category, case_data) VALUES (?, ?, ?, ?)''',
                [(run_id, number, category, json.dumps(data, ensure_ascii=False))
                 for number, (category, data) in snapshot.items()])
    return run_id
