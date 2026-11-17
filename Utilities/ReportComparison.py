"""Compare CoreWeave cases against the latest earlier finalized snapshot."""

from __future__ import annotations


import json

import sqlite3

from datetime import datetime, date

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


def _as_date(value):
    """Interpret a report date or optional date-cutoff consistently."""
    if value is None or value == '':
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value).strip())


def _filter_snapshot_by_creation_date(snapshot, since):
    """Apply the same Salesforce Case Creation Date cutoff to saved history.

    Cases without a parseable creation date are excluded when a cutoff exists,
    matching the current report's pd.to_datetime(..., errors='coerce') filter.
    This is read-only; the historical database is never modified.
    """
    if since is None:
        return snapshot

    filtered = {}
    for number, (category, payload) in snapshot.items():
        created = pd.to_datetime(payload.get('Case Creation Date'), errors='coerce')
        if pd.notna(created) and created.date() >= since:
            filtered[number] = (category, payload)
    return filtered


def _previous_run(conn, comparison_date=None, since=None):
    """Read the latest prior finalized snapshot within the selected date range.

    Skipped calendar days do not prevent comparison. Today's snapshot is never
    used as a baseline. When a start date is supplied, do not select a snapshot
    captured before that date. The cases within the snapshot are independently
    filtered by their Case Creation Date in prepare_comparison().

    Prefer official case_snapshots history. Legacy report_runs are used only
    when the modern snapshot table does not exist.
    """
    day = _as_date(comparison_date) or date.today()
    cutoff = _as_date(since)
    report_day = day.isoformat()
    tables = {row[0] for row in conn.execute(
        "SELECT name FROM sqlite_master WHERE type = 'table'")}

    if 'case_snapshots' in tables:
        query = ('SELECT MAX(snapshot_date) FROM case_snapshots '
                 'WHERE snapshot_date < ?')
        params = [report_day]
        if cutoff is not None:
            query += ' AND snapshot_date >= ?'
            params.append(cutoff.isoformat())
        previous_day = conn.execute(query, params).fetchone()[0]
        if previous_day is None:
            return None, {}
        rows = conn.execute(
            'SELECT case_number, category, case_data FROM case_snapshots '
            'WHERE snapshot_date = ?', (previous_day,))
        return (previous_day, previous_day), {
            number: (category, json.loads(payload))
            for number, category, payload in rows
        }

    # Compatibility with legacy run history; exclude current/future reports.
    if {'report_runs', 'report_run_cases'}.issubset(tables):
        query = ('SELECT run_id, captured_at FROM report_runs '
                 'WHERE substr(captured_at, 1, 10) < ?')
        params = [report_day]
        if cutoff is not None:
            query += ' AND substr(captured_at, 1, 10) >= ?'
            params.append(cutoff.isoformat())
        query += ' ORDER BY captured_at DESC, run_id DESC LIMIT 1'
        previous = conn.execute(query, params).fetchone()
        if previous is not None:
            previous_day = previous[1][:10]
            rows = conn.execute(
                'SELECT case_number, category, case_data '
                'FROM report_run_cases WHERE run_id = ?', (previous[0],))
            return (previous[0], previous_day), {
                number: (category, json.loads(payload))
                for number, category, payload in rows
            }
    return None, {}


def _age_bucket(value):
    """Respect Salesforce's A/B/C/D labels before parsing numeric day counts."""
    import re
    text = _text(value)
    if not text:
        return 'Unknown'
    buckets = {
        'A': 'A: 0-9 days', 'B': 'B: 10-15 days',
        'C': 'C: 16-30 days', 'D': 'D: 31+ days',
    }
    label = re.match(r'^([ABCD])(?:\s|[·:.-]|$)', text, re.IGNORECASE)
    if label:
        return buckets[label.group(1).upper()]
    if re.search(r'(?:>|≥)\s*30\b|31\s*\+', text, re.IGNORECASE):
        return buckets['D']
    match = re.search(r'\d+(?:\.\d+)?', text)
    if not match:
        return 'Unknown'
    age = float(match.group())
    if age <= 9:
        return buckets['A']
    if age <= 15:
        return buckets['B']
    if age <= 30:
        return buckets['C']
    return buckets['D']


THERMAL_CATEGORY = 'H100 Thermal Issues'


def _thermal_sections(current, previous, has_previous):

    """Summarize H100 thermal only; category membership is the definition.


    A case moving into/out of the H100 worksheet is counted as entering/leaving

    the thermal backlog, not necessarily as newly opened or resolved.

    """

    current_thermal = {

        number: payload for number, (category, payload) in current.items()

        if category == THERMAL_CATEGORY

    }

    previous_thermal = {

        number: payload for number, (category, payload) in previous.items()

        if category == THERMAL_CATEGORY

    }

    new_ids = set(current_thermal) - set(previous_thermal)

    removed_ids = set(previous_thermal) - set(current_thermal)

    previous_total = len(previous_thermal) if has_previous else 'N/A'

    net = len(current_thermal) - len(previous_thermal) if has_previous else 'N/A'

    aged = sum(_age_bucket(p.get('Case Age')) == 'D: 31+ days'

               for p in current_thermal.values())

    share = f'{100 * len(current_thermal) / len(current):.1f}%' if current else '0.0%'

    summary = pd.DataFrame([

        ('Current H100 thermal backlog', len(current_thermal)),

        ('Previous H100 thermal backlog', previous_total),

        ('Entered H100 thermal backlog', len(new_ids) if has_previous else 'N/A'),

        ('Left H100 thermal backlog', len(removed_ids) if has_previous else 'N/A'),

        ('Net H100 thermal change', net),

        ('Share of all outstanding cases', share),

        ('H100 thermal aged 31+ days', aged),

    ], columns=['Metric', 'Value'])


    def counts(records, field):

        result = {}

        for payload in records.values():

            if field == 'Age Bucket':

                value = _age_bucket(payload.get('Case Age'))

            else:

                value = _text(payload.get(field)) or 'Unknown'

            result[value] = result.get(value, 0) + 1

        return result


    def compare(field, limit=None, include_aged=False):

        now_counts = counts(current_thermal, field)

        old_counts = counts(previous_thermal, field) if has_previous else {}

        keys = sorted(set(now_counts) | set(old_counts),

                      key=lambda value: (-now_counts.get(value, 0),

                                         -old_counts.get(value, 0), value))

        if field == 'Age Bucket':

            keys = [key for key in (

                'A: 0-9 days', 'B: 10-15 days', 'C: 16-30 days',

                'D: 31+ days', 'Unknown') if key in set(keys)]

        if limit is not None and len(keys) > limit:

            shown = keys[:limit]

            remaining = keys[limit:]

        else:

            shown, remaining = keys, []

        age_by_key = {}

        if include_aged:

            for payload in current_thermal.values():

                key = _text(payload.get(field)) or 'Unknown'

                if _age_bucket(payload.get('Case Age')) == 'D: 31+ days':

                    age_by_key[key] = age_by_key.get(key, 0) + 1

        def make_row(key, members):

            now = sum(now_counts.get(member, 0) for member in members)

            old = sum(old_counts.get(member, 0) for member in members)

            row = {field: key, 'Current': now,

                   'Previous': old if has_previous else 'N/A',

                   'Net Change': now - old if has_previous else 'N/A'}

            if include_aged:

                row['Aged 31+ Days'] = sum(age_by_key.get(member, 0)

                                           for member in members)

            return row

        rows = [make_row(key, [key]) for key in shown]

        if remaining:

            rows.append(make_row('Other', remaining))

        columns = [field, 'Current', 'Previous', 'Net Change']

        if include_aged:

            columns.append('Aged 31+ Days')

        return pd.DataFrame(rows, columns=columns)


    return [

        ('H100 THERMAL SUMMARY', summary),

        ('H100 THERMAL BY SITE', compare('Site', limit=10, include_aged=True)),

        ('H100 THERMAL ROOT CAUSES', compare('Root Cause: Root Cause Name', limit=10)),

        ('H100 THERMAL BY STATUS', compare('Status')),

        ('H100 THERMAL BY AGE', compare('Age Bucket')),

    ]



def _display_site_names_in_old_snapshot(previous, history_folder):
    """Normalize legacy Site codes for display, without changing stored history.

    Previous snapshots may contain `site` or LoCode in the Site column,
    whereas newer reports display `site_name`. For consistent by-Site deltas,
    resolve *each historical address* against the same authoritative
    CoreWeave_Map.json, and replace an old code only when the verified record
    agrees with that code. Unmatched/conflicting addresses are left unchanged.
    """
    if not previous:
        return previous
    map_path = Path(history_folder).resolve().parents[2] / 'Address Cache' / 'CoreWeave_Map.json'
    if not map_path.is_file():
        # The main report will already reject a missing map; keep this helper
        # usable on standalone historical comparison fixtures, too.
        return previous

    from Utilities.AddressStandardizer import SiteRegistry
    registry = SiteRegistry.from_map(map_path)
    adjusted = {}
    for case_number, (category, payload) in previous.items():
        old_label = _text(payload.get('Site'))
        address = _text(payload.get('Asset Location'))
        if not old_label or not address or old_label == 'Shipping':
            adjusted[case_number] = (category, payload)
            continue
        match = registry.resolve(address=address)
        site = match.site if match.status == 'matched' else None
        if site and site.site_name and old_label.casefold() in {
            _text(site.site_code).casefold(), _text(site.locode).casefold()
        }:
            payload = dict(payload)
            payload['Site'] = site.site_name
        adjusted[case_number] = (category, payload)
    return adjusted



def prepare_comparison(cases, categories, history_folder, comparison_date=None, since=None):

    """Compare the same case-date cohort in current and historical snapshots.

    The optional `since` cutoff limits both current cases and prior snapshot
    cases by Case Creation Date, and restricts the chosen historical snapshot
    date. No snapshot or map data is written or changed here.
    """
    comparison_day = _as_date(comparison_date) or date.today()
    cutoff = _as_date(since)

    db_path = Path(history_folder) / 'outstanding_history.sqlite3'

    current = _filter_snapshot_by_creation_date(_snapshot(cases, categories), cutoff)

    if db_path.exists():
        with sqlite3.connect(db_path) as conn:
            previous_run, previous = _previous_run(conn, comparison_day, cutoff)
    else:
        previous_run, previous = None, {}

    previous = _filter_snapshot_by_creation_date(previous, cutoff)

    # Bridge legacy code-based snapshots to site_name for by-Site comparisons.
    # No database rows or map records are updated.
    previous = _display_site_names_in_old_snapshot(previous, history_folder)

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

        ('Comparison source', previous_run[1] if previous_run else
         f'No saved snapshot before {comparison_day.isoformat()}'),
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

        age_status = age_status.reindex(
            ['A: 0-9 days', 'B: 10-15 days',
            'C: 16-30 days', 'D: 31+ days'],
            fill_value=0
        )

        age_status['Total'] = age_status.sum(axis=1)

        age_status = age_status.reset_index()

    sections = [('REPORT SUMMARY', summary)]

    # Thermal analytics immediately follow the general summary, so they remain

    # visible even when the new/removed case lists contain thousands of rows.

    sections.extend(_thermal_sections(current, previous, previous_run is not None))

    sections.extend([

        ('NEWLY LISTED CASES', details(added, current)),

        ('REMOVED FROM OUTSTANDING', details(removed, previous)),

        ('STATUS CHANGES', status_changes),

        ('WORKSHEET MOVEMENTS', movements),

        ('CURRENT CASES BY AGE AND STATUS', age_status),

    ])

    return sections


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