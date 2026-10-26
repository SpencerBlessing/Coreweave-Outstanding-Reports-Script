"""Persistent daily CoreWeave outstanding-case snapshots and comparisons."""
from __future__ import annotations

import csv
import json
import sqlite3
from datetime import date, timedelta
from pathlib import Path

import pandas as pd


def _value(value):
    if pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, date)):
        return value.isoformat()
    return str(value)


def _cases_on(conn, day):
    return {
        row[0]: row[1]
        for row in conn.execute(
            'SELECT case_number, category FROM case_snapshots WHERE snapshot_date = ?',
            (day,),
        )
    }


def _status_rows(conn, day):
    """Return status and category for cases present on a given date."""
    return {
        case_number: (json.loads(payload).get('Status'), category)
        for case_number, category, payload in conn.execute(
            'SELECT case_number, category, case_data FROM case_snapshots WHERE snapshot_date = ?',
            (day,),
        )
    }


def _clean_status(value):
    return str(value).strip() if value is not None else ''


def log_report_history(cases: pd.DataFrame, categories: pd.Series, history_folder: Path,
                       status=lambda message: None, snapshot_date=None):
    """Record a daily case-level snapshot and return day/week comparisons.

    One snapshot per calendar date. Re-running on the same date replaces that date's
    data, so running the report twice does not inflate history.
    """
    day = str(snapshot_date or date.today())
    history_folder = Path(history_folder)
    history_folder.mkdir(parents=True, exist_ok=True)
    db_path = history_folder / 'outstanding_history.sqlite3'
    csv_path = history_folder / 'daily_summary.csv'
    status_csv_path = history_folder / 'status_changes.csv'

    if 'Case Number' not in cases.columns:
        raise ValueError('Case Number is required for history logging')
    if cases['Case Number'].isna().any() or cases['Case Number'].duplicated().any():
        raise ValueError('Case numbers must be populated and unique for history logging')

    categories = categories.reindex(cases.index)
    records = []
    for idx, row in cases.iterrows():
        payload = {str(k): _value(v) for k, v in row.items()}
        records.append((day, str(row['Case Number']), str(categories.loc[idx]),
                        json.dumps(payload, ensure_ascii=False)))

    with sqlite3.connect(db_path) as conn:
        conn.execute('''CREATE TABLE IF NOT EXISTS case_snapshots (
            snapshot_date TEXT NOT NULL,
            case_number TEXT NOT NULL,
            category TEXT NOT NULL,
            case_data TEXT NOT NULL,
            PRIMARY KEY (snapshot_date, case_number)
        )''')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_case_history ON case_snapshots(case_number)')
        conn.execute('DELETE FROM case_snapshots WHERE snapshot_date = ?', (day,))
        conn.executemany('''INSERT INTO case_snapshots
            (snapshot_date, case_number, category, case_data) VALUES (?, ?, ?, ?)''', records)

        current = _cases_on(conn, day)
        results = []
        for label, offset in [('Day over day', 1), ('Week over week', 7)]:
            baseline_day = (date.fromisoformat(day) - timedelta(days=offset)).isoformat()
            previous = _cases_on(conn, baseline_day)
            if not previous:
                results.append((label, baseline_day, None, None, None, None))
                continue
            current_ids, previous_ids = set(current), set(previous)
            results.append((label, baseline_day, len(previous),
                            len(current_ids - previous_ids),
                            len(previous_ids - current_ids),
                            len(current) - len(previous)))

        # Rebuild the human-readable daily log from the database so reruns remain idempotent.
        days = [row[0] for row in conn.execute(
            'SELECT DISTINCT snapshot_date FROM case_snapshots ORDER BY snapshot_date')]
        # Reconstruct observed status transitions across consecutive available snapshots.
        # Rebuilding prevents duplicates and handles same-day reruns correctly.
        status_change_count = 0
        with status_csv_path.open('w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(['Detected Date', 'Previous Snapshot Date', 'Case Number',
                             'Previous Status', 'New Status', 'Previous Category', 'New Category'])
            for previous_day, current_day in zip(days, days[1:]):
                previous_rows = _status_rows(conn, previous_day)
                current_rows = _status_rows(conn, current_day)
                for case_number in sorted(previous_rows.keys() & current_rows.keys()):
                    old_status, old_category = previous_rows[case_number]
                    new_status, new_category = current_rows[case_number]
                    old_status = _clean_status(old_status)
                    new_status = _clean_status(new_status)
                    if old_status != new_status:
                        writer.writerow([current_day, previous_day, case_number,
                                         old_status, new_status, old_category, new_category])
                        if current_day == day:
                            status_change_count += 1

        with csv_path.open('w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(['Date', 'Total Outstanding', 'Regular', 'Self-Maintenance',
                             'H100 Thermal', 'Previous Day Total', 'New Since Previous Day',
                             'Removed Since Previous Day', 'Net Daily Change',
                             'Previous Week Total', 'New Since Previous Week',
                             'Removed Since Previous Week', 'Net Weekly Change'])
            for logged_day in days:
                snapshot = _cases_on(conn, logged_day)
                prior_day = _cases_on(conn, (date.fromisoformat(logged_day) - timedelta(days=1)).isoformat())
                prior_week = _cases_on(conn, (date.fromisoformat(logged_day) - timedelta(days=7)).isoformat())
                def changes(prior):
                    if not prior:
                        return ('', '', '', '')
                    current_set, prior_set = set(snapshot), set(prior)
                    return (len(prior), len(current_set - prior_set),
                            len(prior_set - current_set), len(snapshot) - len(prior))
                writer.writerow([logged_day, len(snapshot),
                                 sum(v == 'CoreWeave Outstanding Cases' for v in snapshot.values()),
                                 sum(v == 'Self-Maintenance' for v in snapshot.values()),
                                 sum(v == 'H100 Thermal Issues' for v in snapshot.values()),
                                 *changes(prior_day), *changes(prior_week)])

    status(f'History saved: {len(records)} cases on {day} ({db_path.name})')
    status(f'Status changes detected for {day}: {status_change_count} ({status_csv_path.name})')
    for label, baseline_day, old_total, added, removed, net in results:
        if old_total is None:
            status(f'{label}: no snapshot for {baseline_day} yet')
        else:
            status(f'{label} vs {baseline_day}: {added} new, {removed} removed, '
                   f'net {net:+d} ({old_total} -> {len(records)})')
    return db_path, csv_path
