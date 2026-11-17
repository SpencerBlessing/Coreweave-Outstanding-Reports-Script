"""Compact operational dashboard for the CoreWeave outstanding report.

Drop-in replacement; existing call remains:
    write_dashboard(writer, cases, categories, comparison, history_folder)
The output remains an ordinary Excel worksheet, included in finalized reports.
"""
from __future__ import annotations

import re
from datetime import datetime

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

NAVY = '17324D'
TEAL = '167A76'
WHITE = 'FFFFFF'
TEXT = '26394B'
MUTED = '65788A'
PALE = 'EAF0F5'
ALT = 'F5F8FA'

AGE_BUCKETS = ['0–9 days', '10–15 days', '16–30 days', '31+ days']


def _col(df, name):
    if name not in df:
        return pd.Series('', index=df.index, dtype='str')
    return df[name].fillna('').astype(str).str.strip()


def _age(value):
    if pd.isna(value) or not str(value).strip():
        return ''
    s = str(value).casefold().strip()
    for letter, bucket in zip('abcd', AGE_BUCKETS):
        if re.match(rf'^{letter}(?:\s*[·:\-]|\b)', s):
            return bucket
    if re.search(r'0\s*[-–]\s*9|(?:≤|<)\s*9', s):
        return AGE_BUCKETS[0]
    if re.search(r'10\s*[-–]\s*15', s):
        return AGE_BUCKETS[1]
    if re.search(r'16\s*[-–]\s*30', s):
        return AGE_BUCKETS[2]
    if re.search(r'31\s*\+|>\s*30', s):
        return AGE_BUCKETS[3]
    match = re.fullmatch(r'\d+(?:\.\d+)?', s)
    if match:
        days = float(s)
        return AGE_BUCKETS[0] if days <= 9 else AGE_BUCKETS[1] if days <= 15 else AGE_BUCKETS[2] if days <= 30 else AGE_BUCKETS[3]
    return ''


def _summary(comparison):
    section = dict(comparison or []).get('REPORT SUMMARY', pd.DataFrame())
    if not {'Metric', 'Value'}.issubset(section.columns):
        return {}
    return dict(zip(section['Metric'], section['Value']))


def _metric(summary, label):
    value = summary.get(label)
    if value is None or pd.isna(value) or str(value).upper() == 'N/A':
        return 'N/A'
    try:
        return f'{int(value):+,}' if label == 'Net backlog change' else f'{int(value):,}'
    except (TypeError, ValueError):
        return str(value)


def write_dashboard(writer, cases, categories, comparison, history_folder):
    """Write one readable, table-based operational dashboard; never mutate case data."""
    ws = writer.book.create_sheet('Dashboard', 0)
    writer.sheets['Dashboard'] = ws
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    ws.sheet_properties.tabColor = TEAL
    ws.freeze_panes = 'C13'
    widths = [3, 39, 18, 17, 17, 17, 17, 3, 25, 17, 17, 17, 3]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    for i in range(1, 95):
        ws.row_dimensions[i].height = 22
    ws.row_dimensions[2].height = 37

    def band(row, left, right, title):
        ws.merge_cells(start_row=row, start_column=left, end_row=row, end_column=right)
        for col in range(left, right + 1):
            ws.cell(row, col).fill = PatternFill('solid', fgColor=NAVY)
        cell = ws.cell(row, left)
        cell.value = title.upper()
        cell.font = Font(name='Aptos', size=10, bold=True, color=WHITE)
        cell.alignment = Alignment(vertical='center', indent=1)
        ws.row_dimensions[row].height = 27

    def table(row, left, headers, rows, max_rows=None):
        for offset, value in enumerate(headers):
            cell = ws.cell(row, left + offset, value)
            cell.fill = PatternFill('solid', fgColor=PALE)
            cell.font = Font(name='Aptos', size=10, bold=True, color=NAVY)
            cell.alignment = Alignment(vertical='center', wrap_text=True)
        ws.row_dimensions[row].height = 26
        limit = len(rows) if max_rows is None else max_rows
        for i in range(limit):
            values = rows[i] if i < len(rows) else []
            for j in range(len(headers)):
                value = values[j] if j < len(values) else None
                cell = ws.cell(row + i + 1, left + j, value)
                cell.font = Font(name='Aptos', size=10, color=TEXT)
                cell.alignment = Alignment(vertical='center', horizontal='right' if j > 0 else 'left')
                if i % 2:
                    cell.fill = PatternFill('solid', fgColor=ALT)
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    cell.number_format = '#,##0'
        return row + limit + 1

    summary = _summary(comparison)
    n = len(cases)
    cats = categories.reindex(cases.index).fillna('')
    ages = _col(cases, 'Case Age').map(_age)
    statuses = _col(cases, 'Status')
    sites = _col(cases, 'Site')
    regions = _col(cases, 'Region')
    rma = _col(cases, 'RMA Number')
    has_rma = rma.ne('')

    ws.merge_cells('B2:L2')
    for row in ws['B2:L2']:
        for c in row:
            c.fill = PatternFill('solid', fgColor=NAVY)
    header = ws['B2']
    header.value = 'COREWEAVE  |  OPERATIONS OVERVIEW'
    header.font = Font(name='Aptos Display', size=19, bold=True, color=WHITE)
    header.alignment = Alignment(vertical='center', indent=1)
    ws.merge_cells('B3:L3')
    ws['B3'] = f'Outstanding case snapshot  •  {datetime.now():%B %d, %Y %I:%M %p}  •  Changes relative to previous report run'
    ws['B3'].font = Font(name='Aptos', size=10, color=MUTED)

    # A small KPI strip, with no emphasis on a single age bucket.
    cards = [
        ('OUTSTANDING', f'{n:,}'),
        ('NEW SINCE PRIOR', _metric(summary, 'Newly listed cases')),
        ('LEFT OUTSTANDING', _metric(summary, 'Removed from outstanding')),
        ('NET CHANGE', _metric(summary, 'Net backlog change')),
        ('STATUS CHANGES', _metric(summary, 'Status changes')),
        ('RMA-LINKED', f'{int(has_rma.sum()):,}'),
        ('SELF-MAINTENANCE', f'{int(cats.eq("Self-Maintenance").sum()):,}'),
        ('H100 THERMAL', f'{int(cats.eq("H100 Thermal Issues").sum()):,}'),
    ]
    positions = [(2,3,5), (4,5,5), (6,6,5), (8,9,5), (10,12,5),
                 (2,3,8), (4,6,8), (8,10,8)]
    for (label, val), (l, r, row) in zip(cards, positions):
        ws.merge_cells(start_row=row, start_column=l, end_row=row, end_column=r)
        ws.merge_cells(start_row=row+1, start_column=l, end_row=row+1, end_column=r)
        for rr in [row, row+1]:
            for cc in range(l, r+1):
                ws.cell(rr, cc).fill = PatternFill('solid', fgColor=PALE)
        label_cell = ws.cell(row, l)
        value_cell = ws.cell(row+1, l)
        label_cell.value, value_cell.value = label, val
        label_cell.font = Font(name='Aptos', size=9, bold=True, color=MUTED)
        value_cell.font = Font(name='Aptos Display', size=17, bold=True, color=NAVY)
        label_cell.alignment = value_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[6].height = ws.row_dimensions[9].height = 34

    # Each table occupies its own row block so labels cannot visually collide.
    band(12, 2, 7, 'Case workload by status and age')
    top_statuses = statuses[statuses.ne('')].value_counts().head(10)
    status_rows = []
    for name, count in top_statuses.items():
        mask = statuses.eq(name)
        status_rows.append((name, int(count), *(int((mask & ages.eq(b)).sum()) for b in AGE_BUCKETS)))
    table(13, 2, ['Status', 'Total', '0–9 d', '10–15 d', '16–30 d', '31+ d'], status_rows, 10)

    band(27, 2, 6, 'Cases by site')
    top_sites = sites[sites.ne('')].value_counts().head(10)
    site_rows = []
    for name, count in top_sites.items():
        mask = sites.eq(name)
        site_rows.append((name, int(count), int((mask & ages.isin(AGE_BUCKETS[:2])).sum()),
                          int((mask & ages.eq('16–30 days')).sum()), int((mask & ages.eq('31+ days')).sum())))
    table(28, 2, ['Site', 'Cases', '0–15 d', '16–30 d', '31+ d'], site_rows, 10)

    band(42, 2, 6, 'Faulty parts and current workload')
    part_cases = {}
    # Unique case ID used so repeated part descriptions on a row do not inflate totals.
    for pos, value in enumerate(_col(cases, 'Faulty Part Number')):
        for part in set(p.strip() for p in re.split(r'\s*\|\s*|\s*,\s*', value) if p.strip()):
            part_cases.setdefault(part, set()).add(pos)
    part_rows = []
    for part, positions_set in sorted(part_cases.items(), key=lambda pair: (-len(pair[1]), pair[0]))[:10]:
        idx = list(positions_set)
        sub_age = ages.iloc[idx]
        sub_rma = has_rma.iloc[idx]
        part_rows.append((part, len(idx), int(sub_age.isin(AGE_BUCKETS[:2]).sum()),
                          int(sub_age.isin(AGE_BUCKETS[2:]).sum()), int(sub_rma.sum())))
    table(43, 2, ['Faulty Part', 'Cases', '0–15 d', '16+ d', 'With RMA'], part_rows, 10)

    band(57, 2, 6, 'Region workload')
    region_rows = []
    for region, count in regions[regions.ne('')].value_counts().head(8).items():
        mask = regions.eq(region)
        region_rows.append((region, int(count), int((mask & ages.isin(AGE_BUCKETS[:2])).sum()),
                            int((mask & ages.isin(AGE_BUCKETS[2:])).sum()), int((mask & has_rma).sum())))
    table(58, 2, ['Region', 'Cases', '0–15 d', '16+ d', 'With RMA'], region_rows, 8)

    band(70, 2, 4, 'Age distribution')
    age_rows = [(bucket, int(ages.eq(bucket).sum()),
                 f'{(ages.eq(bucket).sum()/n):.1%}' if n else '0.0%') for bucket in AGE_BUCKETS]
    table(71, 2, ['Age bucket', 'Cases', 'Share'], age_rows, 4)

    band(70, 9, 10, 'Since previous report')
    change_rows = [
        ('Newly listed cases', _metric(summary, 'Newly listed cases')),
        ('Left outstanding', _metric(summary, 'Removed from outstanding')),
        ('Status transitions', _metric(summary, 'Status changes')),
        ('Category movements', _metric(summary, 'Worksheet movements')),
        ('Net backlog change', _metric(summary, 'Net backlog change')),
    ]
    table(71, 9, ['Change', 'Cases'], change_rows, 5)

    band(80, 2, 3, 'Current worksheet totals')
    table(81, 2, ['Worksheet', 'Cases'], [
        ('CoreWeave Outstanding Cases', int(cats.eq('CoreWeave Outstanding Cases').sum())),
        ('Self-Maintenance', int(cats.eq('Self-Maintenance').sum())),
        ('H100 Thermal Issues', int(cats.eq('H100 Thermal Issues').sum())),
    ], 3)

    ws.merge_cells('B89:L89')
    ws['B89'] = ('Removed = no longer in the outstanding export, not necessarily resolved. '
                'RMA-linked = an RMA number is recorded. See Operational Insights for individual changes.')
    ws['B89'].font = Font(name='Aptos', size=10, italic=True, color=MUTED)
    ws.row_dimensions[89].height = 28
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 3
    ws.print_area = 'B2:L89'

    wb = writer.book

    # Rename Comparison
    if "Comparison" in wb.sheetnames:
        wb["Comparison"].title = "Operational Insights"

    # Reorder worksheets
    sheet_order = [
        "Dashboard",
        "Operational Insights",
        "CoreWeave Outstanding Cases",
        "H100 Thermal Issues",
        "Self-Maintenance",
    ]

    for index, name in enumerate(sheet_order):
        if name in wb.sheetnames:
            wb._sheets.insert(
                index,
                wb._sheets.pop(wb.sheetnames.index(name))
            )
    return ws
