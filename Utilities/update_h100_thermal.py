"""Update H100 thermal cases from the latest CoreWeave outstanding report.

Run from project root:
    python Utilities/update_h100_thermal.py

Requires: pip install openpyxl

- Existing CS tickets missing from outstanding cases move to Status Dashboard (Closed).
- Empty CS tickets match on server serial, with Jira disambiguation when possible.
- For EVERY matching open CS ticket, sync its latest outstanding-case status.
- Fill missing RMA numbers using the latest Salesforce H100 Thermal Issues export.
- Match RMA data by CS ticket AND server serial (and work order when possible).
- Preserve existing RMA values without overwriting or appending.
- Normalize recognized CS IDs to CS-######## and left-align output cells.
- All other thermal information, including NVIDIA GPU serial values, is untouched.
- Original input workbook is never overwritten.
"""
from collections import defaultdict
from copy import copy
from datetime import datetime
from pathlib import Path
import re

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
BASE = ROOT / 'Reports' / 'Coreweave_Outstanding'
THERMAL_FILE = BASE / 'Input' / 'Weekly H100 Thermal Report.xlsx'
THERMAL_ISSUES_DIR = BASE / 'Input'
OUTSTANDING_DIR = BASE / 'Output'
OUTPUT_FILE = OUTSTANDING_DIR / 'Weekly H100 Thermal Report - Updated.xlsx'


def normalized(value):
    return str(value).strip().upper() if value is not None else ''


def normalize_cs(value):
    """Standardize numeric and CS-prefixed case IDs to CS-########.

    Examples: 112490 -> CS-00112490, CS-112490 -> CS-00112490.
    Other identifiers are left untouched rather than guessed.
    """
    if value is None:
        return ''
    raw = str(value).strip().upper()
    match = re.fullmatch(r'(?:CS\s*[-#:]?\s*)?(\d{1,8})(?:\.0+)?', raw)
    if match:
        return f'CS-{int(match.group(1)):08d}'
    return raw


def left_align_used_cells(ws):
    """Left align cells while retaining their existing vertical/wrap settings."""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            alignment = copy(cell.alignment)
            alignment.horizontal = 'left'
            cell.alignment = alignment


def headers_for(sheet, row_number):
    return {normalized(c.value): c.column for c in sheet[row_number] if c.value is not None}


def required(headers, name):
    if name not in headers:
        raise ValueError(f'Missing required column: {name}')
    return headers[name]


def jira_codes(value):
    return set(re.findall(r'\b[A-Z]+-\d+\b', normalized(value)))


def rma_numbers(subject):
    """Extract actual RMA identifiers, not phrases such as 'RMA Request'.

    This outstanding export has no dedicated RMA number column. Only capture
    subjects with a literal RMA number, e.g. RMA-123456 or RMA #123456.
    """
    if subject is None:
        return set()
    return set(re.findall(r'\bRMA\s*[-#: ]\s*(\d{5,})\b', str(subject), re.I))


def merge_rmas(existing, additions):
    """Append new RMA IDs without altering existing text or duplicates."""
    if not additions:
        return existing
    current = '' if existing is None else str(existing).strip()
    known = set(re.findall(r'\d{5,}', current))
    new = [f'RMA-{number}' for number in sorted(additions) if number not in known]
    if not new:
        return existing
    return f'{current}, {", ".join(new)}' if current else ', '.join(new)


def style_header(ws, row, labels):
    for col, label in enumerate(labels, 1):
        cell = ws.cell(row, col, label)
        cell.fill = PatternFill('solid', fgColor='245A81')
        cell.font = Font(bold=True, color='FFFFFF')
        cell.alignment = Alignment(vertical='center')
    ws.row_dimensions[row].height = 24


def locate_outstanding():
    # Pick newest dated generated report (CoreWeave Outstanding Cases MMDDYYYY.xlsx).
    dated = []
    for path in OUTSTANDING_DIR.glob('CoreWeave Outstanding Cases *.xlsx'):
        match = re.fullmatch(r'CoreWeave Outstanding Cases (\d{8})\.xlsx', path.name, re.I)
        if not match:
            continue
        try:
            stamp = datetime.strptime(match.group(1), '%m%d%Y')
        except ValueError:
            continue
        dated.append((stamp, path))
    if not dated:
        raise FileNotFoundError(f'No dated outstanding report found in: {OUTSTANDING_DIR}')
    return max(dated, key=lambda p: (p[0], p[1].stat().st_mtime))[1]



def locate_thermal_issues():
    """Pick newest timestamped thermal-issues export in the Input folder."""
    matches = []
    for path in THERMAL_ISSUES_DIR.glob('H100 Thermal Issues Coreweave-*.xlsx'):
        match = re.fullmatch(
            r'H100 Thermal Issues Coreweave-(\d{4}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2})\.xlsx',
            path.name, re.I)
        if match:
            matches.append((match.group(1), path))
    return max(matches)[1] if matches else None


def normalize_wo(value):
    """Normalize WO numbers without touching source cells."""
    raw = normalized(value)
    match = re.fullmatch(r'(?:WO[- ]?)?(\d{1,8})(?:\.0+)?', raw)
    return f'WO-{int(match.group(1)):08d}' if match else raw


def parse_work_orders(value):
    if value is None:
        return set()
    return {normalize_wo(x) for x in re.split(r'[,;\n]+', str(value)) if normalized(x)}


def load_thermal_rmas(path):
    """Index validated RMA IDs by CS + server serial + WO.

    Salesforce exports leave the case cell blank on subsequent work-order
    line items. Carry forward the preceding CS and server serial so those
    legitimate line-item RMAs are not missed.
    """
    if path is None:
        return {}
    source = load_workbook(path, read_only=True, data_only=True)
    index = defaultdict(lambda: defaultdict(set))
    try:
        sheet = source['H100 Thermal Issues Coreweave']
        cols = headers_for(sheet, 12)
        case_col = next((col - 1 for label, col in cols.items() if label.startswith('CASE NUMBER')), None)
        if case_col is None:
            raise ValueError('Missing Case Number column in thermal issues export')
        serial_col = required(cols, 'SERIAL NUMBER') - 1
        wo_col = required(cols, 'WORK ORDER: WORK ORDER NUMBER') - 1
        rma_col = required(cols, 'WORK ORDER: RMA NUMBER') - 1
        active_case = active_serial = ''
        for row in sheet.iter_rows(min_row=13, values_only=True):
            if row[case_col] is not None and normalized(row[case_col]):
                active_case = normalize_cs(row[case_col])
                active_serial = ''  # do not leak a serial across cases
            if row[serial_col] is not None and normalized(row[serial_col]):
                active_serial = normalized(row[serial_col])
            case, serial = active_case, active_serial
            rma = normalized(row[rma_col])
            if not (case.startswith('CS-') and serial and rma):
                continue
            # The export contains actual XD-prefixed Supermicro RMA IDs.
            if not re.fullmatch(r'[A-Z]{1,5}\d{6,}', rma):
                continue
            wo = normalize_wo(row[wo_col]) if row[wo_col] else ''
            index[(case, serial)][wo].add(rma)
    finally:
        source.close()
    return index


def choose_thermal_rmas(index, case, serial, existing_work_orders):
    """Return only RMA IDs associated with this exact CS/serial.

    If the row has specific work orders and the source identifies them,
    restrict the RMA IDs to those WOs; otherwise use this exact CS/serial.
    """
    by_wo = index.get((case, serial))
    if not by_wo:
        return set()
    work_orders = parse_work_orders(existing_work_orders)
    matched_wos = work_orders & set(by_wo)
    if matched_wos:
        return set().union(*(by_wo[wo] for wo in matched_wos))
    return set().union(*by_wo.values())

def main():
    if not THERMAL_FILE.is_file():
        raise FileNotFoundError(f'Workbook not found: {THERMAL_FILE}')
    outstanding_file = locate_outstanding()
    thermal_issues_file = locate_thermal_issues()
    thermal_rma_index = load_thermal_rmas(thermal_issues_file)

    open_wb = load_workbook(outstanding_file, read_only=True, data_only=True)
    try:
        open_sheet = open_wb['CoreWeave Outstanding Cases']
        cols = headers_for(open_sheet, 1)
        case_idx = required(cols, 'CASE NUMBER') - 1
        serial_idx = required(cols, 'SERIAL NUMBER') - 1
        status_idx = required(cols, 'STATUS') - 1
        subject_idx = cols.get('SUBJECT', 0) - 1
        rma_idx = next((cols[k] - 1 for k in ('RMA NUMBER', 'RMA NUMBERS', 'RMA #') if k in cols), None)

        cases_by_serial = defaultdict(lambda: defaultdict(set))
        case_details = {}
        for row in open_sheet.iter_rows(min_row=2, values_only=True):
            case = normalize_cs(row[case_idx])
            if not case.startswith('CS-'):
                continue
            serial = normalized(row[serial_idx])
            subject = row[subject_idx] if subject_idx >= 0 else None
            status = row[status_idx]
            ids = rma_numbers(subject)
            if rma_idx is not None and row[rma_idx] is not None:
                ids.update(rma_numbers('RMA ' + str(row[rma_idx])))
            if case not in case_details:
                case_details[case] = {'status': status, 'rmas': set(), 'statuses': set()}
            detail = case_details[case]
            detail['rmas'].update(ids)
            if status is not None and str(status).strip():
                detail['statuses'].add(str(status).strip())
                detail['status'] = status
            if serial:
                cases_by_serial[serial][case].update(jira_codes(subject))
    finally:
        open_wb.close()

    wb = load_workbook(THERMAL_FILE)
    sheet = wb['SN_H100']
    headers = headers_for(sheet, 3)
    serial_col = required(headers, 'SERVER SERIAL')
    case_col = required(headers, 'SUPERMICRO CASE #')
    status_col = required(headers, 'STATUS')
    rma_col = required(headers, 'RMA NUMBERS')
    jira_col = headers.get('JIRA TICKET')
    work_orders_col = headers.get('WORK ORDERS')
    n_cols = sheet.max_column
    original_count = sheet.max_row - 3

    closed_rows, closed_row_numbers = [], []
    changes, ambiguous = [], []
    new_cs_count = status_change_count = rma_change_count = 0

    for row_num in range(4, sheet.max_row + 1):
        cs_cell = sheet.cell(row_num, case_col)
        was_missing_case = not normalized(cs_cell.value)
        case = normalize_cs(cs_cell.value)
        serial = normalized(sheet.cell(row_num, serial_col).value)

        if case:
            if case.startswith('CS-') and case not in case_details:
                cells = [sheet.cell(row_num, col) for col in range(1, n_cols + 1)]
                closed_rows.append((row_num, [c.value for c in cells],
                                    [c.number_format for c in cells]))
                closed_row_numbers.append(row_num)
                continue
            if case not in case_details:
                continue  # Unrecognized/non-CS value: do not change it.
            # Normalize only confirmed case IDs, without changing any other cells.
            if cs_cell.value != case:
                cs_cell.value = case
        else:
            options = cases_by_serial.get(serial, {})
            if not options:
                continue
            if len(options) == 1:
                case = next(iter(options))
            else:
                jira = jira_codes(sheet.cell(row_num, jira_col).value) if jira_col else set()
                candidates = [c for c, subjects in options.items() if jira & subjects]
                if len(candidates) != 1:
                    ambiguous.append((row_num, serial, ', '.join(sorted(options))))
                    continue
                case = candidates[0]
            cs_cell.value = case
            new_cs_count += 1

        detail = case_details[case]
        previous_status = sheet.cell(row_num, status_col).value
        previous_rma = sheet.cell(row_num, rma_col).value
        new_status = previous_status
        new_rma = previous_rma
        # If duplicate outstanding records disagree on status, do not guess.
        if len(detail['statuses']) == 1:
            new_status = next(iter(detail['statuses']))
            if str(new_status).strip() != str(previous_status or '').strip():
                sheet.cell(row_num, status_col).value = new_status
                status_change_count += 1
        elif not previous_status and cs_cell.value == case:
            new_status = 'At Supermicro'
            sheet.cell(row_num, status_col).value = new_status
            status_change_count += 1

        # Fill blanks only. Never overwrite or extend an existing RMA field.
        new_rma = previous_rma
        if not normalized(previous_rma):
            source_rmas = choose_thermal_rmas(
                thermal_rma_index, case, serial,
                sheet.cell(row_num, work_orders_col).value if work_orders_col else None)
            if source_rmas:
                new_rma = ', '.join(sorted(source_rmas))
            elif detail['rmas']:
                new_rma = merge_rmas(previous_rma, detail['rmas'])
        if new_rma != previous_rma:
            sheet.cell(row_num, rma_col).value = new_rma
            rma_change_count += 1

        if (was_missing_case or normalized(previous_status) != normalized(new_status)
                or previous_rma != new_rma):
            changes.append((row_num, serial, case if was_missing_case else '',
                            previous_status, new_status, previous_rma, new_rma))

    for row_num in reversed(closed_row_numbers):
        sheet.delete_rows(row_num, 1)

    if 'Status Dashboard' in wb:
        del wb['Status Dashboard']
    dash = wb.create_sheet('Status Dashboard')
    dash.sheet_view.showGridLines = False
    dash.merge_cells('A1:L1')
    dash['A1'] = 'H100 Thermal | Status Dashboard'
    dash['A1'].font = Font(size=16, bold=True, color='FFFFFF')
    dash['A1'].fill = PatternFill('solid', fgColor='183153')
    dash.row_dimensions[1].height = 32

    stats = [
        ('New CS numbers added', new_cs_count),
        ('Status values updated', status_change_count),
        ('RMA values updated', rma_change_count),
        ('Closed rows removed from thermal sheet', len(closed_rows)),
        ('Ambiguous matches skipped', len(ambiguous)),
        ('Rows remaining in SN_H100', original_count - len(closed_rows)),
    ]
    for r, (label, value) in enumerate(stats, 3):
        dash.cell(r, 1, label).font = Font(bold=True, color='183153')
        dash.cell(r, 2, value).font = Font(bold=True, size=12)

    dash['A10'] = 'Open cases | CS, status and RMA changes'
    dash['A10'].font = Font(size=13, bold=True, color='183153')
    change_headers = ['Original thermal row', 'Server serial', 'CS ticket added',
                      'Previous status', 'Current CS status', 'Previous RMA numbers',
                      'Current RMA numbers']
    style_header(dash, 11, change_headers)
    for row_out, record in enumerate(changes, start=12):
        for col, value in enumerate(record, 1):
            cell = dash.cell(row_out, col, value)
            if row_out % 2 == 0:
                cell.fill = PatternFill('solid', fgColor='EDF3F8')
        dash.cell(row_out, 2).number_format = '@'
        dash.cell(row_out, 3).number_format = '@'

    closed_title_row = 13 + len(changes)
    dash.cell(closed_title_row, 1, 'Closed cases (removed from SN_H100)')
    dash.cell(closed_title_row, 1).font = Font(size=13, bold=True, color='183153')
    original_headers = [sheet.cell(3, col).value for col in range(1, n_cols + 1)]
    closed_header_row = closed_title_row + 1
    style_header(dash, closed_header_row, original_headers + ['Original thermal row'])
    for offset, (original_row, values, formats) in enumerate(closed_rows, start=1):
        dest = closed_header_row + offset
        for col, value in enumerate(values, 1):
            cell = dash.cell(dest, col, value)
            cell.number_format = formats[col - 1]
            if offset % 2 == 0:
                cell.fill = PatternFill('solid', fgColor='EDF3F8')
        dash.cell(dest, status_col).value = 'Closed'
        dash.cell(dest, n_cols + 1, original_row)

    widths = {'A': 28, 'B': 24, 'C': 24, 'D': 24, 'E': 28,
              'F': 30, 'G': 30, 'H': 22, 'I': 22, 'J': 22,
              'K': 25, 'L': 24, 'M': 23}
    for col, width in widths.items():
        dash.column_dimensions[col].width = width
    dash.freeze_panes = 'A12'
    left_align_used_cells(sheet)
    left_align_used_cells(dash)
    OUTSTANDING_DIR.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(OUTPUT_FILE)
    except PermissionError as error:
        raise PermissionError(f'Cannot save {OUTPUT_FILE}. Close it in Excel and retry.') from error
    finally:
        wb.close()
    print(f'Outstanding report: {outstanding_file.name}')
    print(f'RMA source: {thermal_issues_file.name if thermal_issues_file else "Not found in Input (RMA lookup skipped)"}')
    print(f'New CS numbers: {new_cs_count}; status changes: {status_change_count}; RMA updates: {rma_change_count}.')
    print(f'Closed rows moved to dashboard: {len(closed_rows)}; ambiguous rows skipped: {len(ambiguous)}.')
    for row_num, serial, options in ambiguous:
        print(f'  Ambiguous row {row_num}: {serial} -> {options}')
    print(f'Saved: {OUTPUT_FILE}')
    print('Original thermal input and NVIDIA GPU serial values have not been changed.')


if __name__ == '__main__':
    main()
