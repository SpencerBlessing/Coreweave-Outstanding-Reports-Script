"""CoreWeave Salesforce report processing. Usable from CLI or OutstandingReport_GUI."""
from __future__ import annotations

import argparse
from pathlib import Path
from openpyxl.utils import get_column_letter
import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
INPUT_FOLDER = BASE_DIR / 'Reports' / 'Coreweave_Outstanding' / 'Input'
OUTPUT_FOLDER = BASE_DIR / 'Reports' / 'Coreweave_Outstanding' / 'Output'
ADDRESS_FILE = BASE_DIR / 'Address Cache' / 'address_input.xlsx'
NEW_ADDRESS_FILE = INPUT_FOLDER / 'new_address.xlsx'
CASE_FILE = INPUT_FOLDER / 'case_input.xlsx'
FAULTY_PART_FILE = INPUT_FOLDER / 'faulty_part_input.xlsx'
RMA_FILE = INPUT_FOLDER / 'rma_input.xlsx'
OUTPUT_FILE = OUTPUT_FOLDER / 'outstanding_case_output.xlsx'

from Utilities.AddressStandardizer import SiteRegistry, standardize_address, address_key, extract_site
from Utilities.CaseNumberStandardizer import standardize_case_number
from Utilities.PartStandardizer import process_faulty_parts
from Utilities.RMAEnrichment import enrich_outstanding_with_rma
from Utilities.SiteEnrichment import enrich_sites
from Utilities.ReportHistory import log_report_history
from Utilities.ReportComparison import prepare_comparison, write_comparison, save_report_run
from Utilities.ReportDashboard import write_dashboard

LOCATION_COLUMNS = ('City', 'State', 'Region')
ADDRESS_COLUMNS = ('Case Number', 'Asset Location', *LOCATION_COLUMNS)


def read_salesforce_excel(path):
    raw = pd.read_excel(path, header=None)
    header_row = next((i for i, row in raw.iterrows()
                       if any(str(v).strip() == 'Case Number' for v in row)), None)
    if header_row is None:
        raise ValueError(f"Could not find 'Case Number' header in {path}")
    df = pd.read_excel(path, header=header_row, dtype={'Case Number': str})
    df.columns = df.columns.astype(str).str.strip()
    return df.dropna(axis=1, how='all')


def clean_cases(df):
    df = df.copy()
    df['Case Number'] = df['Case Number'].apply(standardize_case_number)
    valid = df['Case Number'].notna() & df['Case Number'].astype(str).str.strip().ne('')
    valid &= ~df['Case Number'].astype(str).str.strip().str.casefold().eq('total')
    return df.loc[valid].copy()


def is_filled(value):
    return pd.notna(value) and str(value).strip() != ''


def process_addresses(registry=None, status=lambda message: None):
    """Update the existing persistent cache; never discard historical cases."""
    status('Loading Salesforce address data...')
    incoming = clean_cases(read_salesforce_excel(NEW_ADDRESS_FILE))
    if 'Asset Location' not in incoming:
        raise ValueError('Address export is missing Asset Location')
    for col in ADDRESS_COLUMNS:
        if col not in incoming:
            incoming[col] = pd.NA
    incoming = incoming[list(ADDRESS_COLUMNS)].copy()
    incoming['Asset Location'] = incoming['Asset Location'].apply(standardize_address)

    if ADDRESS_FILE.exists():
        old = clean_cases(pd.read_excel(ADDRESS_FILE, dtype={'Case Number': str}))
        for col in ADDRESS_COLUMNS:
            if col not in old:
                old[col] = pd.NA
        old = old[list(ADDRESS_COLUMNS)].copy()
        old['Asset Location'] = old['Asset Location'].apply(standardize_address)
    else:
        old = pd.DataFrame(columns=ADDRESS_COLUMNS)

    known_locations = {
        'coreweave c/o switch rno sector 1, sparks, nv 89437, usa': {
            'City': 'Sparks', 'State': 'NV', 'Region': 'West Coast (US)'
        },
    }

    def correct(row):
        location = row.get('Asset Location')
        if is_filled(location):
            values = known_locations.get(str(location).strip().casefold())
            if values:
                for col, value in values.items():
                    if not is_filled(row.get(col)):
                        row[col] = value
        return row

    # A location can provide fallback values only when there is no conflict.
    address_values = {}
    conflicting_keys = set()
    for row in old.to_dict('records'):
        key = address_key(row['Asset Location'])
        if not key:
            continue
        values = tuple(str(row[col]).strip() if is_filled(row[col]) else ''
                       for col in LOCATION_COLUMNS)
        prior = address_values.get(key, ('', '', ''))
        if any(a and b and a != b for a, b in zip(prior, values)):
            conflicting_keys.add(key)
        address_values[key] = tuple(a or b for a, b in zip(prior, values))

    cache = {row['Case Number']: row for row in old.to_dict('records')}
    for row in incoming.to_dict('records'):
        case = row['Case Number']
        previous = cache.get(case)
        if previous:
            old_key = address_key(previous['Asset Location'])
            new_key = address_key(row['Asset Location'])
            if not new_key and old_key:
                row['Asset Location'] = previous['Asset Location']
                new_key = old_key
            if old_key == new_key:
                for col in LOCATION_COLUMNS:
                    if is_filled(previous[col]):
                        row[col] = previous[col]
        key = address_key(row['Asset Location'])
        if key and key not in conflicting_keys:
            fallback = address_values.get(key)
            if fallback:
                for index, col in enumerate(LOCATION_COLUMNS):
                    if not is_filled(row[col]) and fallback[index]:
                        row[col] = fallback[index]
        cache[case] = correct(row)

    addresses = pd.DataFrame(list(cache.values()), columns=ADDRESS_COLUMNS)
    if not addresses.empty:
        addresses = addresses.apply(correct, axis=1)
    ADDRESS_FILE.parent.mkdir(parents=True, exist_ok=True)
    addresses.to_excel(ADDRESS_FILE, index=False)
    status(f'Address cache updated in place: {len(addresses)} cases.')
    return addresses


def generate_report(since=None, status_callback=None, directory_file=None):
    status = status_callback or (lambda message: print(message, flush=True))
    for path in (NEW_ADDRESS_FILE, CASE_FILE, FAULTY_PART_FILE):
        if not path.is_file():
            raise FileNotFoundError(f'Missing Salesforce report: {path}')
    registry = SiteRegistry.from_workbook(directory_file) if directory_file else None
    addresses = process_addresses(registry, status)
    status('Loading outstanding cases...')
    cases = clean_cases(read_salesforce_excel(CASE_FILE))
    if 'Subject' not in cases:
        raise ValueError('Case export missing Subject')
    cases['Site'] = cases['Subject'].apply(extract_site)
    # Resolve site from subject if possible, without replacing an address-based conflict.
    if registry is not None:
        addresses = registry.enrich(addresses.assign(Site=addresses['Case Number'].map(
            cases.drop_duplicates('Case Number').set_index('Case Number')['Site'])))

    fields = ['Case Number', 'City', 'State', 'Region']
    fields += [c for c in ('Asset Location', 'Matched LoCode', 'Matched Site',
                           'Match Status', 'Match Method', 'Standardized Address',
                           'Directory Region', 'Directory Facility', 'Match Candidates') if c in addresses]
    lookup = addresses[fields].drop_duplicates('Case Number', keep='last')
    overlap = [c for c in lookup if c != 'Case Number' and c in cases]
    if overlap:
        cases = cases.drop(columns=overlap)
    before = len(cases)
    cases = cases.merge(lookup, on='Case Number', how='left', validate='many_to_one')
    if len(cases) != before:
        raise RuntimeError('Address merge changed case row count')

    if since:
        if 'Case Creation Date' not in cases:
            raise ValueError('Could not find Case Creation Date column')
        dates = pd.to_datetime(cases['Case Creation Date'], errors='coerce')
        cases = cases.loc[dates >= pd.Timestamp(since)].copy()
    status(f'Cases remaining: {len(cases)}')

    status('Standardizing GPU parts...')
    parts = clean_cases(read_salesforce_excel(FAULTY_PART_FILE))
    parts = process_faulty_parts(parts)
    before = len(cases)
    cases = cases.merge(parts, on='Case Number', how='left', validate='many_to_one')
    if len(cases) != before:
        raise RuntimeError('Part merge changed case row count')

    # Optional RMA source: enrich only cases already present in outstanding.
    if RMA_FILE.is_file():
        status('Loading RMA report...')
        rma = read_salesforce_excel(RMA_FILE)
        cases = enrich_outstanding_with_rma(cases, rma, status)
    else:
        status(f'RMA input not found; skipping enrichment: {RMA_FILE.name}')

    # Enrich missing site identifiers.
    cases = enrich_sites(
        cases,
        registry=registry,
        cache_path=BASE_DIR / 'Address Cache' / 'site_mapping_cache.json',
        status=status,
    )

    cols = list(cases.columns)
    
    if 'Site' in cols:
        cols.remove('Site')
        cols.insert(cols.index('Subject') + 1, 'Site')
        cases = cases[cols]
    # Classify H100 thermal cases before splitting remaining worksheets.
    is_h100 = cases['Faulty Part Number'].fillna('').astype(str).str.contains(
        'H100', case=False, regex=False)
    is_overheating = cases['Faulty Part Description'].fillna('').astype(str).str.contains(
    'GPU Over Temp', case=False, regex=False
    )

    is_thermal_issue = cases['Issue Sub-Type Detail'].fillna('').astype(str).str.strip().str.casefold().eq(
        'thermal issue'
    )

    is_thermal_root_cause = (
    cases['Root Cause: Root Cause Name']
    .fillna('')
    .astype(str)
    .str.contains(
        r'therm|theral|tlimit|over.?temp|temperature|overheat',
        case=False,
        regex=True
    )
    )

    is_gpu_temp = (
        cases['GPU Failure Codes']
        .fillna('')
        .astype(str)
        .str.contains('temp', case=False, regex=False)
    )

    is_h100_thermal = is_h100 & (
        is_overheating |
        is_thermal_issue |
        is_thermal_root_cause |
        is_gpu_temp
    )
    is_self = (
        cases['Subject'].fillna('').astype(str).str.upper().str.startswith('SELF')
        | cases['Issue Sub-Type'].fillna('').astype(str).str.strip().str.casefold().eq('self maintenance')
    )
    categories = pd.Series('CoreWeave Outstanding Cases', index=cases.index)
    categories.loc[is_self] = 'Self-Maintenance'
    categories.loc[is_h100_thermal] = 'H100 Thermal Issues'

    history_folder = BASE_DIR / 'Reports' / 'Coreweave_Outstanding' / 'History'
    status('Comparing against previous report...')
    comparison = prepare_comparison(cases, categories, history_folder)

    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    status('Writing workbook...')
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        cases.loc[categories.eq('CoreWeave Outstanding Cases')].to_excel(
            writer, sheet_name='CoreWeave Outstanding Cases', index=False)
        cases.loc[categories.eq('Self-Maintenance')].to_excel(
            writer, sheet_name='Self-Maintenance', index=False)
        cases.loc[categories.eq('H100 Thermal Issues')].to_excel(
            writer, sheet_name='H100 Thermal Issues', index=False)
        write_comparison(writer, comparison)
        status('Generating Excel dashboard...')
        write_dashboard(writer, cases, categories, comparison, history_folder)

        # Automatically size columns on all worksheets except Dashboard.
        for ws in writer.book.worksheets:
            if ws.title == 'Dashboard':
                continue

            ws.freeze_panes = 'A2'
            ws.sheet_view.showGridLines = False

            for col_idx, column in enumerate(ws.columns, start=1):
                letter = get_column_letter(col_idx)

                max_length = max(
                    (len(str(cell.value)) for cell in column if cell.value is not None),
                    default=0
                )

                ws.column_dimensions[letter].width = min(
                    max(max_length + 3, 14), 65
                )

    # Save after workbook generation, so the next run compares against this one.
    run_id = save_report_run(cases, categories, history_folder)
    status(f'Comparison saved against previous report; snapshot run #{run_id}.')
    log_report_history(cases, categories, history_folder, status=status)
    status('CoreWeave report completed successfully.')
    return OUTPUT_FILE


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--since', help='Include cases created on or after YYYY-MM-DD')
    parser.add_argument('--directory', type=Path, help='Site directory workbook')
    args = parser.parse_args()
    generate_report(since=args.since, directory_file=args.directory)
