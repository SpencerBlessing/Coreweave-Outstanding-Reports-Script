"""CoreWeave Salesforce report processing. Usable from CLI or OutstandingReport_GUI."""
from __future__ import annotations
import argparse
import hashlib
import math
import re
from copy import copy
import json
from pathlib import Path
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.views import Selection
import pandas as pd
BASE_DIR = Path(__file__).resolve().parents[2]
INPUT_FOLDER = BASE_DIR / 'Reports' / 'Coreweave_Outstanding' / 'Input'
OUTPUT_FOLDER = BASE_DIR / 'Reports' / 'Coreweave_Outstanding' / 'Output'
MAP_FILE = BASE_DIR / 'Address Cache' / 'CoreWeave_Map.json'
CASE_FILE = INPUT_FOLDER / 'case_input.xlsx'
# Full Salesforce RMA export. Join only cases in the outstanding-case source.
RMA_FILE = INPUT_FOLDER / 'rma_input.xlsx'
OUTPUT_FILE = OUTPUT_FOLDER / 'outstanding_case_output.xlsx'
# Diagnostic workbooks are written outside Output, which contains shareable reports.
DEBUG_FOLDER = BASE_DIR / 'Reports' / 'Coreweave_Outstanding' / 'Debug'
LOCATION_REVIEW_FILE = DEBUG_FOLDER / 'CoreWeave_Location_Review.xlsx'
# Temporary finalized-history inputs. Not an address cache or official history.
# Official SQLite history is updated only after a report is finalized.
FULL_SNAPSHOT_STAGE_FILE = DEBUG_FOLDER / 'outstanding_full_snapshot_stage.csv'
FULL_SNAPSHOT_STAGE_MANIFEST = DEBUG_FOLDER / 'outstanding_full_snapshot_stage.json'
from Utilities.AddressStandardizer import SiteRegistry, address_key, extract_site
from Utilities.CaseNumberStandardizer import standardize_case_number
from Utilities.ReportHistory import log_report_history
from Utilities.ReportComparison import prepare_comparison, write_comparison
from Utilities.ReportDashboard import write_dashboard
LOCATION_COLUMNS = ('City', 'State', 'Region')
ADDRESS_COLUMNS = ('Case Number', 'Asset Location', *LOCATION_COLUMNS)
# These fields support address matching and validation, but belong in diagnostics,
# not in the daily operational case worksheets.
INTERNAL_LOCATION_COLUMNS = frozenset({
    'LoCode', 'Matched LoCode', 'Matched Site', 'Directory Site',
    'Reporting Site', 'Reporting Site Source', 'Match Status', 'Match Method',
    'Standardized Address', 'Directory City', 'Directory State',
    'Directory Region', 'Directory Facility', 'Directory Data Hall',
    'Directory Customer', 'Directory Address Type', 'Match Candidates', 'Site Verification',
})
# One standardized location block. All existing business / RMA columns follow it.
REPORT_COLUMN_ORDER = (
    'Case Number', 'Case Creation Date', 'Case Age', 'Product Name',
    'Serial Number', 'Case Owner', 'Status', 'Subject',
    'Site', 'Asset Location', 'City', 'State', 'Region',
    'Faulty Part Number', 'Root Cause: Root Cause Name',
    'Faulty Part Description', 'GPU Failure Codes', 'Issue Sub-Type',
    'Issue Sub-Type Detail', 'Resolution Category', 'Resolution Code',
    'RMA Number', 'UPS Order Number', 'RMA Created Date',
)
# Baseline widths are comfortable enough for normal report values while keeping
# exceptionally long subjects, addresses, and descriptions wrapped and readable.
CASE_COLUMN_WIDTHS = {
    'Case Number': 20,
    'Case Creation Date': 23,
    'Case Age': 22,
    'Product Name': 31,
    'Serial Number': 27,
    'Case Owner': 26,
    'Status': 32,
    'Subject': 54,
    'Site': 19,
    'Asset Location': 60,
    'City': 25,
    'State': 22,
    'Region': 25,
    'Faulty Part Number': 32,
    'Root Cause: Root Cause Name': 40,
    'Faulty Part Description': 54,
    'GPU Failure Codes': 33,
    'Issue Sub-Type': 33,
    'Issue Sub-Type Detail': 38,
    'Resolution Category': 33,
    'Resolution Code': 32,
    'RMA Number': 30,
    'UPS Order Number': 30,
    'RMA Created Date': 24,
}
def format_report_worksheets(workbook):
    """Left-align and comfortably size the generated sheets without touching data.
    The dashboard already defines intentional narrow spacers and a merged-cell
    layout. Its widths/heights remain intact. Other sheets receive practical
    minimum widths, bounded autosizing, wrapping, and expanded row heights.
    Existing fonts, merged cells, fills, and number formats are preserved.
    """
    case_sheets = {
        'CoreWeave Outstanding Cases',
        'H100 Thermal Issues',
        'Self-Maintenance',
    }
    for ws in workbook.worksheets:
        dashboard = ws.title == 'Dashboard'
        operational = ws.title == 'Operational Insights'
        is_case = ws.title in case_sheets
        if is_case:
            ws.freeze_panes = 'A2'
        # Preserve the dashboard's preconfigured spacer columns and widths.
        if not dashboard:
            for col_idx in range(1, ws.max_column + 1):
                letter = get_column_letter(col_idx)
                header = str(ws.cell(1, col_idx).value or '').strip()
                if is_case:
                    minimum = CASE_COLUMN_WIDTHS.get(header, 23)
                    maximum = max(minimum, 65)
                elif operational:
                    minimum = 38 if col_idx == 1 else (33 if col_idx == 2 else 26)
                    maximum = 58 if col_idx <= 2 else 48
                else:
                    minimum, maximum = 23, 55
                # Only ordinary, non-merged text sets width. A long section
                # title spanning several columns should not widen column A.
                longest = len(header) if is_case else 0
                for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
                    cell = row[0]
                    if isinstance(cell, MergedCell) or cell.value is None:
                        continue
                    text = str(cell.value)
                    if cell.data_type == 'f':
                        # Formulas may be long while their display values are short.
                        continue
                    longest = max(longest, max(map(len, text.splitlines()), default=0))
                    if longest >= maximum:
                        break
                width = max(minimum, min(maximum, longest + 3))
                ws.column_dimensions[letter].width = width
        for row in ws.iter_rows():
            needed_lines = 1
            for cell in row:
                if isinstance(cell, MergedCell) or cell.value is None:
                    continue
                alignment = copy(cell.alignment)
                alignment.horizontal = 'left'
                alignment.vertical = 'center'
                alignment.wrap_text = True
                cell.alignment = alignment
                if dashboard:
                    continue  # Keep dashboard's carefully designed row heights.
                # Approximate wrapped lines so Excel does not clip long text.
                # A small inset allows for cell padding/font glyph differences.
                width = ws.column_dimensions[cell.column_letter].width or 20
                capacity = max(9, int(width) - 4)
                lines = sum(
                    max(1, math.ceil(len(part) / capacity))
                    for part in str(cell.value).split('\n')
                )
                needed_lines = max(needed_lines, lines)
            if not dashboard:
                row_number = row[0].row
                current_height = ws.row_dimensions[row_number].height or 18
                # Excel's maximum row height is 409.5 points; keep rare giant
                # descriptions readable without making rows unbounded.
                proposed_height = min(360, max(23, needed_lines * 16 + 7))
                ws.row_dimensions[row_number].height = max(current_height, proposed_height)
        ws.sheet_view.showGridLines = False
def report_view(cases: pd.DataFrame) -> pd.DataFrame:
    """Exactly 24 operational columns; diagnostics stay in Location Review.
    Salesforce and RMA exports can add extra or unnamed columns. Keep those
    available for upstream enrichment and auditing, but do not expose them in
    the three standardized case worksheets. Optional missing fields remain
    blank so all case worksheets retain the same fixed schema.
    """
    missing_columns = [column for column in REPORT_COLUMN_ORDER
                       if column not in cases.columns]
    visible = cases.copy()
    for column in missing_columns:
        visible[column] = pd.NA
    # Keep exactly 24 columns and one row per CS ticket. When the RMA export
    # contains multiple part rows for a case, display the actual quantities in
    # the existing Faulty Part Number column (for example, "3 x GPU-FRU-...").
    # Non-RMA cases retain their original standardized faulty-part values.
    if 'RMA Part Breakdown' in visible.columns:
        breakdown = visible['RMA Part Breakdown'].fillna('').astype(str).str.strip()
        has_parts = breakdown.ne('')
        visible.loc[has_parts, 'Faulty Part Number'] = breakdown.loc[has_parts]
    return visible.loc[:, list(REPORT_COLUMN_ORDER)].copy()
# Preserve the established 24-column internal schema for dashboard/history,
# while omitting unsupported empty fields in the three shareable case tabs.
# This is presentation-only: no source data or audit data is discarded.
def display_report_view(visible: pd.DataFrame) -> pd.DataFrame:
    always = {
        'Case Number', 'Case Creation Date', 'Case Age', 'Product Name',
        'Serial Number', 'Case Owner', 'Status', 'Subject',
        'Site', 'Asset Location', 'City', 'State', 'Region',
        'Faulty Part Number',
    }
    columns = [
        name for name in REPORT_COLUMN_ORDER
        if name in always or (
            name in visible.columns and
            visible[name].notna().any() and
            visible[name].fillna('').astype(str).str.strip().ne('').any()
        )
    ]
    return visible.loc[:, columns].copy()
def update_dashboard_part_links(workbook, cases: pd.DataFrame, part_tracking: pd.DataFrame):
    """Refresh legacy RMA-linked KPIs from actual open-case part records.
    New All CoreWeave RMA does not supply RMA Number. Distinct case matches,
    not manufactured IDs, define the 'RMA-LINKED' count.
    """
    if 'Dashboard' not in workbook.sheetnames:
        return
    ws = workbook['Dashboard']
    valid_parts = part_tracking.loc[
        part_tracking['Faulty Part Number'].fillna('').astype(str).str.strip().ne('')
        & part_tracking['Quantity'].fillna(0).gt(0)
    ].copy()
    case_ids = set(valid_parts['Case Number'].astype(str))
    ws['B9'] = f'{len(case_ids):,}'
    # Maintain the existing dashboard layout and totals, replacing only the
    # RMA-linked breakdown values that the old metric incorrectly zeroes.
    parts_by_case = (
        valid_parts[['Case Number', 'Faulty Part Number']]
        .drop_duplicates()
        .groupby('Faulty Part Number')['Case Number'].nunique()
        .to_dict()
    )
    regions = cases.loc[cases['Case Number'].astype(str).isin(case_ids)]
    region_counts = (
        regions['Region'].fillna('').astype(str)
        .value_counts().to_dict()
        if 'Region' in regions else {}
    )
    section = None
    for r in range(1, ws.max_row + 1):
        label = ws.cell(r, 2).value
        if label == 'Faulty Part' and ws.cell(r, 6).value == 'With RMA':
            section = 'part'
            continue
        if label == 'Region' and ws.cell(r, 6).value == 'With RMA':
            section = 'region'
            continue
        if label in {'REGION WORKLOAD', 'AGE DISTRIBUTION', 'CURRENT WORKSHEET TOTALS', 'CASE WORKLOAD BY STATUS AND AGE', 'CASES BY SITE'}:
            section = None
            continue
        if section and label and isinstance(label, str) and ws.cell(r, 6).value is not None:
            if section == 'part':
                ws.cell(r, 6, int(parts_by_case.get(label, 0)))
            elif section == 'region':
                ws.cell(r, 6, int(region_counts.get(label, 0)))
def resolve_report_sites(cases: pd.DataFrame) -> pd.DataFrame:
    """Display the verified map's site_name, retaining identifier verification.
    `Matched Site` comes from CoreWeave_Map.json `site_name` and is the human-
    readable operational label. `Directory Site` is the separate canonical
    `site` identifier used to verify any Site code extracted from Subject.
    A blank site_name falls back to the map's Site code, then its LoCode.
    Conflicting/unmatched addresses never acquire an inferred Site.
    """
    cases = cases.copy()
    original_site = cases['Site'].fillna('').astype(str).str.strip()
    mapped_name = cases['Matched Site'].fillna('').astype(str).str.strip()
    mapped_site = cases['Directory Site'].fillna('').astype(str).str.strip()
    locode = cases['Matched LoCode'].fillna('').astype(str).str.strip()
    match_status = cases['Match Status'].fillna('').astype(str).str.casefold()
    matched = match_status.eq('matched')
    # Display the map's name (e.g., PHX01), not its Site code (US-WEST-02A).
    # Preserve the authoritative site code and LoCode in separate diagnostics.
    reporting_site = mapped_name.where(mapped_name.ne(''), mapped_site)
    reporting_site = reporting_site.where(reporting_site.ne(''), locode)
    cases['Subject Site'] = original_site
    cases['Site'] = reporting_site.where(matched, '')
    # Subject verification continues to compare the underlying identifiers,
    # never the friendly display name: US-WEST-02A != PHX01 is not a conflict.
    mapped = matched & mapped_site.ne('')
    fallback = matched & mapped_site.eq('') & locode.ne('')
    no_identifier = matched & mapped_site.eq('') & locode.eq('')
    shipping = (no_identifier & cases['Directory Address Type'].fillna('')
                .astype(str).str.casefold().isin({'shipping', 'shipping_only'}))
    cases.loc[shipping, 'Site'] = 'Shipping'
    cases['Site Verification'] = 'Unresolved'
    cases.loc[match_status.eq('conflict'), 'Site Verification'] = 'Conflict'
    cases.loc[mapped & original_site.eq(mapped_site), 'Site Verification'] = 'Verified'
    cases.loc[mapped & original_site.eq(''), 'Site Verification'] = 'Filled from Map'
    cases.loc[mapped & original_site.ne('') & original_site.ne(mapped_site),
              'Site Verification'] = 'Mismatch'
    cases.loc[fallback & original_site.eq(''), 'Site Verification'] = 'Filled from LoCode'
    cases.loc[fallback & original_site.eq(locode), 'Site Verification'] = 'Verified LoCode'
    cases.loc[fallback & original_site.ne('') & original_site.ne(locode),
              'Site Verification'] = 'Unverified Site (LoCode Only)'
    # A site_name without either identifier can display, but remains a map gap.
    cases.loc[no_identifier & ~shipping, 'Site Verification'] = 'Map Site and LoCode Missing'
    cases.loc[shipping, 'Site Verification'] = 'Known Shipping Location'
    return cases
def resolve_missing_address_sites(cases: pd.DataFrame, registry, status=lambda message: None) -> pd.DataFrame:
    """Resolve blank-address cases using exact Site identifiers verified in the map.

    This never invents an address, matches by geography, or replaces a conflicting
    address assignment. A Site identifier must have one unambiguous display Site
    across the authoritative map. Source evidence remains in Location Review.
    """
    cases = cases.copy()
    with registry.map_path.open('r', encoding='utf-8') as handle:
        records = json.load(handle).get('records', [])

    # Multiple records for a Site are fine; contradictory display Sites are not.
    identifier_sites = {}
    for record in records:
        if not isinstance(record, dict):
            continue
        if str(record.get('address_type', '')).strip().casefold() in {'shipping', 'shipping_only'}:
            continue
        site_name = str(record.get('site_name') or '').strip()
        site_code = str(record.get('site') or '').strip()
        locode = str(record.get('locode') or '').strip()
        display = site_name or site_code or locode
        if not display:
            continue
        for identifier in (site_name, site_code, locode):
            if identifier and len(identifier) >= 3:
                identifier_sites.setdefault(identifier.casefold(), set()).add(display)

    unique = {code: next(iter(names)) for code, names in identifier_sites.items()
              if len(names) == 1}

    # A single verified Site code may cover multiple physical locations with
    # different site_name values (for example, RNO2A at US-SPK01/US-SPK02).
    # Retain the exact Site code in that case, without guessing an address.
    # Never apply this exception to ambiguous LoCodes or site_name labels.
    site_code_groups = {}
    for record in records:
        if not isinstance(record, dict):
            continue
        if str(record.get('address_type', '')).strip().casefold() in {'shipping', 'shipping_only'}:
            continue
        site_code = str(record.get('site') or '').strip()
        if site_code:
            site_code_groups.setdefault(site_code.casefold(), set()).add(site_code)
    for code, names in site_code_groups.items():
        if len(names) == 1 and code not in unique:
            unique[code] = next(iter(names))

    if not unique:
        return cases

    blank_address = cases['Asset Location'].fillna('').astype(str).str.strip().eq('')
    unassigned = cases['Site'].fillna('').astype(str).str.strip().eq('')
    eligible = blank_address & unassigned
    if 'Match Status' in cases:
        eligible &= ~cases['Match Status'].fillna('').astype(str).str.casefold().eq('conflict')

    changed = 0
    for index in cases.index[eligible]:
        subject = str(cases.at[index, 'Subject'] or '')
        # Exact whole identifier matches avoid RNO2 accidentally matching RNO2A.
        tokens = set(re.findall(r'(?<![A-Za-z0-9])[A-Za-z0-9][A-Za-z0-9_-]*(?![A-Za-z0-9])', subject))
        matches = {unique[token.casefold()] for token in tokens if token.casefold() in unique}
        if len(matches) != 1:
            continue
        cases.at[index, 'Site'] = next(iter(matches))
        cases.at[index, 'Site Verification'] = 'Verified Subject Site (Address Missing)'
        changed += 1

    status(f'Site-only map verification: {changed} blank-address cases resolved; '
           'addresses remain blank and flagged for review.')
    return cases


def location_review(cases: pd.DataFrame) -> pd.DataFrame:
    """Keep disagreements and unresolved Site assignments in a private debug workbook."""
    status = cases['Match Status'].fillna('unmatched').astype(str).str.casefold()
    verification = cases['Site Verification'].fillna('').astype(str)
    reason = pd.Series('', index=cases.index, dtype='object')
    reason.loc[status.ne('matched')] = 'Unmatched address'
    reason.loc[status.eq('conflict')] = 'Conflicting location assignments'
    reason.loc[status.eq('matched') & verification.eq('Mismatch')] = 'Subject Site differs from map'
    reason.loc[status.eq('matched') & verification.eq('Unverified Site (LoCode Only)')] = 'Subject Site not verified against LoCode'
    reason.loc[status.eq('matched') & verification.eq('Map Site and LoCode Missing')] = 'No Site or LoCode in map'
    reason.loc[status.eq('matched') & verification.eq('Known Shipping Location')] = 'Known shipping destination (informational)'
    reason.loc[verification.eq('Verified Subject Site (Address Missing)')] = 'Site verified in map from Subject; Asset Location missing'
    problem = reason.ne('')
    columns = [
        'Case Number', 'Subject Site', 'Site', 'Matched Site', 'Directory Site',
        'Matched LoCode', 'Asset Location', 'City', 'State', 'Region', 'Subject', 'Match Status',
        'Directory Address Type',
    ]
    review = cases.loc[problem, [c for c in columns if c in cases.columns]].copy()
    review.insert(1, 'Review Reason', reason.loc[problem])
    return review.rename(columns={
        'Site': 'Report Site', 'Matched Site': 'Map Site Name',
        'Directory Site': 'Map Site', 'Matched LoCode': 'Map LoCode',
        'Directory Address Type': 'Map Address Type',
    })
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
def process_addresses(registry=None, status=lambda message: None):
    """Enrich the outstanding cases export using ONLY CoreWeave_Map.json.
    This deliberately does not read, maintain, or write address_input.xlsx or
    infer verified location data from previous case exports.
    """
    registry = registry if registry is not None else SiteRegistry.from_map(MAP_FILE)
    status('Validating outstanding case addresses against CoreWeave_Map.json...')
    incoming = clean_cases(read_salesforce_excel(CASE_FILE))
    if 'Asset Location' not in incoming.columns:
        raise ValueError('Outstanding cases export is missing Asset Location')
    for column in ADDRESS_COLUMNS:
        if column not in incoming.columns:
            incoming[column] = pd.NA
    # Retain supplied LoCode when available to detect contradictory identifiers.
    columns = list(ADDRESS_COLUMNS)
    if 'LoCode' in incoming.columns:
        columns.append('LoCode')
    incoming = incoming[columns].drop_duplicates('Case Number', keep='last').copy()
    addresses = registry.enrich(incoming)
    matched = addresses['Match Status'].eq('matched')
    # The canonical records (not their lookup aliases) identify whether an
    # address is shipping-only. Some aliases omit address_type; relying on an
    # alias would miss those valid shipping destinations. This reads the SAME
    # CoreWeave_Map.json; no external address source or cache is introduced.
    with registry.map_path.open('r', encoding='utf-8') as handle:
        map_records = json.load(handle).get('records', [])
    shipping_keys = {
        address_key(record.get('address'))
        for record in map_records
        if isinstance(record, dict)
        and str(record.get('address_type', '')).strip().casefold()
        in {'shipping', 'shipping_only'}
    }
    addresses['Directory Address Type'] = ''
    shipping_matches = (matched
                        & addresses['Standardized Address'].map(address_key)
                          .isin(shipping_keys))
    addresses.loc[shipping_matches, 'Directory Address Type'] = 'shipping'
    # The single map decides canonical addresses and location fields on matches.
    addresses.loc[matched, 'Asset Location'] = addresses.loc[matched, 'Standardized Address']
    for column, mapped_column in (
        ('City', 'Directory City'),
        ('State', 'Directory State'),
        ('Region', 'Directory Region'),
    ):
        addresses.loc[matched, column] = addresses.loc[matched, mapped_column].fillna('')
    # Unmatched/conflicting addresses retain their original Salesforce values,
    # clearly marked as unverified; none becomes a new lookup or alias.
    counts = addresses['Match Status'].value_counts().to_dict()
    status(f"CoreWeave map validation: {counts.get('matched', 0)} matched, "
           f"{counts.get('conflict', 0)} conflicts, "
           f"{counts.get('unmatched', 0)} unmatched.")
    return addresses
def _case_membership_signature(cases: pd.DataFrame, categories: pd.Series) -> str:
    """Stable checksum of report case IDs and their destination worksheets."""
    numbers = cases['Case Number'].apply(standardize_case_number)
    categories = categories.reindex(cases.index)
    if numbers.isna().any() or numbers.eq('').any() or numbers.duplicated().any():
        raise ValueError('Snapshot must contain unique, nonblank Case Numbers')
    if categories.isna().any():
        raise ValueError('Every case requires a worksheet category')
    membership = sorted(zip(numbers.astype(str), categories.astype(str)))
    return hashlib.sha256(
        json.dumps(membership, ensure_ascii=False, separators=(',', ':')).encode('utf-8')
    ).hexdigest()
def _stage_complete_snapshot(full_cases, full_categories, filtered_cases, filtered_categories):
    """Preserve ALL enriched rows for finalization, independent of the GUI cutoff.
    Only the finalized report updates the official SQLite history. This staging
    file is simply the full, already-enriched input to that later operation.
    """
    DEBUG_FOLDER.mkdir(parents=True, exist_ok=True)
    stage = report_view(full_cases)
    stage['__category'] = full_categories.reindex(full_cases.index)
    if stage['__category'].isna().any():
        raise ValueError('Missing category in full history staging data')
    _case_membership_signature(stage, stage['__category'])
    # Write the new stage before its manifest; the manifest is the commit marker.
    tmp_stage = FULL_SNAPSHOT_STAGE_FILE.with_suffix('.csv.tmp')
    stage.to_csv(tmp_stage, index=False, encoding='utf-8')
    tmp_stage.replace(FULL_SNAPSHOT_STAGE_FILE)
    manifest = {
        'full_case_count': len(stage),
        'displayed_case_count': len(filtered_cases),
        'displayed_case_signature': _case_membership_signature(filtered_cases, filtered_categories),
        'stage_sha256': hashlib.sha256(FULL_SNAPSHOT_STAGE_FILE.read_bytes()).hexdigest(),
    }
    tmp_manifest = FULL_SNAPSHOT_STAGE_MANIFEST.with_suffix('.json.tmp')
    tmp_manifest.write_text(json.dumps(manifest, indent=2), encoding='utf-8')
    tmp_manifest.replace(FULL_SNAPSHOT_STAGE_MANIFEST)
def commit_finalized_history(report_path, status_callback=None):
    """Save an unfiltered snapshot after a matching final report was created.
    The final workbook may use a creation-date cutoff, but the SQLite snapshot
    always contains ALL outstanding cases from the same Salesforce export.
    """
    status = status_callback or (lambda message: print(message, flush=True))
    if not FULL_SNAPSHOT_STAGE_FILE.is_file() or not FULL_SNAPSHOT_STAGE_MANIFEST.is_file():
        raise FileNotFoundError(
            'Full-history staging data is missing. Regenerate the report with the updated '
            'OutstandingReport.py before finalizing; filtered worksheets cannot be used '
            'to save the complete snapshot.'
        )
    manifest = json.loads(FULL_SNAPSHOT_STAGE_MANIFEST.read_text(encoding='utf-8'))
    stage_digest = hashlib.sha256(FULL_SNAPSHOT_STAGE_FILE.read_bytes()).hexdigest()
    if stage_digest != manifest.get('stage_sha256'):
        raise ValueError('Full-history staging data has changed. Regenerate before finalizing.')
    sheet_names = (
        'CoreWeave Outstanding Cases',
        'Self-Maintenance',
        'H100 Thermal Issues',
    )
    frames = []
    with pd.ExcelFile(report_path) as workbook:
        for name in sheet_names:
            if name not in workbook.sheet_names:
                raise ValueError(f'Missing worksheet: {name}')
            frame = pd.read_excel(workbook, sheet_name=name, dtype={'Case Number': str})
            frame['__category'] = name
            frames.append(frame)
    displayed = pd.concat(frames, ignore_index=True)
    displayed_categories = displayed.pop('__category')
    if (len(displayed) != manifest.get('displayed_case_count') or
            _case_membership_signature(displayed, displayed_categories)
            != manifest.get('displayed_case_signature')):
        raise ValueError(
            'Final report does not match the full-history staging data. '
            'Regenerate the report before finalizing.'
        )
    full_cases = pd.read_csv(
        FULL_SNAPSHOT_STAGE_FILE,
        dtype={'Case Number': str},
        keep_default_na=False,
        encoding='utf-8',
    )
    if '__category' not in full_cases.columns:
        raise ValueError('Missing categories in full-history staging data')
    categories = full_cases.pop('__category')
    full_cases['Case Number'] = full_cases['Case Number'].apply(standardize_case_number)
    _case_membership_signature(full_cases, categories)
    if len(full_cases) != manifest.get('full_case_count'):
        raise ValueError('Full-history staged case count does not match the manifest')
    history_folder = BASE_DIR / 'Reports' / 'Coreweave_Outstanding' / 'History'
    status(f'Saving complete finalized history: {len(full_cases):,} cases '
           f'({len(displayed):,} shown in date-filtered report)...')
    db_path, csv_path = log_report_history(
        full_cases,
        categories,
        history_folder,
        status=status,
    )
    status(f'Official unfiltered daily snapshot saved: {db_path.name}')
    return db_path
def filter_open_cases(frame: pd.DataFrame, source_name: str) -> pd.DataFrame:
    """Exclude closed/merged records; retain working and pending statuses.
    Status is not the same as RMA closure: 'Pending RMA Closure' remains open.
    The outstanding export is the membership authority for report tickets.
    """
    if 'Status' not in frame.columns:
        raise ValueError(f'{source_name} is missing Status')
    state = frame['Status'].fillna('').astype(str).str.strip().str.casefold()
    rejected = state.str.fullmatch(
        r'(?:case\s+)?closed(?:\s*[-(].*)?|merged|cancelled|canceled|void|duplicate',
        na=False,
    )
    return frame.loc[~rejected].copy()
def build_part_failure_tracking(rma_export: pd.DataFrame, cases: pd.DataFrame) -> pd.DataFrame:
    """One row per open RMA part from the current open-case population.
    Site/address always come from the existing map-enriched case rows, never
    from RMA text or a second location lookup.
    """
    columns = [
        'Case Number', 'Status', 'Site', 'Asset Location',
        'Faulty Part Number', 'Quantity', 'Faulty Part Serial Number',
        'Root Cause: Root Cause Name', 'Faulty Part Description',
        'Product Name', 'Serial Number', 'Case Creation Date',
    ]
    lookup_fields = [
        'Case Number', 'Status', 'Site', 'Asset Location',
        'Product Name', 'Serial Number', 'Case Creation Date',
    ]
    lookup = cases.loc[:, lookup_fields].drop_duplicates('Case Number')
    open_rma = filter_open_cases(clean_cases(rma_export), 'All CoreWeave RMA')
    open_rma = open_rma.loc[open_rma['Case Number'].isin(set(lookup['Case Number']))].copy()
    if open_rma.empty:
        return pd.DataFrame(columns=columns)
    open_rma = open_rma.merge(
        lookup, on='Case Number', how='inner', suffixes=('_RMA', ''),
        validate='many_to_one',
    )
    pattern = re.compile(r'^(\d+)\s*[xX×]\s*(.+)$')
    rows = []
    for _, record in open_rma.iterrows():
        raw = record.get('Faulty Part Number')
        if pd.isna(raw) or not str(raw).strip():
            parts = ['']
        else:
            parts = re.split(r'\s*\|\s*|[;\r\n]+', str(raw))
        for item in parts:
            item = item.strip()
            if not item and len(parts) > 1:
                continue
            match = pattern.match(item)
            count = int(match.group(1)) if match else (1 if item else 0)
            part = (match.group(2) if match else item).strip()
            def clean(field):
                value = record.get(field)
                return '' if value is None or pd.isna(value) else value
            rows.append({
                'Case Number': clean('Case Number'),
                'Status': clean('Status'),
                'Site': clean('Site'),
                'Asset Location': clean('Asset Location'),
                'Faulty Part Number': part,
                'Quantity': count,
                'Faulty Part Serial Number': clean('Faulty Part Serial Number'),
                'Root Cause: Root Cause Name': clean('Root Cause: Root Cause Name'),
                'Faulty Part Description': clean('Faulty Part Description'),
                'Product Name': clean('Product Name'),
                'Serial Number': clean('Serial Number'),
                'Case Creation Date': clean('Case Creation Date'),
            })
    return pd.DataFrame(rows, columns=columns).sort_values(
        ['Site', 'Case Number', 'Faulty Part Number'], kind='stable'
    ).reset_index(drop=True)
def summarize_outstanding_parts(export: pd.DataFrame, case_numbers) -> pd.DataFrame:
    """Aggregate full RMA rows for current outstanding cases only.
    Repeated part rows represent quantities. This report includes historical
    closed cases; the outstanding-case source controls membership. Root cause
    and faulty-part description are supplementary, never location authorities.
    """
    from collections import Counter, defaultdict
    if 'Faulty Part Number' not in export.columns:
        raise ValueError('All CoreWeave RMA export is missing Faulty Part Number')
    export = filter_open_cases(clean_cases(export), 'All CoreWeave RMA')
    wanted = set(case_numbers.dropna().astype(str))
    grouped = defaultdict(Counter)
    causes = defaultdict(dict)
    descriptions = defaultdict(dict)
    quantity_pattern = re.compile(r'^(\d+)\s*[xX×]\s*(.+)$')
    part_idx = export.columns.get_loc('Faulty Part Number')
    case_idx = export.columns.get_loc('Case Number')
    cause_idx = (export.columns.get_loc('Root Cause: Root Cause Name')
                 if 'Root Cause: Root Cause Name' in export else None)
    desc_idx = (export.columns.get_loc('Faulty Part Description')
                if 'Faulty Part Description' in export else None)
    for row in export.itertuples(index=False, name=None):
        case = row[case_idx]
        if case not in wanted:
            continue
        value = row[part_idx]
        if pd.notna(value):
            for part in re.split(r'\s*\|\s*|[;\r\n]+', str(value)):
                part = part.strip()
                if not part:
                    continue
                match = quantity_pattern.match(part)
                quantity = int(match.group(1)) if match else 1
                code = (match.group(2) if match else part).strip()
                if code:
                    grouped[case][code] += quantity
        for idx, store in ((cause_idx, causes), (desc_idx, descriptions)):
            if idx is not None and pd.notna(row[idx]):
                text = str(row[idx]).strip()
                if text:
                    store[case].setdefault(text.casefold(), text)
    records = []
    for case in sorted(grouped.keys() | causes.keys() | descriptions.keys()):
        counts = grouped[case]
        formatted = ' | '.join(
            f'{count} x {part}' if count > 1 else part
            for part, count in sorted(counts.items())
        )
        records.append({
            'Case Number': case,
            'RMA Part Breakdown': formatted,
            'RMA Root Cause': ' | '.join(causes[case].values()),
            'RMA Faulty Part Description': ' | '.join(descriptions[case].values()),
        })
    return pd.DataFrame(records, columns=[
        'Case Number', 'RMA Part Breakdown',
        'RMA Root Cause', 'RMA Faulty Part Description',
    ])
def generate_report(since=None, status_callback=None, map_file=None):
    status = status_callback or (lambda message: print(message, flush=True))
    for path in (CASE_FILE, RMA_FILE):
        if not path.is_file():
            raise FileNotFoundError(f'Missing Salesforce report: {path}')
    # Fail clearly if the single authoritative map is missing.
    registry = SiteRegistry.from_map(map_file if map_file is not None else MAP_FILE)
    addresses = process_addresses(registry, status)
    status('Loading outstanding cases...')
    cases = filter_open_cases(
        clean_cases(read_salesforce_excel(CASE_FILE)),
        'CoreWeave All Outstanding Cases',
    )
    if 'Subject' not in cases.columns:
        raise ValueError('Case export missing Subject')
    cases['Site'] = cases['Subject'].apply(extract_site)
    # Join map results by Case Number, without learning from historical cases.
    fields = ['Case Number', 'Asset Location', 'City', 'State', 'Region']
    fields += [column for column in (
        'Matched LoCode', 'Matched Site', 'Directory Site',
        'Reporting Site', 'Reporting Site Source',
        'Match Status', 'Match Method', 'Standardized Address',
        'Directory City', 'Directory State', 'Directory Region',
        'Directory Facility', 'Directory Data Hall', 'Directory Customer',
        'Directory Address Type', 'Match Candidates',
    ) if column in addresses.columns]
    lookup = addresses[fields].drop_duplicates('Case Number', keep='last')
    overlap = [column for column in lookup if column != 'Case Number' and column in cases]
    if overlap:
        cases = cases.drop(columns=overlap)
    before = len(cases)
    cases = cases.merge(lookup, on='Case Number', how='left', validate='many_to_one')
    if len(cases) != before:
        raise RuntimeError('Address merge changed case row count')
    # Display the matched map site_name; preserve Site code/LoCode for audit.
    # Unmatched/conflicting addresses never receive a subject-derived Site.
    cases = resolve_report_sites(cases)
    cases = resolve_missing_address_sites(cases, registry, status)
    # Resolution counts describe the displayed value, not the source Site code.
    resolved = cases['Match Status'].fillna('').astype(str).str.casefold().eq('matched')
    names = cases['Matched Site'].fillna('').astype(str).str.strip().ne('')
    site_codes = cases['Directory Site'].fillna('').astype(str).str.strip().ne('')
    locodes = cases['Matched LoCode'].fillna('').astype(str).str.strip().ne('')
    shipping = cases['Site'].eq('Shipping')
    status('Site display: '
           f"{int((resolved & names & ~shipping).sum())} site names, "
           f"{int((resolved & ~names & site_codes & ~shipping).sum())} Site-code fallbacks, "
           f"{int((resolved & ~names & ~site_codes & locodes & ~shipping).sum())} LoCode fallbacks, "
           f"{int(shipping.sum())} known shipping destinations, "
           f"{int(cases['Site'].eq('').sum())} blank.")
    # Do NOT apply the optional date cutoff yet: full-history snapshots
    # need every case, including its parts, RMA data, and worksheet category.
    status(f'Full outstanding case population: {len(cases)}')
    status('Aggregating parts and failure details from All CoreWeave RMA...')
    part_export = read_salesforce_excel(RMA_FILE)
    part_export = part_export.rename(columns={
        'Asset: Product: Product Name': 'Product Name',
        'Asset: Serial Number': 'Serial Number',
    })
    part_tracking = build_part_failure_tracking(part_export, cases)
    part_data = summarize_outstanding_parts(part_export, cases['Case Number'])
    before = len(cases)
    cases = cases.merge(part_data, on='Case Number', how='left', validate='one_to_one')
    if len(cases) != before:
        raise RuntimeError('Part merge changed case row count')
    # Preserve the 24-column internal schema for history and comparison.
    # Unsupported blank fields are omitted from displayed case tabs.
    for target, source_col in (
        ('Faulty Part Number', 'RMA Part Breakdown'),
        ('Root Cause: Root Cause Name', 'RMA Root Cause'),
        ('Faulty Part Description', 'RMA Faulty Part Description'),
    ):
        incoming = cases[source_col].replace('', pd.NA)
        if target not in cases.columns:
            cases[target] = incoming.fillna('')
        else:
            original = cases[target].replace('', pd.NA)
            cases[target] = original.combine_first(incoming).fillna('')
    status(f'RMA enrichment: {len(part_data):,} outstanding cases matched.')
    cols = list(cases.columns)
    if 'Site' in cols:
        cols.remove('Site')
        cols.insert(cols.index('Subject') + 1, 'Site')
        cases = cases[cols]
    # Classify thermal tickets using available fields, including RMA root
    # causes and faulty-part descriptions when supplied by the new export.
    def field_text(name):
        return cases[name].fillna('').astype(str) if name in cases else pd.Series('', index=cases.index)
    is_h100 = field_text('Faulty Part Number').str.contains('H100', case=False, regex=False)
    thermal_pattern = r'therm|theral|tlimit|over[ -]?temp|temperature|overheat|gpu over temp'
    is_thermal_subject = field_text('Subject').str.contains(thermal_pattern, case=False, regex=True)
    is_overheating = field_text('Faulty Part Description').str.contains(
        'GPU Over Temp', case=False, regex=False
    )
    is_thermal_issue = field_text('Issue Sub-Type Detail').str.strip().str.casefold().eq('thermal issue')
    is_thermal_root_cause = field_text('Root Cause: Root Cause Name').str.contains(
        thermal_pattern, case=False, regex=True
    )
    is_gpu_temp = field_text('GPU Failure Codes').str.contains('temp', case=False, regex=False)
    is_h100_thermal = is_h100 & (
        is_thermal_subject | is_overheating | is_thermal_issue |
        is_thermal_root_cause | is_gpu_temp
    )
    is_self = (
        cases['Subject'].fillna('').astype(str).str.upper().str.startswith('SELF')
        | field_text('Issue Sub-Type').str.strip().str.casefold().eq('self maintenance')
    )
    categories = pd.Series('CoreWeave Outstanding Cases', index=cases.index)
    categories.loc[is_self] = 'Self-Maintenance'
    categories.loc[is_h100_thermal] = 'H100 Thermal Issues'
    # Save the full enriched population for official history at finalization.
    # Filtering below affects only the shareable report and its comparisons.
    full_cases = cases
    full_categories = categories
    if since:
        if 'Case Creation Date' not in cases.columns:
            raise ValueError('Could not find Case Creation Date column')
        dates = pd.to_datetime(cases['Case Creation Date'], errors='coerce')
        selected = dates.ge(pd.Timestamp(since))
        cases = cases.loc[selected].copy()
        categories = categories.loc[cases.index].copy()
    status(f'Displayed cases: {len(cases)} / {len(full_cases)} full snapshot cases')
    history_folder = BASE_DIR / 'Reports' / 'Coreweave_Outstanding' / 'History'
    status('Comparing against previous report...')
    # The GUI's optional creation-date cutoff must also scope history deltas.
    # Otherwise older baseline cases appear as false removals.
    comparison = prepare_comparison(cases, categories, history_folder, since=since)
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    status('Writing streamlined workbook...')
    # Use the rich internal data for comparison and the dashboard, but do not
    # expose duplicate location / validation fields on the case worksheets.
    visible_cases = report_view(cases)
    display_cases = display_report_view(visible_cases)
    # This worksheet uses the same visible case-date filter, when supplied.
    # Never include closed/merged or RMA-only historical cases.
    part_tracking = part_tracking.loc[
        part_tracking['Case Number'].isin(set(cases['Case Number']))
    ].copy()
    # Trim consistently blank optional columns from the part detail as well.
    part_required = {'Case Number', 'Site', 'Faulty Part Number', 'Quantity', 'Status'}
    part_tracking = part_tracking.loc[:, [
        column for column in part_tracking.columns
        if column in part_required or (
            part_tracking[column].notna().any()
            and part_tracking[column].fillna('').astype(str).str.strip().ne('').any()
        )
    ]]
    review_cases = location_review(cases)
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        display_cases.loc[categories.eq('CoreWeave Outstanding Cases')].to_excel(
            writer, sheet_name='CoreWeave Outstanding Cases', index=False)
        display_cases.loc[categories.eq('Self-Maintenance')].to_excel(
            writer, sheet_name='Self-Maintenance', index=False)
        display_cases.loc[categories.eq('H100 Thermal Issues')].to_excel(
            writer, sheet_name='H100 Thermal Issues', index=False)
        write_comparison(writer, comparison)
        status('Generating Excel dashboard...')
        # Keep your existing write_dashboard(...) call here
        write_dashboard(writer, visible_cases, categories, comparison, history_folder)
        update_dashboard_part_links(writer.book, cases, part_tracking)
        # Rename Comparison to Operational Insights
        wb = writer.book
        if 'Comparison' in wb.sheetnames:
            wb['Comparison'].title = 'Operational Insights'
        # Reorder worksheets
        sheet_order = [
            'Dashboard',
            'Operational Insights',
            'CoreWeave Outstanding Cases',
            'H100 Thermal Issues',
            'Self-Maintenance',
        ]
        for index, name in enumerate(sheet_order):
            if name in wb.sheetnames:
                wb._sheets.insert(
                    index,
                    wb._sheets.pop(wb.sheetnames.index(name))
                )
        # Keep only the five supported report tabs; remove any extra
        # Unknown/legacy worksheets without dropping the underlying case rows.
        for worksheet in list(wb.worksheets):
            if worksheet.title not in sheet_order:
                del wb[worksheet.title]
        format_report_worksheets(wb)
        # Removing frozen panes alone leaves openpyxl's old bottomLeft
        # selection behind, producing an invalid sheetView in Excel.
        dashboard = wb['Dashboard']
        dashboard.freeze_panes = None
        dashboard.sheet_view.selection = [Selection(activeCell='A1', sqref='A1')]
        dashboard.sheet_view.topLeftCell = 'A1'
        wb.active = wb.sheetnames.index('Dashboard')
    # Keep case-level location diagnostics entirely outside the shareable
    # preview/final workbooks. Always overwrite the debug workbook, even when
    # there are no review rows, so an older review is never mistaken for current.
    DEBUG_FOLDER.mkdir(parents=True, exist_ok=True)
    status('Writing separate location debug workbook...')
    with pd.ExcelWriter(LOCATION_REVIEW_FILE, engine='openpyxl') as writer:
        review_cases.to_excel(writer, sheet_name='Location Review', index=False)
        ws = writer.book['Location Review']
        ws.freeze_panes = 'A2'
        ws.sheet_view.showGridLines = False
        for col_idx, column in enumerate(ws.columns, start=1):
            letter = get_column_letter(col_idx)
            max_length = max(
                (len(str(cell.value)) for cell in column if cell.value is not None),
                default=0,
            )
            ws.column_dimensions[letter].width = min(max(max_length + 3, 14), 65)
    status(f'Location debugging saved separately: {LOCATION_REVIEW_FILE}')
    # Preview generation never commits official history.
    # History is saved only after successful finalization.
    # The SQLite database is not written here. Only stage the unfiltered
    # data so finalization can save the FULL snapshot, not just displayed rows.
    _stage_complete_snapshot(full_cases, full_categories, cases, categories)
    status(f'Full history staged for finalization: {len(full_cases):,} cases.')
    status('CoreWeave preview generated successfully.')
    return OUTPUT_FILE
if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--since', help='Include cases created on or after YYYY-MM-DD')
    parser.add_argument('--site-map', type=Path, help='CoreWeave_Map.json (optional override)')
    args = parser.parse_args()
    generate_report(since=args.since, map_file=args.site_map)
