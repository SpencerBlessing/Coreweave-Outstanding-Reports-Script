"""Validate incoming CoreWeave addresses against CoreWeave_Map.json only.

The map owns canonical addresses and Site/location metadata. Incoming exports are
observations, never a source of corrected map values. No source files are changed.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import pandas as pd

try:
    from .AddressStandardizer import SiteRegistry, address_key
except ImportError:
    from AddressStandardizer import SiteRegistry, address_key


DEFAULT_MAP = Path(__file__).resolve().parents[1] / 'Address Cache' / 'CoreWeave_Map.json'


class AddressValidator:
    REQUIRED_COLUMNS = ('Case Number', 'Asset Location')
    LOCATION_COLUMNS = ('City', 'State', 'Region')
    ISSUE_STATUSES = ('Missing Site', 'Mismatch', 'Conflict', 'Unresolved', 'No Site or LoCode in Map', 'Unverified Site (LoCode Only)')

    def __init__(self, address_file: str | Path, site_map_path: str | Path | None = None):
        self.address_file = Path(address_file)
        self.site_map_path = Path(site_map_path) if site_map_path else DEFAULT_MAP
        if not self.site_map_path.is_file():
            raise FileNotFoundError(f'CoreWeave map not found: {self.site_map_path}')
        with self.site_map_path.open(encoding='utf-8') as handle:
            self.site_map = json.load(handle)
        if not isinstance(self.site_map, dict) or not all(
            isinstance(self.site_map.get(k), dict) for k in ('by_address', 'by_locode', 'conflicts')
        ):
            raise ValueError('CoreWeave map must contain by_address, by_locode, and conflicts objects')
        self.by_address = self.site_map['by_address']
        self.by_locode = self.site_map['by_locode']
        conflicts = self.site_map['conflicts']
        self.address_conflicts = conflicts.get('address', {})
        self.locode_conflicts = conflicts.get('locode', {})
        # Reuse the same read-only matcher as report generation.
        self.registry = SiteRegistry.from_map(self.site_map_path)
        # Shipping classification comes only from the canonical records in
        # CoreWeave_Map.json. Lookup aliases can omit address_type, so use the
        # resolved canonical address to recognize verified shipping locations.
        self.shipping_address_keys = {
            address_key(record.get('address'))
            for record in self.site_map.get('records', [])
            if isinstance(record, dict)
            and str(record.get('address_type', '')).strip().casefold()
            in {'shipping', 'shipping_only'}
            and address_key(record.get('address'))
        }

    @staticmethod
    def clean_value(value) -> str:
        if value is None or pd.isna(value):
            return ''
        return str(value).strip()

    def load_data(self) -> pd.DataFrame:
        if not self.address_file.is_file():
            raise FileNotFoundError(f'Address file not found: {self.address_file}')
        df = pd.read_excel(self.address_file, dtype={'Case Number': str})
        df.columns = df.columns.astype(str).str.strip()
        missing = set(self.REQUIRED_COLUMNS) - set(df.columns)
        if missing:
            raise ValueError('Missing required columns: ' + ', '.join(sorted(missing)))
        # City/State/Region may be absent from Salesforce; the map is authoritative.
        for col in self.LOCATION_COLUMNS:
            if col not in df.columns:
                df[col] = ''
        return df

    def valid_location_mask(self, df: pd.DataFrame) -> pd.Series:
        """Exclude Salesforce grand-total and empty case rows from validation."""
        values = df['Asset Location'].map(self.clean_value).str.casefold()
        numbers = df['Case Number'].map(self.clean_value).str.casefold()
        return (values.ne('') & values.ne('total')
                & numbers.ne('') & numbers.ne('total'))

    def find_missing_locations(self, df: pd.DataFrame, mapped: pd.DataFrame | None = None) -> pd.DataFrame:
        """Find fields unavailable from both the export AND the CoreWeave map.

        A blank Salesforce field is not a missing location when the authoritative
        map supplies that field. Unknown addresses remain unresolved, not guessed.
        """
        checked = self._mapped(df) if mapped is None else mapped
        missing = pd.Series(False, index=checked.index)
        for col in self.LOCATION_COLUMNS:
            source_blank = checked[col].map(self.clean_value).eq('')
            map_blank = checked[f'Expected {col}'].map(self.clean_value).eq('')
            missing |= source_blank & map_blank
        return checked.loc[missing].copy()

    def find_missing_source_locations(self, df: pd.DataFrame) -> pd.DataFrame:
        """Raw Salesforce omissions, for informational audit only."""
        missing = pd.Series(False, index=df.index)
        for col in self.LOCATION_COLUMNS:
            missing |= df[col].map(self.clean_value).eq('')
        return df.loc[self.valid_location_mask(df) & missing].copy()

    def _resolve(self, row: dict) -> dict:
        """Use exactly the same map resolution rules as the report pipeline."""
        result = self.registry.resolve(
            address=row.get('Asset Location'),
            locode=row.get('LoCode'),
            site_name=row.get('Site'),
        )
        if result.status != 'matched' or result.site is None:
            return {
                'record': None,
                'status': 'conflict' if result.status == 'conflict' else 'unresolved',
                'method': result.method,
            }

        site = result.site
        return {
            'record': {
                'locode': site.locode,
                'site': site.site_code,
                'reporting_site': site.reporting_site,
                'reporting_site_source': site.reporting_site_source,
                'address': site.address,
                'city': site.city,
                'state': site.state,
                'region': site.region,
                'facility': site.facility,
                'data_hall': site.data_hall,
                'customer': site.customer,
            },
            'status': 'matched',
            'method': result.method,
        }

    def _mapped(self, df: pd.DataFrame) -> pd.DataFrame:
        valid = df.loc[self.valid_location_mask(df)].copy()
        valid['Address Key'] = valid['Asset Location'].map(address_key)
        has_site = 'Site' in df.columns
        results = [self._resolve(row) for row in valid.to_dict('records')]

        def field(result, name):
            record = result['record']
            return self.clean_value(record.get(name)) if record else ''

        valid['Match Status'] = [r['status'] for r in results]
        valid['Match Method'] = [r['method'] for r in results]
        valid['Matched LoCode'] = [field(r, 'locode') for r in results]
        valid['Directory Site'] = [field(r, 'site') for r in results]
        valid['Reporting Site'] = [field(r, 'reporting_site') for r in results]
        valid['Reporting Site Source'] = [field(r, 'reporting_site_source') for r in results]
        valid['Directory Address'] = [field(r, 'address') for r in results]
        valid['Directory Region'] = [field(r, 'region') for r in results]
        valid['Match Candidates'] = ''  # Ambiguous Site assignments remain unresolved.
        valid['Standardized Address'] = [
            field(r, 'address') if r['status'] == 'matched' else self.clean_value(raw)
            for r, raw in zip(results, valid['Asset Location'])
        ]
        # A shipping record is an address match with neither a Site nor a
        # LoCode, explicitly classified in the canonical map records.
        shipping = (valid['Match Status'].eq('matched')
                    & valid['Standardized Address'].map(address_key)
                      .isin(self.shipping_address_keys)
                    & valid['Directory Site'].eq('')
                    & valid['Matched LoCode'].eq(''))
        valid['Directory Address Type'] = ''
        valid.loc[shipping, 'Directory Address Type'] = 'shipping'
        valid.loc[shipping, 'Reporting Site'] = 'Shipping'
        valid.loc[shipping, 'Reporting Site Source'] = 'Shipping'
        for location_field in self.LOCATION_COLUMNS:
            valid[f'Expected {location_field}'] = [field(r, location_field.lower()) for r in results]

        def location_verification(row):
            if row['Match Status'] != 'matched':
                return 'Conflict' if row['Match Status'] == 'conflict' else 'Unresolved'
            mismatches = [col for col in self.LOCATION_COLUMNS
                          if self.clean_value(row.get(col)) and self.clean_value(row.get(f'Expected {col}'))
                          and self.clean_value(row[col]).casefold() != self.clean_value(row[f'Expected {col}']).casefold()]
            if mismatches:
                return 'Mismatch: ' + ', '.join(mismatches)
            missing_in_map = [col for col in self.LOCATION_COLUMNS
                              if not self.clean_value(row.get(f'Expected {col}'))]
            if missing_in_map:
                return 'Missing in Map: ' + ', '.join(missing_in_map)
            supplied = [col for col in self.LOCATION_COLUMNS
                        if not self.clean_value(row.get(col)) and self.clean_value(row.get(f'Expected {col}'))]
            return 'Supplied by Map: ' + ', '.join(supplied) if supplied else 'Verified'

        valid['Location Verification'] = [location_verification(row) for row in valid.to_dict('records')]
        valid['Expected Site'] = valid['Reporting Site']

        def site_verification(row):
            if row['Match Status'] != 'matched':
                return 'Conflict' if row['Match Status'] == 'conflict' else 'Unresolved'
            if row['Directory Address Type'] == 'shipping':
                return 'Known Shipping Location'
            expected = self.clean_value(row['Expected Site']).upper()
            if not expected:
                return 'No Site or LoCode in Map'
            fallback = row['Reporting Site Source'] == 'LoCode fallback'
            if not has_site:
                return 'Available via LoCode' if fallback else 'Available in Map'
            existing = self.clean_value(row.get('Site')).upper()
            if fallback:
                if not existing:
                    return 'Available via LoCode'
                if existing == expected:
                    return 'Verified via LoCode'
                # A nonempty subject Site may be correct, but a LoCode is not
                # proof that the separately named internal Site is wrong.
                return 'Unverified Site (LoCode Only)'
            if not existing:
                return 'Missing Site'
            return 'Verified' if existing == expected else 'Mismatch'

        valid['Site Verification'] = [site_verification(row) for row in valid.to_dict('records')]
        return valid

    def find_inconsistent_locations(self, df: pd.DataFrame, mapped: pd.DataFrame | None = None) -> pd.DataFrame:
        """Compare observed location values for the same *canonical* address."""
        valid = self._mapped(df) if mapped is None else mapped
        columns = ['Asset Location', 'Address Key', 'Field', 'Observed Values',
                   'Case Numbers', 'Rows', 'Matched LoCode']
        records = []
        # Collate spelling variants once they have resolved to the same canonical address.
        grouped = valid.assign(_canonical=valid['Standardized Address'].map(address_key))
        for key, group in grouped.groupby('_canonical', sort=False):
            for field_name in self.LOCATION_COLUMNS:
                values = sorted({self.clean_value(x) for x in group[field_name] if self.clean_value(x)})
                if len(values) > 1:
                    records.append({
                        'Asset Location': group['Standardized Address'].iloc[0],
                        'Address Key': key, 'Field': field_name,
                        'Observed Values': ' | '.join(values),
                        'Case Numbers': ', '.join(group['Case Number'].map(self.clean_value).unique()),
                        'Rows': len(group), 'Matched LoCode': group['Matched LoCode'].iloc[0]
                    })
        return pd.DataFrame(records, columns=columns)

    def find_new_locations(self, df: pd.DataFrame, mapped: pd.DataFrame | None = None) -> pd.DataFrame:
        valid = self._mapped(df) if mapped is None else mapped
        unknown = valid.loc[valid['Match Status'].isin(['conflict', 'unresolved'])]
        return unknown.drop_duplicates(subset=['Address Key', 'Match Status']).reset_index(drop=True)

    @classmethod
    def summarize_address_exceptions(cls, site_issues: pd.DataFrame,
                                     blank_addresses: pd.DataFrame) -> pd.DataFrame:
        """One actionable row per normalized address/reason, with all case IDs.

        Uses resolved canonical addresses for matched cases and the Salesforce
        address for unresolved cases. This is a read-only view, not a lookup
        cache or a source of new authoritative mappings.
        """
        columns = ['Asset Location', 'Cases', 'Review Reason', 'Next Action',
                   'Match Status', 'Case Numbers', 'Address Key']
        rows = []
        if not site_issues.empty:
            detail = site_issues.copy()
            detail['_review_address'] = detail['Asset Location'].map(cls.clean_value)
            matched = detail['Match Status'].eq('matched')
            detail.loc[matched, '_review_address'] = (
                detail.loc[matched, 'Standardized Address'].map(cls.clean_value)
            )
            detail['_review_key'] = detail['_review_address'].map(address_key)
            reasons = {
                'No Site or LoCode in Map': 'Missing Site/LoCode',
                'Unresolved': 'Unmatched Address',
                'Conflict': 'Conflicting Assignment',
                'Mismatch': 'Site Mismatch',
                'Missing Site': 'Missing Source Site',
                'Unverified Site (LoCode Only)': 'Unverified Source Site',
            }
            detail['_review_reason'] = detail['Site Verification'].map(
                lambda value: reasons.get(value, value)
            )
            next_actions = {
                'Missing Site/LoCode': 'Verify Site, LoCode or location type in CW-SMC',
                'Unmatched Address': 'Verify address; add an approved map alias if equivalent',
                'Conflicting Assignment': 'Resolve competing Site assignments using CW-SMC',
                'Site Mismatch': 'Review Salesforce Site against the authoritative map',
                'Missing Source Site': 'Review missing Salesforce Site against the map',
                'Unverified Source Site': 'Verify source Site using CW-SMC',
            }
            for (key, reason), group in detail.groupby(
                    ['_review_key', '_review_reason'], sort=False, dropna=False):
                case_numbers = list(dict.fromkeys(
                    case for case in group['Case Number'].map(cls.clean_value)
                    if case
                ))
                rows.append({
                    'Asset Location': group['_review_address'].iloc[0],
                    'Cases': len(group),
                    'Review Reason': reason,
                    'Next Action': next_actions.get(reason, 'Review against CW-SMC and the map'),
                    'Match Status': group['Match Status'].iloc[0],
                    'Case Numbers': ', '.join(case_numbers),
                    'Address Key': key,
                })
        if not blank_addresses.empty:
            cases = list(dict.fromkeys(
                case for case in blank_addresses['Case Number'].map(cls.clean_value)
                if case
            ))
            rows.append({
                'Asset Location': '(blank Salesforce address)',
                'Cases': len(blank_addresses),
                'Review Reason': 'Blank Asset Location',
                'Next Action': 'Populate Asset Location in Salesforce',
                'Match Status': 'missing address',
                'Case Numbers': ', '.join(cases),
                'Address Key': '',
            })
        summary = pd.DataFrame(rows, columns=columns)
        if not summary.empty:
            summary = summary.sort_values(
                ['Cases', 'Review Reason', 'Asset Location'],
                ascending=[False, True, True], kind='stable'
            ).reset_index(drop=True)
        return summary

    def validate(self) -> dict:
        df = self.load_data()
        mapped = self._mapped(df)
        missing = self.find_missing_locations(df, mapped)
        source_missing = self.find_missing_source_locations(df)
        inconsistent = self.find_inconsistent_locations(df, mapped)
        unknown = self.find_new_locations(df, mapped)
        unique_locations = mapped['Address Key'].nunique()
        missing_keys = set(missing['Asset Location'].map(address_key))
        site_counts = mapped['Site Verification'].value_counts().to_dict()
        site_issues = mapped.loc[mapped['Site Verification'].isin(self.ISSUE_STATUSES)].copy()
        shipping_locations = mapped.loc[mapped['Site Verification'].eq('Known Shipping Location')].copy()
        case_numbers = df['Case Number'].map(self.clean_value)
        blank_addresses = df.loc[
            df['Asset Location'].map(self.clean_value).eq('')
            & case_numbers.ne('')
            & case_numbers.str.casefold().ne('total')
        ].copy()
        address_exceptions = self.summarize_address_exceptions(
            site_issues, blank_addresses
        )
        # These categories correspond to the operational report's Site-first,
        # LoCode-fallback rule. A map Site omission is NOT a failure when its
        # verified LoCode is available.
        verification = mapped['Site Verification']
        site_resolution = {
            'map_site': int(mapped['Reporting Site Source'].eq('Site').sum()),
            'locode_fallback': int(mapped['Reporting Site Source'].eq('LoCode fallback').sum()),
            'shipping': int(verification.eq('Known Shipping Location').sum()),
            'no_identifier': int(verification.eq('No Site or LoCode in Map').sum()),
            'conflict': int(mapped['Match Status'].eq('conflict').sum()),
            'unresolved': int(mapped['Match Status'].eq('unresolved').sum()),
        }
        return {
            'dataframe': df, 'mapped': mapped, 'total_rows': len(mapped),
            'unique_locations': unique_locations,
            'complete_locations': max(0, unique_locations - len(missing_keys)),
            'missing': missing, 'source_missing': source_missing, 'inconsistent': inconsistent,
            'new_locations': unknown, 'missing_count': len(missing),
            'inconsistent_count': len(inconsistent),
            'new_location_count': unknown['Address Key'].nunique(),
            'source_missing_count': len(source_missing),
            'match_counts': mapped['Match Status'].value_counts().to_dict(),
            'site_counts': site_counts, 'site_resolution': site_resolution,
            'site_issues': site_issues, 'shipping_locations': shipping_locations,
            'address_exceptions': address_exceptions, 'blank_addresses': blank_addresses,
            'location_issues': mapped.loc[mapped['Location Verification'] != 'Verified'].copy(),
        }


def main():
    parser = argparse.ArgumentParser(description='Validate address exports using CoreWeave_Map.json')
    parser.add_argument('address_file', help='Excel address export to validate')
    parser.add_argument('--site-map', default=str(DEFAULT_MAP), help='Path to CoreWeave_Map.json')
    parser.add_argument('--output', help='Destination Excel validation workbook')
    args = parser.parse_args()
    result = AddressValidator(args.address_file, site_map_path=args.site_map).validate()
    output = Path(args.output) if args.output else Path(args.address_file).with_name(
        Path(args.address_file).stem + '_Validation.xlsx'
    )
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result['address_exceptions'].to_excel(writer, sheet_name='Address Exceptions', index=False)
        for sheet, key in [('Mapped Records', 'mapped'), ('Missing Fields', 'missing'),
                           ('Inconsistent Values', 'inconsistent'), ('Unresolved Sites', 'new_locations')]:
            result[key].to_excel(writer, sheet_name=sheet, index=False)
    print('Rows:', result['total_rows'])
    print('Map matches:', result['match_counts'])
    print('Site checks:', result['site_counts'])
    print('Missing fields after map lookup:', result['missing_count'])
    print('Raw Salesforce missing fields (informational):', result['source_missing_count'])
    print('Conflicting address groups/fields:', result['inconsistent_count'])
    print('Unresolved unique addresses:', result['new_location_count'])
    print('Saved:', output)


if __name__ == '__main__':
    main()
