"""Validate CoreWeave address exports using AddressStandardizer site mapping.

Does not modify input files or assume that the most frequent historical value is correct.
"""
from __future__ import annotations

from pathlib import Path
import argparse
import pandas as pd

try:
    from .AddressStandardizer import SiteRegistry, address_key, standardize_address
except ImportError:
    from AddressStandardizer import SiteRegistry, address_key, standardize_address


class AddressValidator:
    REQUIRED_COLUMNS = ('Case Number', 'Asset Location', 'City', 'State', 'Region')
    LOCATION_COLUMNS = ('City', 'State', 'Region')

    def __init__(self, address_file: str | Path, directory_file: str | Path | None = None,
                 registry: SiteRegistry | None = None):
        self.address_file = Path(address_file)
        self.registry = registry or (SiteRegistry.from_workbook(directory_file) if directory_file else None)

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
        return df

    def valid_location_mask(self, df: pd.DataFrame) -> pd.Series:
        values = df['Asset Location'].map(self.clean_value).str.casefold()
        return values.ne('') & values.ne('total')

    def find_missing_locations(self, df: pd.DataFrame) -> pd.DataFrame:
        missing = pd.Series(False, index=df.index)
        for col in self.LOCATION_COLUMNS:
            missing |= df[col].map(self.clean_value).eq('')
        return df.loc[self.valid_location_mask(df) & missing].copy()

    def _mapped(self, df: pd.DataFrame) -> pd.DataFrame:
        valid = df.loc[self.valid_location_mask(df)].copy()
        valid['Address Key'] = valid['Asset Location'].map(address_key)
        if self.registry is not None:
            resolutions = [self.registry.resolve(address=row.get('Asset Location'),
                          locode=row.get('LoCode'), site_name=row.get('Site'))
                           for row in valid.to_dict('records')]
            valid['Match Status'] = [r.status for r in resolutions]
            valid['Match Method'] = [r.method for r in resolutions]
            valid['Matched LoCode'] = [r.site.locode if r.site else '' for r in resolutions]
            valid['Directory Address'] = [r.site.address if r.site else '' for r in resolutions]
            valid['Directory Region'] = [r.site.region if r.site else '' for r in resolutions]
            valid['Match Candidates'] = [', '.join(r.candidates) for r in resolutions]
            valid['Standardized Address'] = [standardize_address(row.get('Asset Location'), self.registry,
                    row.get('LoCode'), row.get('Site')) for row in valid.to_dict('records')]
        else:
            valid['Match Status'] = 'not_checked'
            valid['Match Method'] = ''
            valid['Matched LoCode'] = ''
            valid['Directory Address'] = ''
            valid['Directory Region'] = ''
            valid['Match Candidates'] = ''
            valid['Standardized Address'] = valid['Asset Location'].map(standardize_address)
        return valid

    def find_inconsistent_locations(self, df: pd.DataFrame) -> pd.DataFrame:
        """Report conflicting observed values; never promote majority to verified truth."""
        valid = self._mapped(df)
        if valid.empty:
            return pd.DataFrame(columns=['Asset Location', 'Field', 'Observed Values', 'Case Numbers'])
        records = []
        for key, group in valid.groupby('Address Key', sort=False):
            for field in self.LOCATION_COLUMNS:
                values = sorted({self.clean_value(x) for x in group[field] if self.clean_value(x)})
                if len(values) > 1:
                    records.append({'Asset Location': group['Asset Location'].iloc[0],
                                    'Address Key': key, 'Field': field,
                                    'Observed Values': ' | '.join(values),
                                    'Case Numbers': ', '.join(group['Case Number'].map(self.clean_value).unique()),
                                    'Rows': len(group),
                                    'Matched LoCode': group['Matched LoCode'].iloc[0]})
        return pd.DataFrame(records, columns=['Asset Location', 'Address Key', 'Field',
                    'Observed Values', 'Case Numbers', 'Rows', 'Matched LoCode'])

    def find_new_locations(self, df: pd.DataFrame) -> pd.DataFrame:
        """With a directory: unmatched/ambiguous/conflicting sites. Otherwise all unique addresses."""
        mapped = self._mapped(df)
        if self.registry is not None:
            mapped = mapped.loc[mapped['Match Status'] != 'matched']
        return mapped.drop_duplicates(subset=['Address Key']).reset_index(drop=True)

    def validate(self) -> dict:
        df = self.load_data()
        mapped = self._mapped(df)
        missing = self.find_missing_locations(df)
        inconsistent = self.find_inconsistent_locations(df)
        unknown = self.find_new_locations(df)
        unique_locations = mapped['Address Key'].nunique()
        missing_keys = set(missing['Asset Location'].map(address_key))
        return {'dataframe': df, 'mapped': mapped,
                'total_rows': len(mapped), 'unique_locations': unique_locations,
                'complete_locations': max(0, unique_locations - len(missing_keys)),
                'missing': missing, 'inconsistent': inconsistent, 'new_locations': unknown,
                'missing_count': len(missing), 'inconsistent_count': len(inconsistent),
                'new_location_count': len(unknown),
                'match_counts': mapped['Match Status'].value_counts().to_dict()}

    @staticmethod
    def _count_unique_missing_locations(df: pd.DataFrame) -> int:
        return df['Asset Location'].map(address_key).nunique() if not df.empty else 0


def main():
    parser = argparse.ArgumentParser(description='Validate CoreWeave address data and site mappings')
    parser.add_argument('address_file', help='Excel report or address cache')
    parser.add_argument('--directory', help='CoreWeave site directory workbook')
    parser.add_argument('--output', help='Validation workbook destination')
    args = parser.parse_args()
    result = AddressValidator(args.address_file, args.directory).validate()
    output = Path(args.output) if args.output else Path(args.address_file).with_name(Path(args.address_file).stem + '_Validation.xlsx')
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet, key in [('Mapped Records', 'mapped'), ('Missing Fields', 'missing'),
                           ('Inconsistent Values', 'inconsistent'), ('Unresolved Sites', 'new_locations')]:
            result[key].to_excel(writer, sheet_name=sheet, index=False)
    print('Rows:', result['total_rows'])
    print('Site matches:', result['match_counts'])
    print('Missing fields:', result['missing_count'])
    print('Conflicting address groups/fields:', result['inconsistent_count'])
    print('Unresolved unique addresses:', result['new_location_count'])
    print('Saved:', output)


if __name__ == '__main__':
    main()
