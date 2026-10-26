"""Conservative corrections for CoreWeave address cache. Only fills blank fields."""
from pathlib import Path
import shutil
import pandas as pd

KNOWN_LOCATIONS = {
    "coreweave c/o switch rno sector 1, sparks, nv 89437, usa": {
        "City": "Sparks",
        "State": "NV",
        "Region": "West Coast (US)",
    },
}


def auto_correct_addresses(path, *, backup=True):
    path = Path(path)
    df = pd.read_excel(path, dtype={'Case Number': str})
    required = {'Asset Location', 'City', 'State', 'Region'}
    if not required.issubset(df.columns):
        raise ValueError(f'Missing columns: {sorted(required - set(df.columns))}')

    def blank(series):
        return series.isna() | series.astype(str).str.strip().eq('')

    normalized = df['Asset Location'].fillna('').astype(str).str.strip().str.casefold()
    filled = {col: 0 for col in ('City', 'State', 'Region')}
    affected = set()
    for address, values in KNOWN_LOCATIONS.items():
        matched = normalized.eq(address)
        for col, value in values.items():
            mask = matched & blank(df[col])
            if mask.any():
                affected.update(df.index[mask].tolist())
                filled[col] += int(mask.sum())
                df.loc[mask, col] = value

    # Reuse region only when the SAME address has one unambiguous existing region.
    for address in KNOWN_LOCATIONS:
        matched = normalized.eq(address)
        regions = df.loc[matched & ~blank(df['Region']), 'Region'].astype(str).str.strip().unique()
        if len(regions) == 1:
            mask = matched & blank(df['Region'])
            if mask.any():
                affected.update(df.index[mask].tolist())
                filled['Region'] += int(mask.sum())
                df.loc[mask, 'Region'] = regions[0]

    if affected:
        if backup:
            backup_path = path.with_name(path.stem + '_before_autocorrect' + path.suffix)
            if backup_path.exists():
                from datetime import datetime
                backup_path = path.with_name(path.stem + '_before_autocorrect_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f') + path.suffix)
            shutil.copy2(path, backup_path)
        else:
            backup_path = None
        df.to_excel(path, index=False)
    else:
        backup_path = None
    return {'rows_updated': len(affected), 'fields_filled': filled, 'backup': backup_path}
