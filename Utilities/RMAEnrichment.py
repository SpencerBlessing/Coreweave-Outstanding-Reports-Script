"""Enrich outstanding cases with RMA fields without adding or removing cases."""
from __future__ import annotations

import pandas as pd

from Utilities.CaseNumberStandardizer import standardize_case_number

# Salesforce relationship labels map to existing outstanding-case columns.
ALIASES = {
    'Case Owner: Full Name': 'Case Owner',
    'Asset: Product: Product Name': 'Product Name',
    'Asset: Serial Number': 'Serial Number',
    'Asset Location: Location Name': 'Asset Location',
    'Created Date': 'RMA Created Date',
}


def filled(value):
    if value is None or pd.isna(value):
        return False
    return str(value).strip().casefold() not in ('', 'nan', 'none', 'nat')


def distinct_values(values):
    """Aggregate RMA values; never split individual source values on delimiters."""
    result, seen = [], set()
    for value in values:
        if not filled(value):
            continue
        value = str(value).strip()
        key = value.casefold()
        if key not in seen:
            result.append(value)
            seen.add(key)
    return ' | '.join(result) if result else pd.NA


def enrich_outstanding_with_rma(outstanding: pd.DataFrame, rma: pd.DataFrame,
                                status=lambda message: None) -> pd.DataFrame:
    """Left join, fill blanks only, preserve original case rows and values."""
    if 'Case Number' not in outstanding or 'Case Number' not in rma:
        raise ValueError('Both reports require Case Number')

    rma = rma.copy()
    rma['Case Number'] = rma['Case Number'].apply(standardize_case_number)
    rma = rma.loc[rma['Case Number'].map(filled)].copy()
    rma = rma.rename(columns=ALIASES)

    # Combine Salesforce aliases and native columns if both are present.
    for name in rma.columns[rma.columns.duplicated()].unique():
        same = rma.loc[:, rma.columns == name]
        rma = rma.loc[:, rma.columns != name].copy()
        rma[name] = same.apply(lambda row: distinct_values(row.tolist()), axis=1)

    # Exclude non-outstanding cases before grouping to reduce work.
    keys = set(outstanding['Case Number'].dropna())
    rma = rma.loc[rma['Case Number'].isin(keys)].copy()
    if rma.empty:
        status('RMA enrichment: no matching outstanding cases.')
        return outstanding.copy()

    fields = [col for col in rma.columns if col != 'Case Number']
    grouped = rma.groupby('Case Number', sort=False, dropna=True)[fields].agg(
        lambda values: distinct_values(values.tolist())
    )

    result = outstanding.copy()
    matches = result['Case Number'].isin(grouped.index).sum()
    for field in fields:
        candidate = result['Case Number'].map(grouped[field])
        if field not in result.columns:
            result[field] = candidate
        else:
            missing = ~result[field].map(filled)
            result.loc[missing, field] = candidate.loc[missing].combine_first(result.loc[missing, field])

    if len(result) != len(outstanding):
        raise RuntimeError('RMA enrichment changed outstanding case count')
    status(f'RMA enrichment: {matches} outstanding rows matched; '
           f'{len(fields)} RMA fields considered.')
    return result
