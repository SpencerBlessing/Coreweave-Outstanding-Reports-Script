"""Normalize and group GPU faulty-part numbers by Salesforce case."""
import re
import pandas as pd

PART_MAP = {
    'AGPU-FRU-H100-80-SXM5-HS': 'GPU-FRU-H100-80-SXM5-HS',
    'GPU-FRU-H100-80-SXM5-H': 'GPU-FRU-H100-80-SXM5-HS',
    'AGPU-FRU-H100-80-SXM5-HST': 'GPU-FRU-H100-80-SXM5-HST',
    'GPU-FRU-H200-141-SXM5-H': 'GPU-FRU-H200-141-SXM5-HS',
    'AGPU-NVHGX-H100-88': 'GPU-NVHGX-H100-88',
}
PART_PATTERN = re.compile(r'(?<![A-Z0-9])(?:AGPU|GPU|BPN|CCP)-[A-Z0-9]+(?:-[A-Z0-9]+)*', re.I)


def standardize_part(value):
    if pd.isna(value):
        return value
    text = re.sub(r'\s+', '', str(value).strip().upper())
    return PART_MAP.get(text, text)


def extract_standardized_parts(value):
    if pd.isna(value) or not str(value).strip():
        return []
    text = str(value).upper()
    # Preserve legacy handling of 'GPU GPU-NVHGX-H100-88'.
    return list(dict.fromkeys(standardize_part(match) for match in PART_PATTERN.findall(text)))


def process_faulty_parts(df):
    required = {'Case Number', 'Faulty Part Number'}
    if not required.issubset(df.columns):
        raise ValueError(f'Faulty parts missing columns: {sorted(required - set(df.columns))}')
    data = df[['Case Number', 'Faulty Part Number']].copy()
    data['Faulty Part Number'] = data['Faulty Part Number'].apply(extract_standardized_parts)
    data = data.explode('Faulty Part Number').dropna(subset=['Case Number', 'Faulty Part Number'])
    data = data[data['Faulty Part Number'].str.contains('GPU', na=False)]
    return (data.groupby('Case Number', as_index=False)['Faulty Part Number']
            .agg(lambda values: ', '.join(dict.fromkeys(values))))
