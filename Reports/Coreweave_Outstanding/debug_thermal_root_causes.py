"""Inspect Salesforce root-cause variants, especially thermal issues on H100 cases.

Usage:
    python debug_thermal_root_causes.py [path_to_outstanding_workbook.xlsx]

With no argument, loads the latest finalized report in
Reports/Coreweave_Outstanding/Output/, or the working report if none exists.
Prints a summary and writes thermal_root_cause_debug.txt beside the workbook.
Does not modify the workbook or history database.
"""
from __future__ import annotations

import argparse
from collections import Counter
from pathlib import Path
import re

import pandas as pd

ROOT_CAUSE = 'Root Cause: Root Cause Name'
PART_NUMBER = 'Faulty Part Number'
CASE_NUMBER = 'Case Number'
THERMAL_PATTERN = re.compile(r'therm|temp|overheat|heat|cool|fan|hot\s*spot|hotspot', re.I)
EXCLUDED_SHEETS = {'Dashboard', 'Comparison'}


def locate_workbook(explicit: Path | None) -> Path:
    if explicit is not None:
        return explicit.expanduser().resolve()
    here = Path(__file__).resolve().parent
    for base in [here, *here.parents]:
        output = base / 'Reports' / 'Coreweave_Outstanding' / 'Output'
        if output.is_dir():
            candidates = sorted(
                output.glob('CoreWeave Outstanding Cases *.xlsx'),
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
            if candidates:
                return candidates[0]
            working = output / 'outstanding_case_output.xlsx'
            if working.is_file():
                return working
    raise FileNotFoundError(
        'No report found automatically. Pass its path as an argument, e.g.\n'
        'python debug_thermal_root_causes.py "C:\\Reports\\outstanding_case_output.xlsx"'
    )


def filled(value: object) -> str:
    return '' if pd.isna(value) else str(value).strip()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', nargs='?', type=Path)
    args = parser.parse_args()
    workbook = locate_workbook(args.workbook)
    if not workbook.is_file():
        raise FileNotFoundError(workbook)

    sheets = pd.read_excel(workbook, sheet_name=None, dtype=str)
    frames = []
    for name, frame in sheets.items():
        frame.columns = frame.columns.astype(str).str.strip()
        if name in EXCLUDED_SHEETS or CASE_NUMBER not in frame:
            continue
        frame = frame.copy()
        frame['_Sheet'] = name
        frames.append(frame)
    if not frames:
        raise ValueError('No case worksheets with a Case Number column were found.')
    cases = pd.concat(frames, ignore_index=True)
    if ROOT_CAUSE not in cases:
        raise ValueError(f'Missing column: {ROOT_CAUSE}')
    if PART_NUMBER not in cases:
        raise ValueError(f'Missing column: {PART_NUMBER}')

    cases[CASE_NUMBER] = cases[CASE_NUMBER].map(filled)
    cases = cases[cases[CASE_NUMBER].ne('')].drop_duplicates(CASE_NUMBER, keep='first')
    part_text = cases[PART_NUMBER].fillna('').astype(str)
    is_h100 = part_text.str.contains('H100', case=False, regex=False)

    def summarize(df: pd.DataFrame, thermal_only: bool) -> Counter:
        counts: Counter = Counter()
        for value in df[ROOT_CAUSE].map(filled):
            # RMA enrichment can join multiple causes using " | ".
            causes = {part.strip() for part in value.split('|') if part.strip()}
            for cause in causes:
                if not thermal_only or THERMAL_PATTERN.search(cause):
                    counts[cause] += 1
        return counts

    lines = [f'Workbook: {workbook}', f'Unique cases: {len(cases):,}',
             f'H100 cases: {int(is_h100.sum()):,}', '']
    reports = [
        ('ALL THERMAL-LIKE ROOT CAUSES', summarize(cases, True)),
        ('H100 THERMAL-LIKE ROOT CAUSES', summarize(cases.loc[is_h100], True)),
        ('ALL ROOT CAUSE VARIANTS ON H100 CASES', summarize(cases.loc[is_h100], False)),
    ]
    for title, counts in reports:
        lines.extend([f'=== {title} ===', f'Unique variants: {len(counts)}'])
        if not counts:
            lines.append('(none found)')
        for cause, count in sorted(counts.items(), key=lambda x: (-x[1], x[0].casefold())):
            lines.append(f'{count:>6,} cases | {cause}')
        lines.append('')
    output = workbook.with_name('thermal_root_cause_debug.txt')
    result = '\n'.join(lines) + '\n'
    output.write_text(result, encoding='utf-8')
    print(result)
    print(f'Saved: {output}')


if __name__ == '__main__':
    main()
