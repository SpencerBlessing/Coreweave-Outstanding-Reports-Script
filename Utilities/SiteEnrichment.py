"""Conservative CoreWeave Site enrichment.

The report's Site column contains regional site codes (e.g. US-WEST-06),
not directory LoCodes (e.g. US-HIO03). Existing values are never overwritten.

Matching priority: explicit code in Subject; authoritative facility LoCode;
authoritative directory address; unambiguous facility/address evidence from
other report cases. Uncertain records remain blank.
"""
from __future__ import annotations

from collections import defaultdict, Counter
from pathlib import Path
import json
import re
import pandas as pd

try:
    from .AddressStandardizer import address_key, clean
except ImportError:
    from AddressStandardizer import address_key, clean

# Accept region codes such as US-WEST-06 / EU-SOUTH-04A, never US-HIO03.
SITE_CODE = re.compile(r"\b(?:US|EU|EMEA|APAC)-(?:WEST|EAST|NORTH|SOUTH|CENTRAL)-\d+[A-Z]?\b", re.I)
FACILITY_CODE = re.compile(r"\b[A-Z]{2}-[A-Z]{2,5}\d{2,4}\b", re.I)


def _site(value):
    value = clean(value).upper()
    return value if SITE_CODE.fullmatch(value) else ''


def _explicit_site(subject):
    codes = set(m.upper() for m in SITE_CODE.findall(clean(subject)))
    return next(iter(codes)) if len(codes) == 1 else ''


def _facilities(subject):
    """Extract LoCodes within any text, including rack paths like US-HIO03.DH110."""
    return {m.upper() for m in FACILITY_CODE.findall(clean(subject))}


def _unique_mapping(pairs):
    mapping = defaultdict(set)
    for key, value in pairs:
        if key and value:
            mapping[key].add(value)
    return {key: next(iter(values)) for key, values in mapping.items() if len(values) == 1}, {
        key for key, values in mapping.items() if len(values) > 1
    }


def _directory_maps(registry):
    if registry is None:
        return {}, {}, set(), set()
    loc_pairs, address_pairs = [], []
    for entry in registry.sites:
        site = _site(entry.region)
        if not site:
            continue
        loc_pairs.append((clean(entry.locode).upper(), site))
        key = address_key(entry.address)
        if key:
            address_pairs.append((key, site))
    loc_map, loc_conflicts = _unique_mapping(loc_pairs)
    addr_map, addr_conflicts = _unique_mapping(address_pairs)
    # Honor the standardizer's human-reviewed address variants.
    try:
        from .AddressStandardizer import ADDRESS_MAP
    except ImportError:
        from AddressStandardizer import ADDRESS_MAP
    aliases = []
    for raw, standardized in ADDRESS_MAP.items():
        canonical = address_key(standardized)
        if canonical in addr_map:
            aliases.append((address_key(raw), addr_map[canonical]))
    for raw, value in aliases:
        if raw and raw not in addr_conflicts:
            if raw in addr_map and addr_map[raw] != value:
                addr_conflicts.add(raw)
                addr_map.pop(raw, None)
            else:
                addr_map[raw] = value
    for key in addr_conflicts:
        addr_map.pop(key, None)
    return loc_map, addr_map, loc_conflicts, addr_conflicts


def enrich_sites(cases: pd.DataFrame, registry=None, status=None, cache_path: str | Path | None = None):
    """Fill *missing* Site values; preserve row order, columns and case count.

    Optional cache_path persists verified region-code pairs from prior exports.
    Contradictory mappings are quarantined rather than decided by frequency.
    No Salesforce inputs or existing Site values are modified in place.
    """
    report = cases.copy()
    status = status or (lambda message: None)
    if 'Site' not in report.columns:
        report['Site'] = ''
    if 'Subject' not in report.columns:
        report['Subject'] = ''
    if 'Asset Location' not in report.columns:
        report['Asset Location'] = ''

    loc_map, address_map, loc_conflicts, address_conflicts = _directory_maps(registry)
    rows = report.to_dict('records')
    observed_loc = []
    observed_addr = []
    # Historic observed evidence is retained as sets to detect conflicts, not votes.
    persisted = {'facility': {}, 'address': {}}
    path = Path(cache_path) if cache_path else None
    if path and path.is_file():
        try:
            loaded = json.loads(path.read_text(encoding='utf-8'))
            for kind in persisted:
                data = loaded.get(kind, {})
                if isinstance(data, dict):
                    persisted[kind] = {k: sorted({_site(s) for s in v if _site(s)})
                                       for k, v in data.items() if isinstance(v, list)}
        except (OSError, ValueError, TypeError) as error:
            status(f'Site cache unreadable, continuing without it: {error}')

    for row in rows:
        original = _site(row.get('Site'))
        explicit = _explicit_site(row.get('Subject'))
        # Learn only from explicitly coded subjects or valid existing Site fields.
        reliable = original or explicit
        if not reliable or (original and explicit and original != explicit):
            continue
        for code in _facilities(row.get('Subject')):
            observed_loc.append((code, reliable))
        address = address_key(row.get('Asset Location'))
        if address:
            observed_addr.append((address, reliable))

    for kind, values in persisted.items():
        for key, entries in values.items():
            for site in entries:
                (observed_loc if kind == 'facility' else observed_addr).append((key, site))

    learned_loc, learned_loc_conflicts = _unique_mapping(observed_loc)
    learned_addr, learned_addr_conflicts = _unique_mapping(observed_addr)
    # Never trust learned associations contrary to the directory.
    for key in list(learned_loc):
        if key in loc_conflicts or (key in loc_map and learned_loc[key] != loc_map[key]):
            learned_loc.pop(key, None)
            learned_loc_conflicts.add(key)
    for key in list(learned_addr):
        if key in address_conflicts or (key in address_map and learned_addr[key] != address_map[key]):
            learned_addr.pop(key, None)
            learned_addr_conflicts.add(key)

    counts = Counter()
    for idx, row in enumerate(rows):
        if clean(row.get('Site')):
            continue  # Keep original values, including unconventional codes.
        subject = row.get('Subject')
        explicit = _explicit_site(subject)
        codes = _facilities(subject)
        addr = address_key(row.get('Asset Location'))
        candidate_sources = {}
        if explicit:
            candidate_sources['subject'] = explicit
        authoritative_facilities = {loc_map[c] for c in codes if c in loc_map}
        if len(authoritative_facilities) == 1:
            candidate_sources['directory_facility'] = next(iter(authoritative_facilities))
        if addr in address_map:
            candidate_sources['directory_address'] = address_map[addr]
        learned_facilities = {learned_loc[c] for c in codes if c in learned_loc}
        if len(learned_facilities) == 1:
            candidate_sources['learned_facility'] = next(iter(learned_facilities))
        if addr in learned_addr:
            candidate_sources['learned_address'] = learned_addr[addr]
        # Any contradictory evidence means no fill, even when a source ranks higher.
        if len(set(candidate_sources.values())) > 1 or len(authoritative_facilities) > 1 or len(learned_facilities) > 1:
            counts['conflicting'] += 1
            continue
        if not candidate_sources:
            counts['unresolved'] += 1
            continue
        priority = ('subject', 'directory_facility', 'directory_address',
                    'learned_facility', 'learned_address')
        method = next(name for name in priority if name in candidate_sources)
        report.iat[idx, report.columns.get_loc('Site')] = candidate_sources[method]
        counts[method] += 1

    if path:
        # Store all evidence including conflicts so future runs don't silently
        # promote a previously conflicting association to a unique mapping.
        by_loc, by_addr = defaultdict(set), defaultdict(set)
        for key, value in observed_loc:
            if key and value:
                by_loc[key].add(value)
        for key, value in observed_addr:
            if key and value:
                by_addr[key].add(value)
        content = {'facility': {key: sorted(values) for key, values in sorted(by_loc.items())},
                   'address': {key: sorted(values) for key, values in sorted(by_addr.items())}}
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix('.tmp')
        tmp.write_text(json.dumps(content, indent=2), encoding='utf-8')
        tmp.replace(path)

    filled = sum(counts[name] for name in ('subject', 'directory_facility', 'directory_address',
                                            'learned_facility', 'learned_address'))
    status(f'Site enrichment: {filled} filled, {counts["conflicting"]} conflicting, '
           f'{counts["unresolved"]} unresolved. '
           f'By source: subject {counts["subject"]}, directory facility {counts["directory_facility"]}, '
           f'directory address {counts["directory_address"]}, learned facility {counts["learned_facility"]}, '
           f'learned address {counts["learned_address"]}.')
    return report
