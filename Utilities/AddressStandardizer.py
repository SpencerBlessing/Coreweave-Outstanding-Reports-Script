"""CoreWeave site directory and conservative address/site resolution.

Requires: pandas, openpyxl. The source workbook is never modified.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
import re
import unicodedata
import pandas as pd

US_TABS = ('US-Coreweave', 'US-Albatross', 'US-Phoenix', 'US-Merlin', 'US-Osprey')
STREET_WORDS = {
    'rd':'road', 'dr':'drive', 'st':'street', 'ave':'avenue', 'blvd':'boulevard',
    'pkwy':'parkway', 'cir':'circle', 'ln':'lane', 'hwy':'highway',
    'e':'east', 'w':'west', 'n':'north', 's':'south',
}
COUNTRIES = {'usa','us','united states','united states of america'}

# Historical manually reviewed address spellings; kept separate from directory authority.
ADDRESS_MAP = {'1 superloop cir, sparks, nv 89437-5623, usa': '1 Superloop Circle, Sparks, NV 89437-5623, USA',
 '1 superloop circle, sparks, nv 89437, usa': '1 Superloop Circle, Sparks, NV 89437-5623, USA',
 '1 superloop circle, sparks, nv 89437-5623, usa': '1 Superloop Circle, Sparks, NV 89437-5623, USA',
 '2 superloop cir, sparks, nv 89437-5623, usa': '2 Superloop Circle, Sparks, NV 89437-5623, USA',
 '2 superloop circle, sparks, nv 89437, usa': '2 Superloop Circle, Sparks, NV 89437-5623, USA',
 '2 superloop circle, sparks, nv 89437-5623, usa': '2 Superloop Circle, Sparks, NV 89437-5623, USA',
 '3 superloop cir, sparks, nv 89437, usa': '3 Superloop Circle, Sparks, NV 89437, USA',
 '3 superloop circle, sparks, nv 89437, usa': '3 Superloop Circle, Sparks, NV 89437, USA',
 'coreweave c/o switch rno s1, 2 superloop circle, sparks, nv 89437, usa': 'CoreWeave c/o Switch RNO '
                                                                           'Sector 1, 2 Superloop '
                                                                           'Circle, Sparks, NV 89437, '
                                                                           'USA',
 'coreweave c/o switch rno, s1 2 superloop circle, sparks, nv 89437, usa': 'CoreWeave c/o Switch RNO '
                                                                           'Sector 1, 2 Superloop '
                                                                           'Circle, Sparks, NV 89437, '
                                                                           'USA',
 'coreweave c/o switch rno sector 1, 2 superloop circle, sparks, nv 89437, usa': 'CoreWeave c/o Switch '
                                                                                 'RNO Sector 1, 2 '
                                                                                 'Superloop Circle, '
                                                                                 'Sparks, NV 89437, USA',
 'coreweave c/o switch rno sector 1, sparks, nv, us, 89437': 'CoreWeave c/o Switch RNO Sector 1, '
                                                             'Sparks, NV 89437, USA',
 'coreweave c/o switch rno sector 1, sparks, nv 89437, usa': 'CoreWeave c/o Switch RNO Sector 1, 2 '
                                                             'Superloop Circle, Sparks, NV 89437, USA',
 'switch rno sector 1, coreweave c/o, 2, superloop circle, sparks, nv 89437': 'CoreWeave c/o Switch RNO '
                                                                              'Sector 1, 2 Superloop '
                                                                              'Circle, Sparks, NV '
                                                                              '89437, USA',
 'coreweave c/o switch rno sector 2, sparks, nv, us, 89437': 'CoreWeave c/o Switch RNO Sector 2, '
                                                             'Sparks, NV 89437, USA',
 'coreweave c/o switch rno sector 2, sparks, nv 89437, usa': 'CoreWeave c/o Switch RNO Sector 2, '
                                                             'Sparks, NV 89437, USA',
 '100 digital crossroad drive, hammond, in, us, 46320': '100 Digital Crossroad Drive, Hammond, IN '
                                                        '46320-3070, USA',
 '1000 coit rd, plano, tx, us, 75075': '1000 Coit Road, Plano, TX 75075-5813, USA',
 '3500 e plano pkwy, plano, tx, us, 75074': '3500 East Plano Parkway, Plano, TX 75074-7205, USA',
 '120 satellite blvd, suwanee, ga, us, 30024': '120 Satellite Boulevard, Suwanee, GA 30024-2780, USA',
 '120 satellite boulevard northwest, suwanee, ga 30024-2780, usa': '120 Satellite Boulevard, Suwanee, '
                                                                   'GA 30024-2780, USA',
 '1401 meadowville technology pkwy, chester, va, us, 23836': '1401 Meadowville Technology Parkway, '
                                                             'Chester, VA 23836-2843, USA',
 '1510 lookout dr richardson tx 75082': '1510 Lookout Drive, Richardson, TX 75082, USA',
 '1510 lookout dr., richardson, tx, us, 75082': '1510 Lookout Drive, Richardson, TX 75082, USA',
 '155 palmer lane, marble, nc 28905': '155 Palmer Lane, Marble, NC 28905-8506, USA',
 '1650 e. higgins road, elk grove village, il, us, 60007': '1650 East Higgins Road, Elk Grove Village, '
                                                           'IL 60007-1606, USA',
 '1700 n river road, douglasville, ga, us, 30122': '1700 North River Road, Lithia Springs, GA '
                                                   '30122-4043, USA',
 '1919 park avenue, weehawken, nj 07086-6712, usa': '1919 Park Avenue, Weehawken, NJ 07086-6712, USA',
 '1919 park avenue, weehawken township, nj 07086-6712, usa': '1919 Park Avenue, Weehawken, NJ '
                                                             '07086-6712, USA',
 '200 webro road, parssippany, nj, us, 07054': '200 Webro Road, Parsippany, NJ 07054-2823, USA',
 '200 webro road, parsippany-troy hills, nj 07054-2823, usa': '200 Webro Road, Parsippany, NJ '
                                                              '07054-2823, USA',
 '229 west 7th street, cincinnati, oh, us, 45202': '229 West 7th Street, Cincinnati, OH 45202-2308, USA',
 '3 nutmeg ln, london, gb, e14 2ax': '3 Nutmeg Lane, London E14 2AX, UK',
 '3301 hibbetts road, austin, tx, us, 78725': '3301 Hibbetts Road, Austin, TX 78725-1005, USA',
 '375 riverside pkwy, lithia springs, ga, us, 30122': '375 Riverside Parkway, Lithia Springs, GA '
                                                      '30122-3891, USA',
 '4905 north point parkway, alpharetta, ga, us, 30009': '4905 North Point Parkway, Alpharetta, GA '
                                                        '30022-3781, USA',
 '5660 w. badura ave nap15, las vegas, nv, us, 89118': '5660 Badura Avenue, Las Vegas, NV 89118-4712, '
                                                       'USA',
 '7135 s decatur blvd, las vegas, nv, us, 89118': '7135 South Decatur Boulevard, Las Vegas, NV '
                                                  '89118-4376, USA',
 '9663 87th avenue southeast, ellendale, nd, us, 58436-4501': '9663 87th Avenue Southeast, Ellendale, '
                                                              'ND 58436-4501, USA',
 'calle 25, 6, 08040 barcelona, spain': 'Calle 25, nº 6, Parc Logístic de la Zona Franca, 08040 '
                                        'Barcelona, Spain',
 'carrer 25, nº6, parc logistic de la zona franca': 'Calle 25, nº 6, Parc Logístic de la Zona Franca, '
                                                    '08040 Barcelona, Spain',
 'carrer 25, nº6, parc logistic de la zona franca 08040, barcelona, spain': 'Calle 25, nº 6, Parc '
                                                                            'Logístic de la Zona '
                                                                            'Franca, 08040 Barcelona, '
                                                                            'Spain',
 'de la, calle, 25, 08040 barcelona, spain': 'Calle 25, nº 6, Parc Logístic de la Zona Franca, 08040 '
                                             'Barcelona, Spain',
 'de la, carrer 25, 6, 08040 barcelona, spain': 'Calle 25, nº 6, Parc Logístic de la Zona Franca, 08040 '
                                                'Barcelona, Spain',
 'número 25 6 carrer, barcelona, es, 08040': 'Calle 25, nº 6, Parc Logístic de la Zona Franca, 08040 '
                                             'Barcelona, Spain',
 'avenida ribera baja, 01213 rivabellosa, alava, spain': 'Avenida Ribera Baja, 01213 Ribabellosa, '
                                                         'Álava, Spain',
 'avenida ribera baja, 01213, alava, spain': 'Avenida Ribera Baja, 01213 Ribabellosa, Álava, Spain',
 'plataforma logistica arasur, avenida ribera baja, 01213 ribabellosa, alava, spain': 'Plataforma '
                                                                                      'Logística '
                                                                                      'ARASUR, Avenida '
                                                                                      'Ribera Baja, '
                                                                                      '01213 '
                                                                                      'Ribabellosa, '
                                                                                      'Álava, Spain',
 'plataforma logístíca arasur avenida ribera baja': 'Plataforma Logística ARASUR, Avenida Ribera Baja, '
                                                    '01213 Ribabellosa, Álava, Spain',
 'avenida ribera baja, 01213 ribabellosa, álava, spain': 'Plataforma Logística ARASUR, Avenida Ribera '
                                                         'Baja, 01213 Ribabellosa, Álava, Spain',
 'connect way manor royal, crawley, gb, rh10 9bd': 'Unit 2 Connect Way, Manor Royal, Crawley RH10 9BD, '
                                                   'UK',
 'digital realty, unit 2, crawley rh10 9bd, uk': 'Unit 2 Connect Way, Manor Royal, Crawley RH10 9BD, UK',
 'unit 2 connect way, crawley, uk': 'Unit 2 Connect Way, Manor Royal, Crawley RH10 9BD, UK',
 '99 carretera de la santa creu de calafell, sant boi de llobregat, es, 08830': 'Carretera de la Santa '
                                                                                'Creu de Calafell, 99, '
                                                                                '08840 Viladecans, '
                                                                                'Barcelona, Spain',
 '2000 galloping hill road, kenilworth, nj, us, 07033': '2000 Galloping Hill Road, Kenilworth, NJ '
                                                        '07033, USA',
 '2500 w union hills dr., phoenix, az, us, 85251': '2500 West Union Hills Drive, Phoenix, AZ 85251, USA'}

SITE_ADDRESS_ALIASES = {'7135 South Decatur Boulevard, Las Vegas, NV 89118-4376, USA': 'US-LAS01',
 '5660 Badura Avenue, Las Vegas, NV 89118-4712, USA': 'US-LAS03',
 '1510 Lookout Drive, Richardson, TX 75082, USA': 'US-RIN01',
 '1650 East Higgins Road, Elk Grove Village, IL 60007-1606, USA': 'US-EVI01',
 '229 West 7th Street, Cincinnati, OH 45202-2308, USA': 'US-CVG01',
 '200 Webro Road, Parsippany, NJ 07054-2823, USA': 'US-PPY01',
 '1401 Meadowville Technology Parkway, Chester, VA 23836-2843, USA': 'US-CSZ01'}


def clean(value) -> str:
    if value is None or pd.isna(value):
        return ''
    text = re.sub(r'\s+', ' ', str(value)).strip()
    return '' if text.casefold() in {'nan', 'none', 'null'} else text


def address_key(value) -> str:
    """Matching-only key. Does not modify source or published address."""
    s = unicodedata.normalize('NFKD', clean(value)).encode('ascii', 'ignore').decode().casefold()
    s = re.sub(r'\b([a-z]{2})\s+(\d{5})-\d{4}\b', r'\1 \2', s)
    s = re.sub(r'\b(?:usa|united states of america|united states)\b', '', s)
    s = re.sub(r'[^a-z0-9]+', ' ', s)
    words = s.split()
    words = [STREET_WORDS.get(w, w) for w in words]
    return ' '.join(words)


def standardize_us_zip(value):
    """Reduce US ZIP+4 to ZIP5 for display; leave other postal formats alone."""
    if not isinstance(value, str):
        return value
    return re.sub(r'\b([A-Z]{2})\s+(\d{5})-\d{4}\b', r'\1 \2', value, flags=re.I)


def standardize_address(value, registry=None, locode=None, site_name=None) -> str:
    """Legacy correction + verified site address, never overwrite a conflict."""
    original = clean(value)
    if not original:
        return original
    corrected = standardize_us_zip(ADDRESS_MAP.get(original.casefold(), original))
    if registry is None:
        return corrected
    result = registry.resolve(address=original, locode=locode, site_name=site_name)
    if result.status == 'matched' and result.site and result.site.address:
        # A site-name-only match does not verify a supplied street address.
        if result.method != 'site_name' or not original:
            return result.site.address
    return corrected if result.status != 'conflict' else original


def physical_address(value) -> str:
    """Preserve address as supplied, except whitespace. Never infer from shipping notes."""
    return clean(value)


@dataclass(frozen=True)
class Site:
    locode: str
    site_name: str
    region: str
    data_hall: str
    facility: str
    address: str
    source_sheet: str


@dataclass(frozen=True)
class Resolution:
    status: str  # matched, ambiguous, conflict, unmatched
    method: str
    site: Site | None = None
    candidates: tuple[str, ...] = ()


class SiteRegistry:
    def __init__(self, sites: list[Site], aliases: dict[str, str] | None = None):
        self.sites = sites
        self.by_locode = defaultdict(list)
        self.by_address = defaultdict(set)
        self.by_name = defaultdict(set)
        for site in sites:
            self.by_locode[site.locode].append(site)
            if site.address:
                self.by_address[address_key(site.address)].add(site.locode)
            if site.site_name:
                self.by_name[site.site_name.casefold()].add(site.locode)
        self.aliases = {}
        for raw, locode in {**SITE_ADDRESS_ALIASES, **(aliases or {})}.items():
            key = address_key(raw)
            code = clean(locode).upper()
            if code not in self.by_locode:
                raise ValueError(f'Alias references unknown LoCode: {code}')
            if key in self.aliases and self.aliases[key] != code:
                raise ValueError(f'Conflicting alias for {raw}')
            self.aliases[key] = code

    @classmethod
    def from_workbook(cls, path: str | Path, aliases: dict[str, str] | None = None):
        path = Path(path)
        sheets = pd.read_excel(path, sheet_name=list(US_TABS) + ['EMEA Sites'], dtype=str)
        sites = []
        seen = set()
        for sheet, df in sheets.items():
            df.columns = [clean(c) for c in df.columns]
            if sheet == 'EMEA Sites':
                # EMEA uses first column as LoCode; 'Site' is a region code.
                mappings = [('Unoffical site name', 'Location', 'Site', 'Data Hall', 'DC Facility', 'DC Address')]
            else:
                name_field = 'Es' if sheet == 'US-Osprey' else 'Unoffical site name'
                mappings = [('LoCode', name_field, 'Region name', 'Data Hall', 'DC Facility', 'DC Address')]
            for code_col, name_col, region_col, hall_col, facility_col, address_col in mappings:
                if not all(col in df for col in (code_col, name_col, region_col, hall_col, facility_col, address_col)):
                    raise ValueError(f'Unexpected columns in {sheet}')
                for _, row in df.iterrows():
                    code = clean(row[code_col]).upper()
                    address = physical_address(row[address_col])
                    if not re.fullmatch(r'[A-Z]{2}-[A-Z0-9]+', code) or not address:
                        continue
                    record = Site(code, clean(row[name_col]), clean(row[region_col]),
                                  clean(row[hall_col]), clean(row[facility_col]), address, sheet)
                    # Preserve different data halls and source entries; drop identical rows only.
                    signature = (record.locode, record.data_hall, address_key(record.address), sheet)
                    if signature not in seen:
                        sites.append(record)
                        seen.add(signature)
        return cls(sites, aliases)

    def resolve(self, address=None, locode=None, site_name=None) -> Resolution:
        key = address_key(address)
        code = clean(locode).upper()
        name = clean(site_name).casefold()
        address_matches = set(self.by_address.get(key, ())) if key else set()
        if key in self.aliases:
            address_matches.add(self.aliases[key])
        # Existing manually standardized addresses can bridge directory spellings.
        corrected = ADDRESS_MAP.get(clean(address).casefold())
        if corrected:
            corrected_key = address_key(corrected)
            address_matches.update(self.by_address.get(corrected_key, ()))
            if corrected_key in self.aliases:
                address_matches.add(self.aliases[corrected_key])
        code_matches = {code} if code in self.by_locode else set()
        name_matches = set(self.by_name.get(name, ())) if name else set()
        if code_matches and address_matches and not (code_matches & address_matches):
            return Resolution('conflict', 'locode_vs_address', candidates=tuple(sorted(code_matches | address_matches)))
        if code_matches and name_matches and not (code_matches & name_matches):
            return Resolution('conflict', 'locode_vs_site_name', candidates=tuple(sorted(code_matches | name_matches)))
        if address_matches and name_matches and not (address_matches & name_matches):
            return Resolution('conflict', 'address_vs_site_name', candidates=tuple(sorted(address_matches | name_matches)))
        if code_matches:
            matches, method = code_matches, 'locode'
        elif address_matches:
            matches, method = address_matches, 'address'
        elif name_matches:
            matches, method = name_matches, 'site_name'
        else:
            return Resolution('unmatched', 'none')
        if len(matches) != 1:
            return Resolution('ambiguous', method, candidates=tuple(sorted(matches)))
        matched_code = next(iter(matches))
        records = self.by_locode[matched_code]
        # One LoCode can have multiple halls or addresses. Never silently choose one.
        if len({address_key(s.address) for s in records}) > 1:
            narrowed = [s for s in records if address_key(s.address) == key]
            if narrowed:
                return Resolution('matched', method, narrowed[0], (matched_code,))
            return Resolution('ambiguous', 'multiple_addresses_for_locode', candidates=(matched_code,))
        return Resolution('matched', method, records[0], (matched_code,))

    def enrich(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add fields without overwriting source columns or silently fixing conflicts."""
        results = []
        for row in df.to_dict('records'):
            result = self.resolve(row.get('Asset Location'), row.get('LoCode'), row.get('Site'))
            site = result.site
            results.append({
                'Match Status': result.status,
                'Match Method': result.method,
                'Matched LoCode': site.locode if site else '',
                'Matched Site': site.site_name if site else '',
                'Standardized Address': standardize_address(row.get('Asset Location'), self, row.get('LoCode'), row.get('Site')),
                'Directory Region': site.region if site else '',
                'Directory Facility': site.facility if site else '',
                'Match Candidates': ', '.join(result.candidates),
            })
        return pd.concat([df.reset_index(drop=True), pd.DataFrame(results)], axis=1)


def get_site_reference(address, registry=None, locode=None, site_name=None):
    """Compatibility helper; return site metadata for confident matches only."""
    if registry is None:
        return None
    result = registry.resolve(address=address, locode=locode, site_name=site_name)
    if result.status != 'matched' or result.site is None:
        return None
    site = result.site
    return {'site_name': site.site_name, 'locode': site.locode, 'region': site.region,
            'data_hall': site.data_hall, 'facility': site.facility, 'address': site.address}


def extract_site(subject):
    """Extract site identifier from Salesforce case subject."""
    subject = clean(subject)
    if not subject:
        return ''
    parts = [part.strip() for part in subject.split('|')]
    if len(parts) >= 2 and re.fullmatch(r'[A-Z]{3,}\d+|[A-Z]+-[A-Z]+-?\d+[A-Z]?', parts[1], flags=re.I):
        return parts[1].upper()
    match = re.search(r'\bUS-[A-Z]+-?\d+[A-Z]?\b', subject, flags=re.I)
    return match.group(0).upper() if match else ''


def run_report(report_path, directory_path, output_path=None):
    """Enrich all report worksheets containing Asset Location, preserving source file."""
    report_path = Path(report_path)
    output_path = Path(output_path) if output_path else report_path.with_name(report_path.stem + '_Standardized.xlsx')
    registry = SiteRegistry.from_workbook(directory_path)
    sheets = pd.read_excel(report_path, sheet_name=None, dtype={'Case Number': str})
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for name, frame in sheets.items():
            if 'Asset Location' in frame.columns:
                frame = registry.enrich(frame)
            frame.to_excel(writer, sheet_name=name, index=False)
    return output_path


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Standardize CoreWeave addresses and map sites')
    parser.add_argument('report', help='Input report XLSX')
    parser.add_argument('directory', help='CoreWeave site directory XLSX')
    parser.add_argument('--output', help='Output XLSX (default: *_Standardized.xlsx)')
    args = parser.parse_args()
    print(run_report(args.report, args.directory, args.output))
