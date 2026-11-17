"""CoreWeave address standardization and enrichment.

CoreWeave_Map.json is the ONLY authority for canonical addresses and Site data.
This module never learns from case history or modifies the map. An unknown or
conflicting address is preserved as received until the map is reviewed manually.
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass
from functools import lru_cache
from collections import defaultdict
import json
from pathlib import Path
import re
import unicodedata

import pandas as pd


DEFAULT_MAP = Path(__file__).resolve().parents[1] / 'Address Cache' / 'CoreWeave_Map.json'

# Keep the same key rules used when CoreWeave_Map.json was built.
STREET_WORDS = {
    'rd': 'road', 'dr': 'drive', 'st': 'street', 'ave': 'avenue',
    'blvd': 'boulevard', 'pkwy': 'parkway', 'cir': 'circle',
    'ln': 'lane', 'hwy': 'highway', 'wy': 'way',
    'e': 'east', 'w': 'west', 'n': 'north', 's': 'south',
}


def clean(value) -> str:
    if value is None or pd.isna(value):
        return ''
    text = re.sub(r'\s+', ' ', str(value)).strip()
    return '' if text.casefold() in {'nan', 'none', 'null'} else text


def address_key(value) -> str:
    """Matching key only; never a replacement for a canonical map address."""
    text = unicodedata.normalize('NFKD', clean(value)).encode(
        'ascii', 'ignore'
    ).decode().casefold()
    text = re.sub(r'\b([a-z]{2})\s+(\d{5})-\d{4}\b', r'\1 \2', text)
    text = re.sub(r'\b(?:usa|united states of america|united states)\b', '', text)
    words = re.sub(r'[^a-z0-9]+', ' ', text).split()
    return ' '.join(STREET_WORDS.get(word, word) for word in words)


# These rules affect comparison only. They never change published map records.
DIRECTIONS = {
    'n': 'north', 's': 'south', 'e': 'east', 'w': 'west',
    'ne': 'northeast', 'nw': 'northwest', 'se': 'southeast', 'sw': 'southwest',
}
STREET_TYPES = {
    'road', 'street', 'drive', 'avenue', 'boulevard', 'parkway',
    'circle', 'lane', 'highway', 'way', 'place', 'court', 'terrace',
    'trail', 'loop',
}
DIRECTION_WORDS = set(DIRECTIONS.values())
US_STATE_CODES = {
    'al', 'ak', 'az', 'ar', 'ca', 'co', 'ct', 'de', 'fl', 'ga', 'hi',
    'id', 'il', 'in', 'ia', 'ks', 'ky', 'la', 'me', 'md', 'ma', 'mi',
    'mn', 'ms', 'mo', 'mt', 'ne', 'nv', 'nh', 'nj', 'nm', 'ny', 'nc',
    'nd', 'oh', 'ok', 'or', 'pa', 'ri', 'sc', 'sd', 'tn', 'tx', 'ut',
    'vt', 'va', 'wa', 'wv', 'wi', 'wy', 'dc',
}


def _postal_code(text: str) -> str:
    """Extract a postal code for comparison, without altering the source."""
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode().upper()
    # Preserve Swedish 3+2 postal groups (e.g. 791 77).
    if re.search(r'\b(?:SWEDEN|SVERIGE|SE)\b', text):
        match = re.search(r'\b\d{3}\s+\d{2}\b', text)
        if match:
            return re.sub(r'\D', '', match.group())
    # Prefer explicit British or Canadian alphanumeric codes over digit-only ZIPs.
    for pattern in (
        r'\b[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}\b',
        r'\b[A-Z]\d[A-Z]\s*\d[A-Z]\d\b',
        r'\b\d{5}(?:-\d{4})?\b',
    ):
        match = re.search(pattern, text)
        if match:
            return re.sub(r'\W', '', match.group()).upper()[:5] if match.group()[0].isdigit() else re.sub(r'\W', '', match.group()).upper()
    return ''


def _address_profile(text: str):
    """Physical street + location identity, deliberately not fuzzy/geocoded.

    Restrict indexing to the primary address; a parenthetical shipping note
    must NEVER become an alternative address (e.g. Coit Rd vs Jomar Dr).
    """
    original = clean(text)
    primary = re.split(r'\s*\(|\s+-\s+(?=coreweave|shipping|ship\b)', original, maxsplit=1, flags=re.I)[0]
    words = address_key(primary).split()
    words = [DIRECTIONS.get(word, word) for word in words]
    # Written-out street ordinals and their numeric equivalents denote the
    # same street (e.g. West Seventh St == W 7th St). Matching only.
    ordinal_names = {
        'first': '1st', 'second': '2nd', 'third': '3rd',
        'fourth': '4th', 'fifth': '5th', 'sixth': '6th',
        'seventh': '7th', 'eighth': '8th', 'ninth': '9th',
        'tenth': '10th',
    }
    words = [ordinal_names.get(word, word) for word in words]
    # Comparison-only spelling correction. Published names always come from the map.
    words = [{'parssippany': 'parsippany'}.get(word, word) for word in words]
    # A valid street number must be followed by a street name and type, without
    # another building number in between ("Sector 1, 2 Superloop Circle").
    start = end = None
    for i, word in enumerate(words):
        if not re.fullmatch(r'\d{1,5}', word):
            continue
        for j in range(i + 2, min(i + 8, len(words))):
            if words[j] in STREET_TYPES and not any(w.isdigit() for w in words[i + 1:j]):
                start, end = i, j
                break
        if start is not None:
            break
    if start is None:
        # Spanish numbered streets: "Calle 25 nº6", "Carrer 25, 6" and
        # "número 25 6 carrer" are equivalent. Street number and building
        # number MUST remain ordered: Calle 6 nº25 is a different street.
        if any(w in {'calle', 'carrer'} for w in words):
            expanded = []
            for word in words:
                match = re.fullmatch(r'(?:n|no|num|numero)(\d+)', word)
                expanded.extend(['numero', match.group(1)] if match else [word])
            postal = _postal_code(primary)
            # Don't mistake the postal code for a building number.
            nums = [w for w in expanded if w.isdigit() and w != postal]
            locality = frozenset(w for w in expanded if w in {'barcelona', 'madrid'})
            if len(nums) in (1, 2) and postal and locality:
                return {
                    'number': '/'.join(nums) if len(nums) == 2 else nums[0] + '/?',
                    'stem': ('spanish_street',), 'direction': (),
                    'postal': postal, 'locality': locality, 'unit': '',
                    'partial_house': len(nums) == 1,
                }
        for i, word in enumerate(words):
            if word in {'avenida', 'av', 'avda'}:
                street = words[i + 1:i + 3]
                if len(street) == 2 and all(w.isalpha() for w in street):
                    postal = _postal_code(primary)
                    tail = [w for w in words[i + 3:] if w not in {
                        'spain', 'espana', 'plataforma', 'logistica', 'arasur'
                    } and not w.isdigit()]
                    return {
                        'number': '', 'stem': ('avenida', *street),
                        'direction': (), 'postal': postal,
                        'locality': frozenset(tail), 'unit': '',
                    }
        # Swedish street names can appear before or after the house number.
        # Match only with a Swedish postal format and country indicator.
        if (re.search(r'\b(?:sweden|sverige|se)\b', primary, re.I)
                and re.search(r'\b\d{3}\s+\d{2}\b', primary)):
            for i, word in enumerate(words):
                if not (len(word) >= 7 and word.endswith(('vagen', 'gatan', 'grand'))):
                    continue
                adjacent = (i + 1 if i + 1 < len(words) and words[i + 1].isdigit()
                            else i - 1 if i > 0 and words[i - 1].isdigit() else None)
                if adjacent is None:
                    continue
                return {
                    'number': words[adjacent], 'stem': ('swedish_street', word),
                    'direction': (), 'postal': _postal_code(primary),
                    'locality': frozenset(w for j, w in enumerate(words)
                                          if j not in (i, adjacent) and w.isalpha()
                                          and w not in {'sweden', 'sverige', 'se'}),
                    'unit': '',
                }
        return None

    street_words = words[start + 1:end]
    stem = tuple(w for w in street_words if w not in DIRECTION_WORDS)
    # Do not erase the street name or its type, even if abbreviated.
    if not stem:
        return None
    # The direction can be omitted by Salesforce, but conflicting directions
    # are never interchangeable.
    direction = tuple(w for w in street_words if w in DIRECTION_WORDS)
    trailing = words[end + 1:]
    # 'Avenue SE' and 'Drive NW' are part of the street, not the city.
    if trailing and trailing[0] in DIRECTION_WORDS:
        direction = tuple(sorted(set(direction + (trailing[0],))))
        trailing = trailing[1:]
    post = _postal_code(primary)
    explicit_subunit = trailing.pop(0) if trailing and trailing[0].isdigit() and trailing[0] != post else ''
    trailing = [w for w in trailing if w not in US_STATE_CODES and w not in {
        'usa', 'us', 'united', 'states', 'kingdom', 'uk', 'gb', 'canada', 'spain',
        'espana', 'sweden', 'norway',
    } and not re.fullmatch(r'\d{5}|\d{9}', w)]
    # A postal fragment like E14/2AX should not be considered a locality.
    trailing = [w for w in trailing if not any(ch.isdigit() for ch in w)]
    # Guard against silently merging subunits at the same street address.
    unit = re.search(r'(?:\b(?:suite|ste|unit|floor|fl|building|bldg|hall)\s*|#\s*)([a-z0-9-]+)', primary, re.I)
    unit_value = unit.group(1).casefold() if unit and unit.start() > original.casefold().find(str(words[start])) else explicit_subunit
    markers = frozenset(re.findall(r'\b(?:nap|hall|sector|phase)\s*-?\s*\d+\b', primary.casefold()))
    return {
        'number': words[start], 'stem': stem, 'direction': direction,
        'postal': post, 'locality': frozenset(trailing), 'unit': unit_value,
        'markers': markers,
    }


def _compatible_profiles(incoming: dict, mapped: dict) -> bool:
    if incoming.get('partial_house') and incoming['stem'] == ('spanish_street',):
        # Omitted building number is acceptable only when postal AND city
        # agree and the map yields exactly one possible address later.
        if not mapped['number'].startswith(incoming['number'][:-1]):
            return False
        if not (incoming['postal'] and incoming['postal'] == mapped['postal']
                and incoming['locality'] & mapped['locality']):
            return False
    elif (incoming['number'], incoming['stem']) != (mapped['number'], mapped['stem']):
        return False
    if incoming['direction'] and mapped['direction'] and incoming['direction'] != mapped['direction']:
        return False
    if incoming['unit'] != mapped['unit']:
        return False
    if incoming.get('markers', frozenset()) != mapped.get('markers', frozenset()):
        return False
    # A different ZIP is not a formatting discrepancy and must not be guessed.
    if incoming['postal'] and mapped['postal'] and incoming['postal'] != mapped['postal']:
        return False
    left, right = incoming['locality'], mapped['locality']
    if left and right and not (left & right):
        return False
    # NAP identifiers and similar hall tags are physical differentiators.
    hall_pattern = re.compile(r'^(?:nap|hall|sector|phase)\w*\d+$')
    halls_a = {w for w in left if hall_pattern.fullmatch(w)}
    halls_b = {w for w in right if hall_pattern.fullmatch(w)}
    if halls_a != halls_b:
        return False
    # Require a shared postcode or a location anchor, not just the street name.
    if incoming['postal'] and mapped['postal']:
        return True
    return bool(left and right and left & right)




def _verified_variant_match(incoming: dict, mapped: dict, original: str, record: dict) -> bool:
    """Permit only verified historical representations of a map location.

    These are comparison rules, not alternate address authorities. In every
    case the complete canonical address and all metadata come from the map.
    The normal unique-candidate and conflict checks still apply.
    """
    if (incoming['number'], incoming['stem']) != (mapped['number'], mapped['stem']):
        return False
    if incoming['direction'] != mapped['direction'] or incoming['unit'] != mapped['unit']:
        return False

    source_words = set(address_key(original).split())
    canonical_city = address_key(record.get('city'))
    canonical_state = clean(record.get('state')).casefold()
    house, stem = incoming['number'], incoming['stem']
    old_post, new_post = incoming['postal'], mapped['postal']
    old_city = incoming['locality']
    markers = incoming.get('markers', frozenset())
    mapped_markers = mapped.get('markers', frozenset())

    # Both ZIPs are documented for this street; map's 30022 wins.
    if (house == '4905' and stem == ('point',)
            and old_post == '30009' and new_post == '30022'
            and canonical_city == 'alpharetta' and canonical_state == 'ga'
            and 'ga' in source_words and 'alpharetta' in old_city
            and markers == mapped_markers):
        return True

    # Douglasville is the historic municipality for the same 1700 N River Rd.
    if (house == '1700' and stem == ('river',)
            and old_post == new_post == '30122'
            and canonical_city == 'lithia springs' and canonical_state == 'ga'
            and 'ga' in source_words and 'douglasville' in old_city
            and markers == mapped_markers):
        return True

    # Old Grand Rapids mailing locality/ZIP, same 6100 E Paris Ave SE.
    if (house == '6100' and stem == ('paris',)
            and old_post == '49512' and new_post == '49316'
            and canonical_city == 'caledonia' and canonical_state == 'mi'
            and 'mi' in source_words and {'grand', 'rapids'} <= old_city
            and markers == mapped_markers):
        return True

    # NAP15 is a verified facility identifier at 5660 W Badura, not a
    # different house number. The map remains authoritative for the Site.
    if (house == '5660' and stem == ('badura',)
            and old_post == new_post == '89118'
            and canonical_city == 'las vegas' and canonical_state == 'nv'
            and 'nv' in source_words and {'las', 'vegas'} <= old_city
            and markers == frozenset({'nap15'}) and not mapped_markers):
        return True

    # West Sussex is a broader administrative area for the verified Unit 2
    # Connect Way location. Require explicit unit, street, and UK country;
    # an omitted postcode alone is NOT adequate for any other address.
    if (house == '2' and stem == ('connect',)
            and not old_post and new_post == 'RH109BD'
            and canonical_city == 'crawley'
            and canonical_state in {'united kingdom', 'uk', 'gb'}
            and re.search(r'\bunit\s*2\b', original, re.I)
            and 'sussex' in old_city
            and source_words & {'uk', 'gb'}
            and markers == mapped_markers):
        return True

    # The ARASUR campus can be identified by its distinctive name and the
    # full Avenida Ribera Baja street, even when the postcode is omitted.
    if (not house and stem == ('avenida', 'ribera', 'baja')
            and not old_post and new_post == '01213'
            and canonical_city == 'rivabellosa' and canonical_state == 'spain'
            and {'plataforma', 'logistica', 'arasur'} <= source_words
            and markers == mapped_markers):
        return True

    return False


def _verified_incomplete_street(text: str, by_address: dict, conflicts: dict):
    """Resolve a uniquely identified Crawley street when the unit is omitted.

    The incoming address must contain the estate, city AND full postcode.
    Never choose between two units/buildings at the same street and postcode.
    """
    words = set(address_key(text).split())
    if not {'connect', 'way', 'manor', 'royal', 'crawley'} <= words:
        return None, 'unmatched'
    if _postal_code(text) != 'RH109BD':
        return None, 'unmatched'
    if re.search(r'\b(?:unit|suite|building|bldg|#)\s*\w+', text, re.I):
        return None, 'unmatched'
    if re.search(r'\b\d{1,5}\b', re.sub(r'\bRH10\s*9BD\b', '', text, flags=re.I)):
        return None, 'unmatched'

    hits = {}
    for record in by_address.values():
        if not isinstance(record, dict):
            continue
        if (address_key(record.get('city')) != 'crawley'
                or clean(record.get('state')).casefold() not in {'united kingdom', 'uk', 'gb'}):
            continue
        canonical = clean(record.get('address'))
        if _postal_code(canonical) != 'RH109BD':
            continue
        profile = _address_profile(canonical)
        if not profile or profile['stem'] != ('connect',):
            continue
        if not {'manor', 'royal'} <= profile['locality']:
            continue
        key = address_key(canonical)
        if key in conflicts:
            return None, 'conflict'
        hits[(key, clean(record.get('site')), clean(record.get('locode')))] = record
    if len(hits) == 1:
        return next(iter(hits.values())), 'matched'
    return None, 'conflict' if len(hits) > 1 else 'unmatched'


def standardize_us_zip(value):
    """Formatting helper only. This does not identify or verify an address."""
    if not isinstance(value, str):
        return value
    return re.sub(r'\b([A-Z]{2})\s+(\d{5})-\d{4}\b', r'\1 \2', value, flags=re.I)


def physical_address(value) -> str:
    """Preserve physical/shipping address text except for extra whitespace."""
    return clean(value)


@dataclass(frozen=True)
class Site:
    locode: str
    site_name: str
    region: str                 # Geographic reporting region (not the Site code).
    data_hall: str
    facility: str
    address: str
    source_sheet: str
    site_code: str = ''          # CoreWeave Site identifier, e.g. US-WEST-09.
    city: str = ''
    state: str = ''
    customer: str = ''

    @property
    def reporting_site(self) -> str:
        """Use the verified LoCode only when the map has no Site code."""
        return self.site_code or self.locode

    @property
    def reporting_site_source(self) -> str:
        if self.site_code:
            return 'Site'
        return 'LoCode fallback' if self.locode else ''


@dataclass(frozen=True)
class Resolution:
    status: str                 # matched, conflict, unmatched
    method: str                 # address, locode, address_vs_locode, or none
    site: Site | None = None
    candidates: tuple[str, ...] = ()


def _site_from_record(record: dict) -> Site:
    return Site(
        locode=clean(record.get('locode')).upper(),
        site_name=clean(record.get('site_name')),
        region=clean(record.get('region')),
        data_hall=clean(record.get('data_hall')),
        facility=clean(record.get('facility')),
        address=clean(record.get('address')),
        source_sheet=clean(record.get('source')),
        site_code=clean(record.get('site')),
        city=clean(record.get('city')),
        state=clean(record.get('state')),
        customer=clean(record.get('customer')),
    )


class SiteRegistry:
    """Read-only view of the single CoreWeave map, with no separate lookup data."""

    def __init__(self, map_path: str | Path | None = None):
        self.map_path = Path(map_path) if map_path is not None else DEFAULT_MAP
        if not self.map_path.is_file():
            raise FileNotFoundError(f'CoreWeave map not found: {self.map_path}')

        with self.map_path.open('r', encoding='utf-8') as handle:
            data = json.load(handle)
        if not isinstance(data, dict) or not all(
            isinstance(data.get(name), dict)
            for name in ('by_address', 'by_locode', 'conflicts')
        ):
            raise ValueError('CoreWeave map must contain by_address, by_locode and conflicts objects')

        self.by_address = data['by_address']
        self.by_locode = data['by_locode']
        self.address_conflicts = data['conflicts'].get('address', {})
        self.locode_conflicts = data['conflicts'].get('locode', {})
        if not isinstance(self.address_conflicts, dict) or not isinstance(self.locode_conflicts, dict):
            raise ValueError('CoreWeave map conflicts must contain address/locode dictionaries')
        self.sites = [_site_from_record(record) for record in data.get('records', [])]

        # Recognize all explicit map conflicts, even keys omitted from by_address.
        self._conflict_profiles = [_address_profile(key) for key in self.address_conflicts]
        self._conflict_profiles = [p for p in self._conflict_profiles if p is not None]

        # Ephemeral index built from the SAME map, not a second address cache.
        # Include existing confirmed by_address aliases as searchable spellings.
        self._street_index = defaultdict(list)
        seen = set()
        for key, record in self.by_address.items():
            if not isinstance(record, dict):
                raise ValueError(f'Invalid CoreWeave map address entry: {key}')
            for candidate_text in (key, record.get('address', '')):
                profile = _address_profile(candidate_text)
                if profile is None:
                    continue
                identity = (key, candidate_text)
                if identity in seen:
                    continue
                seen.add(identity)
                self._street_index[(profile['number'], profile['stem'])].append((profile, record))
                if (profile['stem'] == ('spanish_street',)
                        and '/' in profile['number'] and not profile.get('partial_house')):
                    street_number = profile['number'].split('/', 1)[0]
                    self._street_index[(street_number + '/?', profile['stem'])].append((profile, record))

    @classmethod
    def from_map(cls, path: str | Path | None = None) -> SiteRegistry:
        return cls(path)

    def _unique_facility_postcode_unit(self, text: str):
        """Resolve an incomplete street address only with multiple independent anchors.

        Example: Digital Realty, Unit 2, Crawley RH10 9BD can identify a
        canonical Unit 2 address from the map when the facility, postal code,
        unit number, and city all agree uniquely. A facility name by itself
        never establishes an address.
        """
        postal = _postal_code(text)
        unit_match = re.search(r'\bunit\s*#?\s*([a-z0-9-]+)\b', text, re.I)
        if not postal or unit_match is None:
            return None, 'unmatched'
        unit = unit_match.group(1).casefold()
        incoming_words = set(address_key(text).split())
        hits = {}
        for record in self.by_address.values():
            canonical = clean(record.get('address'))
            if _postal_code(canonical) != postal:
                continue
            mapped_unit = re.search(r'\bunit\s*#?\s*([a-z0-9-]+)\b', canonical, re.I)
            if mapped_unit is None or mapped_unit.group(1).casefold() != unit:
                continue
            city = clean(record.get('city'))
            if not city or not set(address_key(city).split()).issubset(incoming_words):
                continue
            facility_words = [word for word in address_key(record.get('facility')).split()
                              if len(word) >= 4 and word not in {
                                  'datacenter', 'centre', 'center', 'facility', 'building',
                              }]
            if len(facility_words) < 2 or not set(facility_words).issubset(incoming_words):
                continue
            canonical_key = address_key(canonical)
            if canonical_key in self.address_conflicts:
                return None, 'conflict'
            hits[(canonical_key, clean(record.get('site')), clean(record.get('locode')))] = record
        if len(hits) == 1:
            return next(iter(hits.values())), 'matched'
        return None, 'conflict' if len(hits) > 1 else 'unmatched'

    @lru_cache(maxsize=8192)
    def _flexible_lookup(self, text: str):
        """Return (record, status) only when one map identity is supported."""
        profile = _address_profile(text)
        if profile is None:
            record, status = self._unique_facility_postcode_unit(text)
            if status != 'unmatched':
                return record, status
            return _verified_incomplete_street(text, self.by_address, self.address_conflicts)
        # Known ambiguous addresses must remain ambiguous when prefixes are omitted.
        if any(_compatible_profiles(profile, bad) for bad in self._conflict_profiles):
            return None, 'conflict'
        hits = {}
        for candidate, record in self._street_index.get((profile['number'], profile['stem']), []):
            if not (_compatible_profiles(profile, candidate)
                    or _verified_variant_match(profile, candidate, text, record)):
                continue
            canonical_key = address_key(record.get('address'))
            # Preserve known contradictions even when Salesforce omits prefixes.
            if canonical_key in self.address_conflicts:
                return None, 'conflict'
            # Identical LoCode with different physical locations should not be
            # collapsed, nor should conflicting map records be arbitrarily chosen.
            hits[(canonical_key, clean(record.get('site')), clean(record.get('locode')))] = record
        if len(hits) == 1:
            return next(iter(hits.values())), 'matched'
        return None, 'conflict' if len(hits) > 1 else 'unmatched'

    def resolve(self, address=None, locode=None, site_name=None) -> Resolution:
        """Resolve the address first; a LoCode cannot certify an unknown address.

        A unique LoCode can supply an address only when the address is empty.
        `site_name` is retained for caller compatibility but is not used for
        guessing a street address or Site assignment.
        """
        raw_address = clean(address)
        key = address_key(raw_address)
        code = clean(locode).upper()

        if key:
            if key in self.address_conflicts:
                choices = self.address_conflicts[key].get('sites', [])
                return Resolution('conflict', 'address', candidates=tuple(choices))

            record = self.by_address.get(key)
            method = 'address'
            if record is None:
                # Use normalization and unique physical address identity, not
                # invented aliases or similarity guesses, to find map records.
                record, status = self._flexible_lookup(raw_address)
                if status != 'matched':
                    return Resolution(status, 'normalized_address')
                method = 'normalized_address'
            if not isinstance(record, dict):
                raise ValueError(f'Invalid CoreWeave map address entry: {key}')

            mapped_code = clean(record.get('locode')).upper()
            if code and mapped_code and code != mapped_code:
                return Resolution(
                    'conflict', 'address_vs_locode',
                    candidates=tuple(sorted({code, mapped_code})),
                )
            return Resolution('matched', method, _site_from_record(record), (mapped_code,))

        if not code:
            return Resolution('unmatched', 'none')
        if code in self.locode_conflicts:
            choices = self.locode_conflicts[code].get('sites', [])
            return Resolution('conflict', 'locode', candidates=tuple(choices))
        record = self.by_locode.get(code)
        if record is None:
            return Resolution('unmatched', 'locode')
        if not isinstance(record, dict):
            raise ValueError(f'Invalid CoreWeave map LoCode entry: {code}')
        return Resolution('matched', 'locode', _site_from_record(record), (code,))

    def enrich(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add map-derived fields without overwriting the incoming source fields."""
        records = []
        for row in df.to_dict('records'):
            result = self.resolve(row.get('Asset Location'), row.get('LoCode'), row.get('Site'))
            site = result.site
            records.append({
                'Match Status': result.status,
                'Match Method': result.method,
                'Matched LoCode': site.locode if site else '',
                'Matched Site': site.site_name if site else '',
                'Directory Site': site.site_code if site else '',
                'Reporting Site': site.reporting_site if site else '',
                'Reporting Site Source': site.reporting_site_source if site else '',
                'Standardized Address': site.address if site else clean(row.get('Asset Location')),
                'Directory City': site.city if site else '',
                'Directory State': site.state if site else '',
                'Directory Region': site.region if site else '',
                'Directory Facility': site.facility if site else '',
                'Directory Data Hall': site.data_hall if site else '',
                'Directory Customer': site.customer if site else '',
                'Match Candidates': ', '.join(result.candidates),
            })
        enriched = df.reset_index(drop=True).copy()
        if records:
            derived = pd.DataFrame(records)
            for column in derived:
                if column not in enriched.columns:
                    enriched[column] = derived[column]
        return enriched


@lru_cache(maxsize=8)
def _cached_registry(path: str, modified_ns: int, size: int) -> SiteRegistry:
    """In-memory reuse only; edits to the map invalidate the cache."""
    return SiteRegistry.from_map(path)


def default_registry() -> SiteRegistry:
    map_path = DEFAULT_MAP.resolve()
    stat = map_path.stat()  # Clear FileNotFoundError if the sole map is missing.
    return _cached_registry(str(map_path), stat.st_mtime_ns, stat.st_size)


def standardize_address(value, registry=None, locode=None, site_name=None) -> str:
    """Return the map's canonical address, or keep an unmatched address intact."""
    original = clean(value)
    if not original and not clean(locode):
        return original
    registry = registry if registry is not None else default_registry()
    result = registry.resolve(original, locode, site_name)
    return result.site.address if result.status == 'matched' and result.site else original


def get_site_reference(address, registry=None, locode=None, site_name=None):
    """Return verified map fields, or None for unknown/conflicting addresses."""
    registry = registry if registry is not None else default_registry()
    result = registry.resolve(address, locode, site_name)
    if result.status != 'matched' or result.site is None:
        return None
    site = result.site
    return {
        'site_name': site.site_name, 'site': site.reporting_site,
        'map_site': site.site_code, 'site_source': site.reporting_site_source,
        'locode': site.locode, 'city': site.city, 'state': site.state,
        'region': site.region, 'facility': site.facility,
        'data_hall': site.data_hall, 'customer': site.customer,
        'address': site.address,
    }


def extract_site(subject):
    """Extract a Site identifier from a Salesforce case subject."""
    subject = clean(subject)
    if not subject:
        return ''
    parts = [part.strip() for part in subject.split('|')]
    if len(parts) >= 2 and re.fullmatch(r'[A-Z]{3,}\d+|[A-Z]+-[A-Z]+-?\d+[A-Z]?', parts[1], flags=re.I):
        return parts[1].upper()
    match = re.search(r'\bUS-[A-Z]+-?\d+[A-Z]?\b', subject, flags=re.I)
    return match.group(0).upper() if match else ''


def run_report(report_path, map_path=None, output_path=None):
    """Enrich report sheets using CoreWeave_Map.json; leave input untouched."""
    report_path = Path(report_path)
    output_path = Path(output_path) if output_path else report_path.with_name(
        report_path.stem + '_Standardized.xlsx'
    )
    registry = SiteRegistry.from_map(map_path)
    sheets = pd.read_excel(report_path, sheet_name=None, dtype={'Case Number': str})
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for name, frame in sheets.items():
            if 'Asset Location' in frame.columns:
                frame = registry.enrich(frame)
            frame.to_excel(writer, sheet_name=name, index=False)
    return output_path


def main():
    parser = argparse.ArgumentParser(description='Standardize addresses against CoreWeave_Map.json')
    parser.add_argument('report', help='Input report XLSX')
    parser.add_argument('--site-map', default=str(DEFAULT_MAP), help='Path to CoreWeave_Map.json')
    parser.add_argument('--output', help='Output XLSX (default: *_Standardized.xlsx)')
    args = parser.parse_args()
    print(run_report(args.report, args.site_map, args.output))


if __name__ == '__main__':
    main()
