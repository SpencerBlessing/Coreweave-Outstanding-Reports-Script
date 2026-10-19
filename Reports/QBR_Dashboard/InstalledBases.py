
import os
import re
import unicodedata
from pathlib import Path

import mysql.connector
import pandas as pd


# ---------- Configuration ----------

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent.parent

OUTPUT_FOLDER = BASE_DIR / "Output"
OUTPUT_FILE = OUTPUT_FOLDER / "Installed Bases.xlsx"
ADDRESS_FILE = PROJECT_DIR / "Address Cache" / "address_input.xlsx"

DB_CONFIG = {
    "host": "10.2.9.8",
    "port": 3306,
    "user": "readonly",
    "password": "nowrite",
    "database": "gsp2",
}

# Keep this enabled only if a province-level label is acceptable
# in the dashboard's City column.
ALLOW_PROVINCE_AS_CITY = True

# ---------- SQL query ----------

QUERY = """
WITH all_systems AS (
    SELECT
        system_serial,
        system_address,
        system_skuNumber,
        MIN(service_startDate) AS service_startDate,
        MAX(service_expireDate) AS service_expireDate
    FROM systemServiceV
    WHERE endUser_id IN (2859, 3419)
      AND system_address IS NOT NULL
      AND service_startDate IS NOT NULL
    GROUP BY
        system_serial,
        system_address,
        system_skuNumber
)
SELECT
    s.system_serial AS `System Serial`,
    s.system_skuNumber AS `System SKU`,
    s.service_startDate AS `Service Start Date`,
    s.service_expireDate AS `Service Expired Date`,
    s.system_address AS `System Address`,
    l.formattedAddress AS `Formatted Address`,
    COALESCE(l.mail_locality, l.locality) AS `City`,
    COALESCE(
        l.mail_administrativeArea,
        l.administrative_area_level_1
    ) AS `State`,
    COALESCE(l.mail_country, l.country) AS `Country`,
    l.locality AS `Original Locality`,
    l.administrative_area_level_1 AS `Original State`,
    l.country AS `Original Country`
FROM all_systems AS s
LEFT JOIN location AS l
    ON s.system_address = l.address
ORDER BY s.system_serial;
"""

# ---------- Normalization ----------

def normalize(series):
    """Normalize text for consistent address and location matching."""
    values = (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("ascii")
        .str.upper()
        .str.replace(r"[^A-Z0-9]+", " ", regex=True)
        .str.strip()
    )
    return values

def is_blank(series):
    return series.fillna("").astype(str).str.strip().eq("")

def first_nonblank(series):
    """Return the first available nonblank value."""
    for value in series:
        if pd.notna(value) and str(value).strip():
            return str(value).strip()
    return ""

# ---------- Address enrichment ----------

def apply_address_cache(df):
    # Ensure expected columns exist.
    for column in ["City", "State", "Region", "Country"]:
        if column not in df.columns:
            df[column] = ""
        df[column] = df[column].fillna("").astype(str).str.strip()

    # Restore missing values from the original GSP fields.
    for target, source in {
        "City": "Original Locality",
        "State": "Original State",
    }.items():
        if source in df.columns:
            missing = is_blank(df[target])
            fallback = df[source].fillna("").astype(str).str.strip()
            df.loc[missing, target] = fallback[missing]

    # Recover city from Swedish postal addresses, e.g. 791 77 Falun.
    if "System Address" in df.columns:
        address = df["System Address"].fillna("").astype(str)
        missing_city = is_blank(df["City"])

        postal_city = address.str.extract(
            r"\b\d{3}\s?\d{2}\s+([^,]+)",
            expand=False,
        )

        mask = (
            missing_city
            & address.str.contains(r"\bSweden\s*$", case=False, regex=True)
            & postal_city.notna()
        )
        df.loc[mask, "City"] = postal_city[mask].str.strip()

    # Optional Álava fallback. This is a province, not a verified city.
    if ALLOW_PROVINCE_AS_CITY and "Formatted Address" in df.columns:
        mask = (
            is_blank(df["City"])
            & df["Formatted Address"].fillna("").astype(str).str.contains(
                r"\bAlava\b|\bÁlava\b",
                case=False,
                regex=True,
            )
        )
        df.loc[mask, "City"] = "Álava"

    # Correct known country-parsing issue before building lookup keys.
    mask = (
        df["City"].str.strip().str.upper().eq("VOLO")
        & df["Country"].str.strip().str.upper().eq("IL")
    )
    df.loc[mask, "State"] = "Illinois"
    df.loc[mask, "Country"] = "United States"

    # For these countries, use the country name as State when missing.
    # This is a reporting placeholder, not an actual state/province.
    missing_state = is_blank(df["State"])
    country = df["Country"].str.strip().str.upper()

    mask = missing_state & country.isin(["UK", "NORWAY", "SWEDEN"])
    df.loc[mask, "State"] = df.loc[mask, "Country"]

    # Load and validate the address cache.
    if not ADDRESS_FILE.exists():
        raise FileNotFoundError(f"Address cache not found: {ADDRESS_FILE}")

    cache = pd.read_excel(ADDRESS_FILE, dtype=str)
    cache.columns = cache.columns.astype(str).str.strip()

    required = ["Asset Location", "City", "State", "Region"]
    missing_columns = [c for c in required if c not in cache.columns]

    if missing_columns:
        raise ValueError(f"Address cache missing columns: {missing_columns}")

    for column in required:
        cache[column] = cache[column].fillna("").astype(str).str.strip()

    print("\nMost common cache entries with City and State but no Region:")
    cache_missing_region = cache[
        cache["City"].ne("")
        & cache["State"].ne("")
        & cache["Region"].eq("")
    ]
    print(
        cache_missing_region[["City", "State", "Region"]]
        .drop_duplicates()
        .head(40)
        .to_string(index=False)
    )

    # Normalize addresses for exact matching.
    cache["_address_key"] = normalize(cache["Asset Location"])
    df["_address_key"] = normalize(df["System Address"])

    # Ensure Country exists in the address cache.
    if "Country" not in cache.columns:
        cache["Country"] = ""

    cache["Country"] = (
        cache["Country"].fillna("").astype(str).str.strip()
    )

    cache["_city_key"] = normalize(cache["City"])
    cache["_state_key"] = normalize(cache["State"])
    cache["_country_key"] = normalize(cache["Country"])

    # Exclude addresses whose nonblank cache values conflict.
    address_conflicts = set()

    for address_key, group in cache.groupby("_address_key"):
        if not address_key:
            continue

        for column in ["City", "State", "Region"]:
            values = normalize(group[column])
            if values[values.ne("")].nunique() > 1:
                address_conflicts.add(address_key)
                break

    safe_cache = cache[
        cache["_address_key"].ne("")
        & ~cache["_address_key"].isin(address_conflicts)
    ].copy()

    # Preserve the first available nonblank value for each address.
    safe_addresses = safe_cache.groupby("_address_key").agg({
        "City": first_nonblank,
        "State": first_nonblank,
        "Region": first_nonblank,
    })

    # Fill missing values using exact normalized address matches.
    for column in ["City", "State", "Region"]:
        mapped = df["_address_key"].map(safe_addresses[column])
        missing = is_blank(df[column])
        df.loc[missing, column] = mapped[missing].fillna("")

    # Build location keys only after corrections and address enrichment.
    df["_city_key"] = normalize(df["City"])
    df["_state_key"] = normalize(df["State"])
    df["_country_key"] = normalize(df["Country"])

    # Build unambiguous region lookups from cache entries.
    def build_region_lookup(key_column, valid_rows):
        valid = cache[
            valid_rows
            & cache[key_column].ne("|")
            & cache["Region"].ne("")
        ]

        lookup = {}

        for key, group in valid.groupby(key_column):
            regions = group["Region"].drop_duplicates().tolist()
            if len(regions) == 1:
                lookup[key] = regions[0]

        return lookup


    # Build lookup keys first.
    cache["_city_state_key"] = (
        cache["_city_key"] + "|" + cache["_state_key"]
    )
    cache["_city_country_key"] = (
        cache["_city_key"] + "|" + cache["_country_key"]
    )

    df["_city_state_key"] = df["_city_key"] + "|" + df["_state_key"]
    df["_city_country_key"] = df["_city_key"] + "|" + df["_country_key"]

    # Identify cache entries that have enough information for each lookup.
    valid_city_state = (
        cache["_city_key"].ne("")
        & cache["_state_key"].ne("")
    )

    valid_city_country = (
        cache["_city_key"].ne("")
        & cache["_country_key"].ne("")
    )

    city_state_regions = build_region_lookup(
        "_city_state_key",
        valid_city_state,
    )

    city_country_regions = build_region_lookup(
        "_city_country_key",
        valid_city_country,
    )

    # First use the cache's City + State mapping.
    missing_region = is_blank(df["Region"])
    mapped = df["_city_state_key"].map(city_state_regions)
    df.loc[missing_region, "Region"] = mapped[missing_region].fillna("")

    # Then try City + Country for international cache entries.
    missing_region = is_blank(df["Region"])
    mapped = df["_city_country_key"].map(city_country_regions)
    df.loc[missing_region, "Region"] = mapped[missing_region].fillna("")

    # Keep these labels consistent with your organization's reporting.
    us_state_regions = {
        "AL": "Southeast (US)", "ALABAMA": "Southeast (US)",
        "AK": "West Coast (US)", "ALASKA": "West Coast (US)",
        "AZ": "West Coast (US)", "ARIZONA": "West Coast (US)",
        "AR": "Central (US)", "ARKANSAS": "Central (US)",
        "CA": "West Coast (US)", "CALIFORNIA": "West Coast (US)",
        "CO": "Central (US)", "COLORADO": "Central (US)",
        "CT": "East Coast (US)", "CONNECTICUT": "East Coast (US)",
        "DE": "East Coast (US)", "DELAWARE": "East Coast (US)",
        "FL": "Southeast (US)", "FLORIDA": "Southeast (US)",
        "GA": "Southeast (US)", "GEORGIA": "Southeast (US)",
        "HI": "West Coast (US)", "HAWAII": "West Coast (US)",
        "ID": "West Coast (US)", "IDAHO": "West Coast (US)",
        "IL": "Central (US)", "ILLINOIS": "Central (US)",
        "IN": "Central (US)", "INDIANA": "Central (US)",
        "IA": "Central (US)", "IOWA": "Central (US)",
        "KS": "Central (US)", "KANSAS": "Central (US)",
        "KY": "Southeast (US)", "KENTUCKY": "Southeast (US)",
        "LA": "Southeast (US)", "LOUISIANA": "Southeast (US)",
        "ME": "East Coast (US)", "MAINE": "East Coast (US)",
        "MD": "East Coast (US)", "MARYLAND": "East Coast (US)",
        "MA": "East Coast (US)", "MASSACHUSETTS": "East Coast (US)",
        "MI": "Central (US)", "MICHIGAN": "Central (US)",
        "MN": "Central (US)", "MINNESOTA": "Central (US)",
        "MS": "Southeast (US)", "MISSISSIPPI": "Southeast (US)",
        "MO": "Central (US)", "MISSOURI": "Central (US)",
        "MT": "West Coast (US)", "MONTANA": "West Coast (US)",
        "NE": "Central (US)", "NEBRASKA": "Central (US)",
        "NV": "West Coast (US)", "NEVADA": "West Coast (US)",
        "NH": "East Coast (US)", "NEW HAMPSHIRE": "East Coast (US)",
        "NJ": "East Coast (US)", "NEW JERSEY": "East Coast (US)",
        "NM": "West Coast (US)", "NEW MEXICO": "West Coast (US)",
        "NY": "East Coast (US)", "NEW YORK": "East Coast (US)",
        "NC": "Southeast (US)", "NORTH CAROLINA": "Southeast (US)",
        "ND": "Central (US)", "NORTH DAKOTA": "Central (US)",
        "OH": "Central (US)", "OHIO": "Central (US)",
        "OK": "Central (US)", "OKLAHOMA": "Central (US)",
        "OR": "West Coast (US)", "OREGON": "West Coast (US)",
        "PA": "East Coast (US)", "PENNSYLVANIA": "East Coast (US)",
        "RI": "East Coast (US)", "RHODE ISLAND": "East Coast (US)",
        "SC": "Southeast (US)", "SOUTH CAROLINA": "Southeast (US)",
        "SD": "Central (US)", "SOUTH DAKOTA": "Central (US)",
        "TN": "Southeast (US)", "TENNESSEE": "Southeast (US)",
        "TX": "Central (US)", "TEXAS": "Central (US)",
        "UT": "West Coast (US)", "UTAH": "West Coast (US)",
        "VT": "East Coast (US)", "VERMONT": "East Coast (US)",
        "VA": "East Coast (US)", "VIRGINIA": "East Coast (US)",
        "WA": "West Coast (US)", "WASHINGTON": "West Coast (US)",
        "WV": "East Coast (US)", "WEST VIRGINIA": "East Coast (US)",
        "WI": "Central (US)", "WISCONSIN": "Central (US)",
        "WY": "West Coast (US)", "WYOMING": "West Coast (US)",
        "DC": "East Coast (US)",
    }

    country_key = df["Country"].str.strip().str.upper()
    state_key = normalize(df["State"])

    is_us = country_key.isin(["USA", "UNITED STATES"])

    mapped_regions = state_key.map(us_state_regions)

    df.loc[is_us & mapped_regions.notna(), "Region"] = (
        mapped_regions[is_us & mapped_regions.notna()]
    )

    # Use Europe as the fallback region for Spanish sites.
    missing_region = is_blank(df["Region"])
    is_spain = df["Country"].str.strip().str.upper().eq("SPAIN")
    df.loc[missing_region & is_spain, "Region"] = "Europe"

    # Diagnostics.
    cache_address_keys = set(
        cache.loc[cache["_address_key"].ne(""), "_address_key"]
    )

    missing_city = is_blank(df["City"])
    unresolved_addresses = df.loc[
        missing_city, ["System Address"]
    ].drop_duplicates()

    unresolved_keys = normalize(unresolved_addresses["System Address"])
    cache_matches = unresolved_keys.isin(cache_address_keys).sum()

    print(f"Conflicting cache addresses skipped: {len(address_conflicts):,}")
    print(f"Unique addresses still missing City: {len(unresolved_addresses):,}")
    print(f"Those addresses matching cache keys: {cache_matches:,}")

    for column in ["City", "State", "Region"]:
        missing_count = is_blank(df[column]).sum()
        percentage = missing_count / len(df) * 100 if len(df) else 0
        print(
            f"Rows still missing {column}: "
            f"{missing_count:,} ({percentage:.2f}%)"
        )

    df.drop(
        columns=[
            "_address_key",
            "_city_key",
            "_state_key",
            "_country_key",
            "_city_state_key",
            "_city_country_key",
        ],
        errors="ignore",
        inplace=True,
    )

    # Fill missing/unknown countries only when location is recognizable.
    country_missing = (
        df["Country"].fillna("").str.strip().str.upper().isin(["", "UNKNOWN"])
    )

    # Locations with strong evidence for a country
    country_corrections = {
        ("FALUN", "SWEDEN"): "Sweden",
        ("SPARKS", "NV"): "USA",
        ("BENSENVILLE", "IL"): "USA",
    }

    for (city, state), country in country_corrections.items():
        mask = (
            country_missing
            & df["City"].fillna("").str.strip().str.upper().eq(city)
            & df["State"].fillna("").str.strip().str.upper().eq(state)
        )
        df.loc[mask, "Country"] = country

    # Infer missing countries from explicit address evidence.
    country_missing = (
        df["Country"].fillna("").str.strip().str.upper().isin(["", "UNKNOWN"])
    )

    address = df["System Address"].fillna("").str.upper()

    country_rules = [
        (address.str.contains(r", USA\b|, UNITED STATES\b", regex=True), "USA"),
        (address.str.contains(r", CANADA\b", regex=True), "Canada"),
    ]

    for mask, country in country_rules:
        df.loc[country_missing & mask, "Country"] = country

    return df

# ---------- Main export ----------

def main():
    if not DB_CONFIG["password"]:
        raise RuntimeError(
            "Database password not set. Set the GSP_DB_PASSWORD "
            "environment variable before running this script."
        )

    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    connection = None

    print("\n" + "=" * 60)
    print("          COREWEAVE INSTALLED BASES EXPORT")
    print("=" * 60)

    try:
        # 1. Connect to GSP.
        print("\n[1/5] Connecting to GSP database...")
        connection = mysql.connector.connect(**DB_CONFIG)
        print("      Connected successfully.")

        # 2. Run the installed-bases query.
        print("\n[2/5] Running installed-bases query...")

        cursor = connection.cursor()

        try:
            cursor.execute(QUERY)
            columns = [item[0] for item in cursor.description]
            rows = cursor.fetchall()
        finally:
            cursor.close()

        df = pd.DataFrame(rows, columns=columns)
        print(f"      Records retrieved: {len(df):,}")

        # 3. Enrich location information.
        print("\n[3/5] Applying location enrichment...")
        df = apply_address_cache(df)

        # Audit unknown countries.
        unknown_country = df[
            df["Country"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
            .isin(["", "UNKNOWN"])
        ]

        print("\n--- LOCATION AUDIT ---")
        print(f"Records with unknown country: {len(unknown_country):,}")

        if not unknown_country.empty:
            print("\nMost common states:")
            print(unknown_country["State"].value_counts().head(10))

            print("\nMost common cities:")
            print(unknown_country["City"].value_counts().head(10))

            audit_columns = [
                "City", "State", "Country", "Region", "System Address"
            ]
            available = [
                col for col in audit_columns if col in df.columns
            ]

            print("\nSample unresolved addresses:")
            print(
                unknown_country[available]
                .head(15)
                .to_string(index=False)
            )

        # Check for missing State and Region values.
        for column in ["State", "Region"]:
            missing = (
                df[column].isna()
                | df[column].astype(str).str.strip().eq("")
            )

            print(f"\nMissing {column}: {missing.sum():,}")

            if missing.any():
                diagnostic_columns = [
                    "System Address",
                    "Formatted Address",
                    "City",
                    "State",
                    "Country",
                    "Original Locality",
                    "Original State",
                    "Original Country",
                ]

                available = [
                    col for col in diagnostic_columns
                    if col in df.columns
                ]

                print(f"Sample records missing {column}:")
                print(
                    df.loc[missing, available]
                    .head(10)
                    .to_string(index=False)
                )

       
        # Organize columns in a more useful order.
        preferred_columns = [
            # System identification
            "System Serial",
            "System SKU",

            # Location
            "City",
            "State",
            "Country",
            "Region",
            "Formatted Address",

            # Service coverage
            "Service Start Date",
            "Service Expired Date",
        ]

        # Keep all remaining columns afterward, preserving their original order.
        remaining_columns = [
            column for column in df.columns
            if column not in preferred_columns
        ]

        ordered_columns = [
            column for column in preferred_columns
            if column in df.columns
        ]

        df = df[ordered_columns + remaining_columns].copy()


        # Convert dates to Excel-compatible datetime values.
        for column in ["Service Start Date", "Service Expired Date"]:
            df[column] = pd.to_datetime(
                df[column],
                errors="coerce",
            )

        # Remove internal address and original-location audit fields from export.
        df = df.drop(
            columns=[
                "System Address",
                "Original Locality",
                "Original State",
                "Original Country",
            ],
            errors="ignore",
        )

        # 4. Export the workbook.
        print("\n[4/5] Exporting Excel file...")

        with pd.ExcelWriter(
            OUTPUT_FILE,
            engine="openpyxl",
            date_format="mm/dd/yy",
            datetime_format="mm/dd/yy",
        ) as writer:
            df.to_excel(
                writer,
                index=False,
                sheet_name="Installed Bases",
            )

            worksheet = writer.sheets["Installed Bases"]
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            # Display service dates without time.
            for column_name in ["Service Start Date", "Service Expired Date"]:
                column_index = df.columns.get_loc(column_name) + 1

                for row in range(2, worksheet.max_row + 1):
                    worksheet.cell(
                        row=row,
                        column=column_index,
                    ).number_format = "mm/dd/yy"

            # Automatically size columns.
            for column_cells in worksheet.columns:
                max_length = max(
                    len(str(cell.value or ""))
                    for cell in column_cells
                )

                worksheet.column_dimensions[
                    column_cells[0].column_letter
                ].width = min(max_length + 2, 35)
                
        # 5. Final summary.
        print("\n[5/5] Export completed successfully.")
        print("-" * 60)
        print(f"Output file : {OUTPUT_FILE}")
        print(f"Total rows  : {len(df):,}")
        print("-" * 60)
        print("DATA QUALITY SUMMARY")

        for column in ["City", "State", "Region"]:
            missing_count = is_blank(df[column]).sum()
            percentage = (
                missing_count / len(df) * 100
                if len(df) else 0
            )

            print(
                f"Missing {column:<6}: "
                f"{missing_count:>7,} ({percentage:5.2f}%)"
            )

        print(
            f"Unknown Country: "
            f"{df['Country'].fillna('').astype(str).str.strip().str.upper().isin(['', 'UNKNOWN']).sum():,}"
        )

        print("=" * 60)
        print("  INSTALLED BASES EXPORT FINISHED")
        print("=" * 60 + "\n")

    finally:
        if connection and connection.is_connected():
            connection.close()
            print("Database connection closed.")

if __name__ == "__main__":
    main()