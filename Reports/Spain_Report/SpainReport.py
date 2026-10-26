from pathlib import Path
from datetime import datetime
import re
import pandas as pd


# ============================================================
# PATHS
# ============================================================

PROJECT_DIR = Path(__file__).resolve().parents[2]

INPUT_DIR = (
    PROJECT_DIR
    / "Reports"
    / "Coreweave_Outstanding"
    / "Output"
)

input_reports = sorted(
    INPUT_DIR.glob("CoreWeave Outstanding Cases *.xlsx"),
    key=lambda path: path.stat().st_mtime,
    reverse=True,
)

if not input_reports:
    raise FileNotFoundError(
        f"No CoreWeave Outstanding Cases Excel reports found in:\n{INPUT_DIR}"
    )

INPUT_FILE = input_reports[0]

OUTPUT_DIR = PROJECT_DIR / "Reports" / "Spain_Report" / "Output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# SETTINGS
# ============================================================

HEADER_SEARCH_ROWS = 40

# Output fields, in the preferred order.
# Fields not present in the source export are skipped.
PREFERRED_COLUMNS = [
    "Case Number",
    "Date/Time Opened",
    "Open Age",
    "System Model",
    "Serial Number",
    "Case Owner",
    "Status",
    "Asset Location",
    "Address",
    "City",
    "State",
    "Region",
    "Country",
    "Faulty Part",
]

# Common Salesforce header variations mapped to the preferred names.
COLUMN_ALIASES = {
    "case number": "Case Number",
    "date/time opened": "Date/Time Opened",
    "created date": "Date/Time Opened",
    "case created date": "Date/Time Opened",
    "open age": "Open Age",
    "case age": "Open Age",
    "system model": "System Model",
    "product": "System Model",
    "serial number": "Serial Number",
    "serial": "Serial Number",
    "case owner": "Case Owner",
    "owner": "Case Owner",
    "status": "Status",
    "asset location": "Asset Location",
    "address": "Address",
    "city": "City",
    "state": "State",
    "country": "Country",
    "region": "Region",
    "faulty part": "Faulty Part",
    "faulty part number": "Faulty Part",
}


def normalize_header(value):
    """Normalize a column header for alias matching."""
    if pd.isna(value):
        return ""
    return " ".join(str(value).replace("\n", " ").split()).strip()

def clean_address_spacing(address):
    if not isinstance(address, str):
        return address

    # Collapse repeated spaces, tabs, and line breaks.
    address = re.sub(r"\s+", " ", address).strip()

    # Normalize spaces around commas.
    address = re.sub(r"\s*,\s*", ", ", address)

    return address

def find_header_row(raw_df):
    """Find the row containing 'Case Number' in a Salesforce export."""
    rows_to_check = min(HEADER_SEARCH_ROWS, len(raw_df))

    for row_index in range(rows_to_check):
        values = [
            normalize_header(value).casefold()
            for value in raw_df.iloc[row_index].tolist()
        ]
        if "case number" in values:
            return row_index

    raise ValueError(
        f"Could not find a 'Case Number' header in the first "
        f"{rows_to_check} rows. Check that the correct Salesforce export "
        f"is in: {INPUT_DIR}"
    )


def load_report(file_path):
    print(f"Loading report: {file_path}")

    raw_df = pd.read_excel(file_path, header=None, dtype=object)
    header_row = find_header_row(raw_df)

    headers = [normalize_header(v) for v in raw_df.iloc[header_row]]
    df = raw_df.iloc[header_row + 1:].copy()
    df.columns = headers

    df = df.loc[
        :, [c for c in df.columns if c and not c.casefold().startswith("unnamed:")]
    ]

    rename_map = {
        col: COLUMN_ALIASES[col.casefold()]
        for col in df.columns
        if col.casefold() in COLUMN_ALIASES
    }
    df = df.rename(columns=rename_map)

    if df.columns.duplicated().any():
        merged = {}
        for col in dict.fromkeys(df.columns):
            matching = df.loc[:, df.columns == col]
            merged[col] = matching.bfill(axis=1).iloc[:, 0]
        df = pd.DataFrame(merged, index=df.index)

    df = df.dropna(how="all")

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].map(
                lambda v: v.strip() if isinstance(v, str) else v
            )

    return df, header_row


def standardize_address(value):
    """Normalize known Spain site address variants to canonical addresses."""
    if pd.isna(value):
        return value

    address = re.sub(r"\s+", " ", str(value)).strip()
    address = re.sub(r"\s*,\s*", ", ", address)
    normalized = address.casefold()

    # Barcelona — Zona Franca
    is_barcelona_25 = (
        "08040" in normalized
        and "barcelona" in normalized
        and (
            "zona franca" in normalized
            or bool(
                re.search(
                    r"\b(?:carrer|calle)\b.*\b25\b",
                    normalized,
                )
            )
        )
    )

    if is_barcelona_25:
        return (
            "Carrer 25, nº 6, Parc Logístic de la Zona Franca, "
            "08040 Barcelona, Spain"
        )

    # Álava — ARASUR / Ribabellosa
    is_alava_arasur = (
        "01213" in normalized
        and any(
            term in normalized
            for term in (
                "ribera baja",
                "arasur",
                "ribabellosa",
                "rivabellosa",
            )
        )
    )

    if is_alava_arasur:
        return (
            "Plataforma Logística ARASUR, Avenida Ribera Baja, "
            "01213 Ribabellosa, Álava, Spain"
        )

    # Santa Creu de Calafell — normalize to Viladecans
    is_santa_creu = (
        "santa creu" in normalized
        and "carretera" in normalized
    )

    if is_santa_creu and (
        "viladecans" in normalized
        or "sant boi de llobregat" in normalized
        or "08830" in normalized
        or "08840" in normalized
    ):
        return (
            "Carretera de la Santa Creu de Calafell, 99, "
            "08840 Viladecans, Barcelona, Spain"
        )

    # Leave other locations unchanged, with whitespace cleaned.
    return address




def find_latest_previous_report(output_file):
    """Find the newest existing Spain report by its timestamped filename."""
    reports = sorted(
        OUTPUT_DIR.glob("Spain_Open_Cases_*.xlsx"),
        key=lambda path: path.name,
        reverse=True,
    )

    return reports[0] if reports else None

def load_previous_report(report_path):
    """Load previous Spain tickets, supporting different worksheet names."""
    sheets = pd.read_excel(report_path, sheet_name=None)

    frames = [
        df
        for sheet_name, df in sheets.items()
        if sheet_name.casefold() != "comparison"
        and "Case Number" in df.columns
    ]

    if not frames:
        raise ValueError(
            f"No ticket worksheet containing 'Case Number' found in "
            f"{report_path}"
        )

    previous_df = pd.concat(frames, ignore_index=True)

    previous_df["Case Number"] = (
        previous_df["Case Number"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    previous_df = previous_df[
        previous_df["Case Number"].ne("")
        & previous_df["Case Number"].ne("NAN")
    ]

    # Remove the unnecessary column from older reports.
    previous_df = previous_df.drop(columns=["Site"], errors="ignore")

    return previous_df.drop_duplicates(
        subset=["Case Number"],
        keep="last",
    )

def get_age_bucket(value):
    """Convert Open Age into the report's age buckets."""
    if pd.isna(value):
        return "Unknown"

    match = re.search(r"\d+(?:\.\d+)?", str(value))
    if not match:
        return "Unknown"

    age = float(match.group())

    if age <= 9:
        return "A: 0-9 days"
    elif age <= 15:
        return "B: 10-15 days"
    elif age <= 30:
        return "C: 16-30 days"
    return "D: 31+ days"



def build_comparison_tables(current_df, output_file):
    """Build summary, ticket-change lists, and current age/status breakdown."""

    def normalize_case_numbers(df):
        df = df.copy()

        def normalize(value):
            if pd.isna(value):
                return ""

            case_number = str(value).strip().upper()

            # Restore the CS- prefix on numeric-only case numbers.
            if case_number.isdigit():
                case_number = f"CS-{case_number.zfill(8)}"

            return case_number

        df["Case Number"] = df["Case Number"].apply(normalize)

        return df[
            df["Case Number"].ne("")
        ].drop_duplicates(
            subset="Case Number",
            keep="last",
        )

    current = normalize_case_numbers(current_df)

    previous_file = find_latest_previous_report(output_file)

    if previous_file:
        previous = normalize_case_numbers(
            load_previous_report(previous_file)
        )
        previous_name = previous_file.name
    else:
        previous = pd.DataFrame(columns=current.columns)
        previous_name = "No previous report found"

    current_ids = set(current["Case Number"])
    previous_ids = set(previous["Case Number"])

    opened_ids = current_ids - previous_ids
    removed_ids = previous_ids - current_ids
    shared_ids = current_ids & previous_ids

    opened = current[
        current["Case Number"].isin(opened_ids)
    ].copy()

    removed = previous[
        previous["Case Number"].isin(removed_ids)
    ].copy()

    # Compare statuses for tickets present in both reports.
    status_changes = pd.DataFrame()

    if (
        previous_file
        and "Status" in current.columns
        and "Status" in previous.columns
    ):
        old_status = (
            previous.set_index("Case Number")["Status"]
            .fillna("")
            .astype(str)
        )
        new_status = (
            current.set_index("Case Number")["Status"]
            .fillna("")
            .astype(str)
        )

        changed_ids = [
            case_id for case_id in shared_ids
            if old_status.get(case_id, "").strip()
            != new_status.get(case_id, "").strip()
        ]

        status_changes = current[
            current["Case Number"].isin(changed_ids)
        ][["Case Number"]].copy()

        status_changes["Previous Status"] = (
            status_changes["Case Number"].map(old_status)
        )
        status_changes["Current Status"] = (
            status_changes["Case Number"].map(new_status)
        )

    summary = pd.DataFrame([
        ["Previous report", previous_name],
        ["Current report", output_file.name],
        ["Tickets in previous report", len(previous_ids)],
        ["Tickets in current report", len(current_ids)],
        ["Tickets opened / newly listed", len(opened_ids)],
        ["Tickets closed / removed from report", len(removed_ids)],
        ["Tickets with status changes", len(status_changes)],
        ["Net change in open tickets",
         len(current_ids) - len(previous_ids)],
    ], columns=["Metric", "Value"])

    # Current open tickets grouped by age and status.
    age_status = current.copy()

    if "Open Age" in age_status.columns:
        age_status["Age Bucket"] = age_status["Open Age"].apply(
            get_age_bucket
        )
    else:
        age_status["Age Bucket"] = "Unknown"

    if "Status" not in age_status.columns:
        age_status["Status"] = "Unknown"

    age_status["Status"] = (
        age_status["Status"]
        .fillna("Unknown")
        .replace("", "Unknown")
    )

    age_order = [
        "A: 0-9 days",
        "B: 10-15 days",
        "C: 16-30 days",
        "D: 31+ days",
        "Unknown",
    ]

    age_by_status = pd.crosstab(
        age_status["Age Bucket"],
        age_status["Status"],
    ).reindex(age_order, fill_value=0)

    age_by_status.index.name = "Open Age"

    # Add total per age bucket.
    age_by_status["Total"] = age_by_status.sum(axis=1)

    return summary, opened, removed, status_changes, age_by_status


def normalize_case_numbers(df):
    df = df.copy()

    def normalize(value):
        if pd.isna(value):
            return ""

        case_number = str(value).strip().upper()

        # Restore the CS- prefix on older numeric-only case numbers.
        if case_number.isdigit():
            case_number = f"CS-{case_number.zfill(8)}"

        return case_number

    df["Case Number"] = df["Case Number"].apply(normalize)

    return df[
        df["Case Number"].ne("")
    ].drop_duplicates(
        subset=["Case Number"],
        keep="last",
    )

def main():
    print("=" * 60)
    print("SPAIN OPEN-CASES REPORT")
    print("=" * 60)

    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"Input report not found:\n{INPUT_FILE}")

    df, header_row = load_report(INPUT_FILE)

    # The source report has no Country column.
    # Identify Spain from the Asset Location address.
    if "Asset Location" not in df.columns:
        raise ValueError("Asset Location column not found.")

    location = (
        df["Asset Location"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
    )

    spain_df = df[
        location.str.contains(r"\bspain\b", regex=True)
    ].copy()

    # Clean addresses FIRST, then standardize known locations.
    spain_df["Asset Location"] = (
        spain_df["Asset Location"]
        .apply(clean_address_spacing)
        .apply(standardize_address)
    )

    print("\nStandardized Asset Locations:")
    location_counts = spain_df["Asset Location"].value_counts(dropna=False)

    for address, count in location_counts.items():
        print(f"{count:>5} cases | {address}")

    print(
        "\nUnique Asset Locations after standardization: "
        f"{spain_df['Asset Location'].nunique(dropna=False)}"
    )
    if "Case Number" not in spain_df.columns:
        raise ValueError("Case Number column not found.")
    
    missing_case_numbers = (
        spain_df["Case Number"].isna()
        | spain_df["Case Number"].astype(str).str.strip().eq("")
    ).sum()

    spain_df = spain_df[
        spain_df["Case Number"].notna()
        & spain_df["Case Number"].astype(str).str.strip().ne("")
    ].copy()

    spain_df["Case Number"] = spain_df["Case Number"].astype(str).str.strip()

    duplicate_cases = int(spain_df["Case Number"].duplicated().sum())
    spain_df = spain_df.drop_duplicates("Case Number", keep="first")

    # Remove unnecessary columns.
    spain_df = spain_df.drop(columns=["Site"], errors="ignore")

    # Keep the report columns in a consistent order.
    output_columns = [
        col for col in PREFERRED_COLUMNS if col in spain_df.columns
    ]
    extra_columns = [
        col for col in spain_df.columns
        if col not in output_columns and col != "Site"
    ]
    spain_df = spain_df[output_columns + extra_columns]

    output_file = OUTPUT_DIR / (
        f"Spain_Open_Cases_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    )

    # Split Spain cases by region into separate worksheets.
    region_text = (
        spain_df["Asset Location"]
        .fillna("")
        .astype(str)
        .str.casefold()
    )

    alava_df = spain_df[
        region_text.str.contains("arasur|ribera baja|ribabellosa|rivabellosa", regex=True)
    ].copy()

    barcelona_df = spain_df[
        region_text.str.contains(
            "barcelona|viladecans|sant boi de llobregat|zona franca",
            regex=True,
        )
    ].copy()

    
    summary, opened, removed, status_changes, age_by_status = (
        build_comparison_tables(spain_df, output_file)
    )

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        alava_df.to_excel(writer, sheet_name="Alava", index=False)
        barcelona_df.to_excel(writer, sheet_name="Barcelona", index=False)

        # Create the Comparison worksheet.
        writer.book.create_sheet("Comparison")

        sections = [
            ("REPORT SUMMARY", summary, False),
            ("NEWLY LISTED TICKETS", opened, False),
            ("REMOVED TICKETS", removed, False),
            ("STATUS CHANGES", status_changes, False),
            (
                "CURRENT OPEN TICKETS BY AGE AND STATUS",
                age_by_status,
                True,
            ),
        ]

        row = 0

        for title, dataframe, include_index in sections:
            ws = writer.sheets["Comparison"]
            ws.cell(row=row + 1, column=1, value=title)

            dataframe.to_excel(
                writer,
                sheet_name="Comparison",
                index=include_index,
                startrow=row + 1,
            )

            # Leave space before the next section.
            row += len(dataframe) + 4

        # Format all worksheets.
        for ws in writer.sheets.values():
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cells in ws.columns:
                letter = cells[0].column_letter
                width = max(
                    (
                        len(str(cell.value))
                        for cell in cells
                        if cell.value is not None
                    ),
                    default=12,
                )
                ws.column_dimensions[letter].width = min(
                    max(width + 2, 12), 48
                )

    print(f"Alava cases:     {len(alava_df)}")
    print(f"Barcelona cases: {len(barcelona_df)}")
    print(f"Tickets opened:  {len(opened)}")
    print(f"Tickets removed: {len(removed)}")
    print(f"Status changes:  {len(status_changes)}")
    print(f"Output saved:    {output_file}")

if __name__ == "__main__":
    main()