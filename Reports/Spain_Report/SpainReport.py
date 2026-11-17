from pathlib import Path

from datetime import datetime, date

import json
import re
import sqlite3

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

# Use the SAME read-only history database maintained by the CoreWeave report.
# Spain report generation never writes to this database.
HISTORY_DB = (
    PROJECT_DIR / "Reports" / "Coreweave_Outstanding"
    / "History" / "outstanding_history.sqlite3"
)

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
    "case creation date": "Date/Time Opened",

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

    case_sheets = [

        "CoreWeave Outstanding Cases",

        "H100 Thermal Issues",

    ]

    with pd.ExcelFile(file_path) as workbook:

        available_sheets = [

            name for name in case_sheets

            if name in workbook.sheet_names

        ]

        if not available_sheets:

            raise ValueError(

                "No CoreWeave case worksheets found in the report."

            )

        frames = [

            pd.read_excel(

                workbook,

                sheet_name=name,

                dtype={"Case Number": str}

            )

            for name in available_sheets

        ]

    df = pd.concat(frames, ignore_index=True)

    df.columns = [normalize_header(col) for col in df.columns]

    header_row = 0

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

def exclude_self_maintenance(df):

    """Exclude Self-Maintenance tickets, including any misfiled in break-fix tabs.

    Match the main CoreWeave report's classification: Subject beginning SELF
    or Issue Sub-Type equal to Self Maintenance. Also recognize explicit
    Self-Maintenance text within Subject, case-insensitively.
    """

    if df.empty:
        return df.copy(), 0

    excluded = pd.Series(False, index=df.index)

    if "Subject" in df.columns:
        subjects = df["Subject"].fillna("").astype(str).str.strip()
        excluded |= subjects.str.casefold().str.startswith("self")
        excluded |= subjects.str.contains(
            r"self[\s_-]*maintenance", case=False, regex=True
        )

    if "Issue Sub-Type" in df.columns:
        issue_type = (
            df["Issue Sub-Type"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.casefold()
        )
        excluded |= issue_type.isin({"self maintenance", "self-maintenance"})

    return df.loc[~excluded].copy(), int(excluded.sum())


def sort_newest_cases_first(df):

    """Show the newest opening dates first; missing/invalid dates go last."""

    if "Date/Time Opened" not in df.columns or df.empty:
        return df

    opened = pd.to_datetime(
        df["Date/Time Opened"], errors="coerce", format="mixed"
    )
    return (
        df.assign(__opened_sort=opened)
        .sort_values("__opened_sort", ascending=False, na_position="last", kind="stable")
        .drop(columns="__opened_sort")
        .reset_index(drop=True)
    )


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

def source_report_date(source_report):
    """Use the finalized CoreWeave report date, or today's date if unavailable.

    The main history comparison only considers snapshots strictly before the
    current report date. This also handles rerunning an older finalized report.
    """
    match = re.search(r"CoreWeave Outstanding Cases (\d{8})\.xlsx$", Path(source_report).name, re.I)
    if match:
        try:
            return datetime.strptime(match.group(1), "%m%d%Y").date()
        except ValueError:
            pass
    return date.today()


def load_previous_snapshot(source_report=INPUT_FILE, db_path=HISTORY_DB):
    """Read the last earlier finalized CoreWeave snapshot from its shared DB.

    History rows use the main report's `case_snapshots` schema:
        snapshot_date, case_number, category, case_data (JSON).

    Apply precisely the same Spain and break-fix filters to historical cases
    that the current Spain report applies to its source workbook. Never infer
    missing addresses or Site values from anything outside the source payload.
    No history database or no earlier snapshot means no comparison baseline;
    it does NOT mean all today's tickets should be classified as newly opened.
    """
    db_path = Path(db_path)
    empty = pd.DataFrame(columns=["Case Number", "Status"])
    if not db_path.is_file():
        print(f"History database not found: {db_path}")
        return empty, None

    cutoff_date = source_report_date(source_report).isoformat()
    # Read-only connection; the CoreWeave main report is the sole DB writer.
    with sqlite3.connect(db_path.as_uri() + "?mode=ro", uri=True) as conn:
        table_exists = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='case_snapshots'"
        ).fetchone()
        if not table_exists:
            print(f"History database has no case_snapshots table: {db_path}")
            return empty, None

        row = conn.execute(
            "SELECT MAX(snapshot_date) FROM case_snapshots WHERE snapshot_date < ?",
            (cutoff_date,),
        ).fetchone()
        previous_date = row[0] if row else None
        if not previous_date:
            print(f"No CoreWeave history snapshot prior to {cutoff_date}.")
            return empty, None

        records = conn.execute(
            """SELECT case_number, category, case_data
               FROM case_snapshots WHERE snapshot_date = ?""",
            (previous_date,),
        ).fetchall()

    previous_records = []
    for case_number, category, payload in records:
        if str(category).strip().casefold() == "self-maintenance":
            continue
        try:
            values = json.loads(payload)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Invalid historical case JSON for {case_number!r} on {previous_date}") from exc
        if not isinstance(values, dict):
            raise ValueError(f"Invalid historical case data for {case_number!r} on {previous_date}")
        values["Case Number"] = str(case_number)  # DB primary key is authoritative
        previous_records.append(values)

    if not previous_records:
        return empty, previous_date

    previous = pd.DataFrame.from_records(previous_records)
    previous.columns = [normalize_header(col) for col in previous.columns]
    previous = previous.rename(columns={
        col: COLUMN_ALIASES[col.casefold()]
        for col in previous.columns if col.casefold() in COLUMN_ALIASES
    })
    if previous.columns.duplicated().any():
        merged = {}
        for col in dict.fromkeys(previous.columns):
            matching = previous.loc[:, previous.columns == col]
            merged[col] = matching.bfill(axis=1).iloc[:, 0]
        previous = pd.DataFrame(merged, index=previous.index)

    previous, excluded = exclude_self_maintenance(previous)
    if "Asset Location" not in previous.columns:
        raise ValueError("Historical CoreWeave snapshot is missing Asset Location")
    spain = (previous["Asset Location"].fillna("").astype(str)
             .str.strip().str.casefold().str.contains(r"\bspain\b", regex=True))
    previous = previous.loc[spain].copy()
    previous["Asset Location"] = (previous["Asset Location"]
                                   .apply(clean_address_spacing)
                                   .apply(standardize_address))
    previous = normalize_case_numbers(previous)
    print(f"Using CoreWeave history snapshot {previous_date}: "
          f"{len(previous):,} Spain break-fix cases (excluded {excluded:,} self-maintenance cases).")
    return previous, previous_date


def get_age_bucket(value):
    """Interpret the actual CoreWeave Case Age labels, not their embedded numbers.

    The primary report stores e.g. ``D · Open > 30 days``. Extracting the first
    number would interpret that as 30 days, incorrectly assigning category C.
    Fallback numeric parsing is retained for older Excel exports.
    """
    if pd.isna(value):
        return "Unknown"

    text = str(value).strip()
    labels = {
        "A": "A: 0-9 days",
        "B": "B: 10-15 days",
        "C": "C: 16-30 days",
        "D": "D: 31+ days",
    }
    category = re.match(r"^([A-D])\s*(?:[:·.\-]|\s+Open\b)", text, flags=re.I)
    if category:
        return labels[category.group(1).upper()]

    # Older exports may contain a numeric age rather than a labeled bucket.
    try:
        age = float(text.replace(",", ""))
    except ValueError:
        return "Unknown"
    if age < 0:
        return "Unknown"
    if age <= 9:
        return labels["A"]
    if age <= 15:
        return labels["B"]
    if age <= 30:
        return labels["C"]
    return labels["D"]


AGE_ORDER = ["A: 0-9 days", "B: 10-15 days", "C: 16-30 days", "D: 31+ days"]


def build_breakdown_comparisons(current, previous, has_baseline):
    """Show previous/current/net counts, like the main CoreWeave report."""

    def counts(frame, field):
        if field not in frame.columns or frame.empty:
            return pd.Series(dtype="int64")
        return frame[field].fillna("Unknown").astype(str).str.strip().replace("", "Unknown").value_counts()

    def compare(title, current_counts, previous_counts, order=None):
        categories = set(current_counts.index) | (set(previous_counts.index) if has_baseline else set())
        if order is None:
            keys = sorted(categories, key=lambda x: (-int(current_counts.get(x, 0)), str(x)))
        else:
            keys = list(order)  # The Unknown age row is intentionally excluded.
        rows = []
        for key in keys:
            now = int(current_counts.get(key, 0))
            before = int(previous_counts.get(key, 0)) if has_baseline else None
            rows.append({title: key, "Current": now,
                         "Previous": before if has_baseline else "N/A",
                         "Net Change": now - before if has_baseline else "N/A"})
        return pd.DataFrame(rows, columns=[title, "Current", "Previous", "Net Change"])

    status_comparison = compare("Status", counts(current, "Status"), counts(previous, "Status"))

    def age_counts(frame):
        if "Open Age" not in frame.columns or frame.empty:
            return pd.Series(dtype="int64")
        return frame["Open Age"].map(get_age_bucket).value_counts()

    age_comparison = compare("Age Bucket", age_counts(current), age_counts(previous), AGE_ORDER)

    def region_counts(frame):
        if "Asset Location" not in frame.columns or frame.empty:
            return pd.Series(dtype="int64")
        loc = frame["Asset Location"].fillna("").astype(str).str.casefold()
        # Match the same two area filters used to write the Alava/Barcelona tabs.
        alava = loc.str.contains("arasur|ribera baja|ribabellosa|rivabellosa", regex=True)
        barcelona = loc.str.contains("barcelona|viladecans|sant boi de llobregat|zona franca", regex=True)
        return pd.Series({"Alava": int(alava.sum()), "Barcelona": int(barcelona.sum())})

    region_comparison = compare("Region", region_counts(current), region_counts(previous), ["Alava", "Barcelona"])
    return region_comparison, status_comparison, age_comparison


def build_comparison_tables(current_df, output_file, source_report=INPUT_FILE, db_path=HISTORY_DB):

    """Compare current Spain cases with the previous CoreWeave DB snapshot."""

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

    previous, previous_date = load_previous_snapshot(source_report, db_path)
    has_baseline = previous_date is not None
    if has_baseline:
        previous = normalize_case_numbers(previous)
    else:
        previous = pd.DataFrame(columns=current.columns)

    current_ids = set(current["Case Number"])
    previous_ids = set(previous["Case Number"])
    opened_ids = current_ids - previous_ids if has_baseline else set()
    removed_ids = previous_ids - current_ids if has_baseline else set()
    shared_ids = current_ids & previous_ids if has_baseline else set()

    opened = current[current["Case Number"].isin(opened_ids)].copy()
    removed = previous[previous["Case Number"].isin(removed_ids)].copy()

    # Status comparisons use normalized case numbers and data from the DB.
    status_changes = pd.DataFrame(columns=["Case Number", "Previous Status", "Current Status"])
    if has_baseline and "Status" in current.columns and "Status" in previous.columns:
        old_status = previous.set_index("Case Number")["Status"].fillna("").astype(str)
        new_status = current.set_index("Case Number")["Status"].fillna("").astype(str)
        changed_ids = sorted(
            case_id for case_id in shared_ids
            if old_status.get(case_id, "").strip() != new_status.get(case_id, "").strip()
        )
        if changed_ids:
            status_changes = pd.DataFrame({
                "Case Number": changed_ids,
                "Previous Status": [old_status.loc[case_id] for case_id in changed_ids],
                "Current Status": [new_status.loc[case_id] for case_id in changed_ids],
            })

    summary = pd.DataFrame([
        ["Comparison source", previous_date or "No previous database snapshot"],
        ["Previous snapshot", previous_date or "No previous database snapshot"],
        ["Current source date", source_report_date(source_report).isoformat()],
        ["Current source workbook", Path(source_report).name],
        ["Current Spain output", output_file.name],
        ["Tickets in previous snapshot", len(previous_ids) if has_baseline else "N/A"],
        ["Tickets present in both snapshots", len(shared_ids) if has_baseline else "N/A"],
        ["Tickets in current report", len(current_ids)],
        ["Tickets opened / newly listed", len(opened_ids) if has_baseline else "N/A"],
        ["Tickets removed from outstanding", len(removed_ids) if has_baseline else "N/A"],
        ["Tickets with status changes", len(status_changes) if has_baseline else "N/A"],
        ["Net change in open tickets", (len(current_ids) - len(previous_ids))
         if has_baseline else "N/A"],
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

    # Keep the four defined age buckets only; do not display an Unknown row.
    # Cases with missing ages remain in the ticket sheets and total counts,
    # but are omitted from this age/status breakdown rather than assigned
    # an invented age.
    unknown_age_count = int(age_status["Age Bucket"].eq("Unknown").sum())
    if unknown_age_count:
        print(f"Age/status breakdown: {unknown_age_count:,} ticket(s) have "
              "no valid age and are omitted from the age buckets.")

    age_order = AGE_ORDER

    age_by_status = pd.crosstab(

        age_status["Age Bucket"],

        age_status["Status"],

    ).reindex(age_order, fill_value=0)

    age_by_status.index.name = "Open Age"

    # Add total per age bucket.

    age_by_status["Total"] = age_by_status.sum(axis=1)

    region_comparison, status_comparison, age_comparison = build_breakdown_comparisons(
        current, previous, has_baseline
    )
    return (summary, opened, removed, status_changes, age_by_status,
            region_comparison, status_comparison, age_comparison)

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

    # Only break-fix cases belong in the Alava/Barcelona reports or Comparison.
    df, excluded_count = exclude_self_maintenance(df)
    print("Excluded the Self-Maintenance worksheet from the Spain report.")
    print(f"Additional Self-Maintenance cases excluded from other case sheets: {excluded_count}")

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

    # Keep the report columns in a consistent order.

    output_columns = [

        col for col in PREFERRED_COLUMNS if col in spain_df.columns

    ]

    extra_columns = [

        col for col in spain_df.columns

        if col not in output_columns

    ]

    spain_df = spain_df[output_columns + extra_columns]

    # Descending by Case Creation Date / Date/Time Opened, so 2026 appears first.
    spain_df = sort_newest_cases_first(spain_df)

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

    (summary, opened, removed, status_changes, age_by_status,
     region_comparison, status_comparison, age_comparison) = (
        build_comparison_tables(spain_df, output_file, source_report=INPUT_FILE)
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

            ("BREAK-FIX CASES BY REGION (DATABASE COMPARISON)", region_comparison, False),
            ("BREAK-FIX CASES BY STATUS (DATABASE COMPARISON)", status_comparison, False),
            ("BREAK-FIX CASES BY AGE (DATABASE COMPARISON)", age_comparison, False),

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
