
from pathlib import Path
import math

import pandas as pd


# --------------------------------------------------
# Configuration
# --------------------------------------------------

INPUT_FILE = Path(
    r"Reports\H100_Weekly_Completed\Input"
    r"\Consolidated_H100_Replacements_Weekly_10_18_to_NV.xlsx"
)

OUTPUT_DIR = Path("Reports/H100_Weekly_Completed/Output")
DEBUG_FILE = OUTPUT_DIR / "H100_Forecast_Debug.xlsx"

#TARGET_SITES = ["SBN1", "PDX2", "ORD3"]
TARGET_SITES = None
MERGE_FILE = OUTPUT_DIR / "H100_Case_Enrichment.xlsx"

# --------------------------------------------------
# Helpers
# --------------------------------------------------

def clean_text(series):
    return series.astype("string").str.strip().replace("", pd.NA)


def find_column(df, name):
    columns = {
        str(column).strip().casefold(): column
        for column in df.columns
    }

    if name.casefold() not in columns:
        raise ValueError(
            f"Could not find column: {name}\n"
            f"Available columns: {list(df.columns)}"
        )

    return columns[name.casefold()]


def print_section(title):
    print("\n" + "=" * 70)
    print(title)
    print("=" * 70)


def print_date_range(label, series):
    dates = pd.to_datetime(series, errors="coerce").dropna()

    if dates.empty:
        print(f"{label}: No valid dates found")
    else:
        print(
            f"{label}: {dates.min():%Y-%m-%d} "
            f"through {dates.max():%Y-%m-%d}"
        )

def get_date_range():
    """Prompt for an optional start and end date."""
    print("\nEnter a date range for Date Opened.")
    print("Leave both blank to include all available dates.")
    print("Format: YYYY-MM-DD (example: 2026-07-01)")

    while True:
        start_input = input("Start date: ").strip()
        end_input = input("End date: ").strip()

        if not start_input and not end_input:
            return None, None

        try:
            start_date = (
                pd.to_datetime(start_input)
                if start_input
                else None
            )
            end_date = (
                pd.to_datetime(end_input)
                if end_input
                else None
            )

            if start_date and end_date and start_date > end_date:
                print("Start date must be on or before end date.")
                continue

            return start_date, end_date

        except (ValueError, TypeError):
            print("Invalid date. Use YYYY-MM-DD.")


# --------------------------------------------------
# Load and prepare ticket-level data
# --------------------------------------------------

def load_tickets():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input workbook not found:\n{INPUT_FILE.resolve()}"
        )

    print_section("SOURCE WORKBOOK")

    file_info = INPUT_FILE.resolve()
    modified = pd.Timestamp(INPUT_FILE.stat().st_mtime, unit="s")

    print(f"File: {file_info}")
    print(f"Last modified: {modified:%Y-%m-%d %H:%M:%S}")
    print("Date filter: None — all available dates")
    print(
        f"Sites requested: {', '.join(TARGET_SITES)}"
        if TARGET_SITES is not None
        else "Sites requested: ALL LOCATIONS"
    )
    print("Status filter: At Supermicro")

    # Find the actual header row.
    preview = pd.read_excel(
        INPUT_FILE,
        header=None,
        dtype=str,
    )

    header_row = None

    for index, row in preview.iterrows():
        values = {
            str(value).strip().casefold()
            for value in row.dropna()
        }

        if (
            "nvidia gpu serial (sxm)" in values
            and "date opened" in values
        ):
            header_row = index
            break

    if header_row is None:
        raise ValueError("Could not locate the workbook header row.")

    print(f"Detected Excel header row: {header_row + 1}")

    df = pd.read_excel(INPUT_FILE, header=header_row)

    df.columns = [
        str(column).strip().replace("\n", " ")
        for column in df.columns
    ]

    site_col = find_column(df, "Site")
    failed_col = find_column(df, "Failed GPUs on server")
    jira_col = find_column(df, "Jira ticket")
    case_col = find_column(df, "Supermicro case #")
    status_col = find_column(df, "Status")
    date_col = find_column(df, "Date opened")

    print(f"Rows read from workbook: {len(df):,}")
    print_date_range("Source Date Opened range", df[date_col])

    df["_site"] = clean_text(df[site_col]).str.upper()
    df["_status"] = clean_text(df[status_col])
    df["_jira"] = clean_text(df[jira_col])
    df["_case"] = clean_text(df[case_col])
    df["_date_opened"] = pd.to_datetime(
        df[date_col],
        errors="coerce",
    )
    df["_failed_gpus"] = pd.to_numeric(
        df[failed_col],
        errors="coerce",
    )

    df = df[
       df["_status"].str.casefold() == "at supermicro"
    ].copy()

    print_section("SITE FILTER DEBUG")

    print("Rows by site before filtering:")
    print(
        df["_site"]
        .value_counts(dropna=False)
        .to_string()
    )

    before_filter = len(df)
    if TARGET_SITES is not None:
        requested_sites = [site.upper() for site in TARGET_SITES]
        df = df[df["_site"].isin(requested_sites)].copy()
        print(f"Site filter enabled: {requested_sites}")
    else:
        print("Site filter disabled — including ALL locations")

    print(f"\nRows before site filter: {before_filter:,}")
    print(f"Rows after site filter:  {len(df):,}")
    print(f"Rows excluded:           {before_filter - len(df):,}")

    if df.empty:
        raise ValueError(
            f"No records found for sites: {TARGET_SITES}"
        )

    print_date_range("Date range included", df["_date_opened"])
    
    print("\nRows by target site:")
    print(
        df["_site"]
        .value_counts()
        .reindex(TARGET_SITES, fill_value=0)
        .to_string()
    )

    print_section("TICKET DEDUPLICATION DEBUG")

    # Prefer Jira ticket, then Supermicro case number.
    df["_ticket_id"] = df["_jira"].fillna(df["_case"])

    # Keep records with missing IDs separate rather than merging them.
    missing_id = df["_ticket_id"].isna()
    df.loc[missing_id, "_ticket_id"] = [
        f"UNKNOWN_ROW_{index}"
        for index in df.index[missing_id]
    ]

    print(f"GPU rows before deduplication: {len(df):,}")
    print(f"Rows with no ticket/case ID:   {int(missing_id.sum()):,}")
    print(f"Unique ticket IDs:             {df['_ticket_id'].nunique():,}")

    # Warn if a ticket ID appears under multiple sites.
    site_counts = df.groupby("_ticket_id")["_site"].nunique()
    conflicting_sites = site_counts[site_counts > 1]

    if not conflicting_sites.empty:
        print(
            f"WARNING: {len(conflicting_sites)} ticket IDs "
            "appear under multiple sites."
        )
        print(conflicting_sites.to_string())
    else:
        print("Site consistency check: PASS")

    # Failed GPUs are repeated on each GPU row.
    # Use max to retain one ticket-level value.
    tickets = (
        df.groupby("_ticket_id", as_index=False)
        .agg(
            Site=("_site", "first"),
            Status=("_status", "first"),
            Date_Opened=("_date_opened", "min"),
            Failed_GPUs=("_failed_gpus", "max"),
        )
    )

    print(f"Ticket-level records created: {len(tickets):,}")
    print(
        "Rows removed by deduplication: "
        f"{len(df) - len(tickets):,}"
    )

    return df, tickets


# --------------------------------------------------
# Forecast and diagnostics
# --------------------------------------------------

def main():
    print_section("H100 REPLACEMENT FORECAST")

    raw_rows, tickets = load_tickets()

    sites = sorted(tickets["Site"].dropna().unique().tolist())

    if not sites:
        raise ValueError("No valid site names found in the ticket data.")

    print(f"\nAll locations included ({len(sites)}):")
    print(", ".join(sites))

    print_section("TICKET STATUS SUMMARY")

    status_summary = pd.crosstab(
        tickets["Site"],
        tickets["Status"].fillna("Unknown"),
    ).reindex(TARGET_SITES, fill_value=0)

    print(status_summary.to_string())

    print("\nUnique tickets by site:")
    print(
        tickets["Site"]
        .value_counts()
        .reindex(TARGET_SITES, fill_value=0)
        .to_string()
    )

    print_section("FAILED-GPU DATA VALIDATION")

    positive = tickets[
        tickets["Failed_GPUs"].notna()
        & (tickets["Failed_GPUs"] > 0)
    ].copy()

    zero = tickets[tickets["Failed_GPUs"] == 0].copy()
    missing = tickets[tickets["Failed_GPUs"].isna()].copy()
    negative = tickets[tickets["Failed_GPUs"] < 0].copy()

    print(f"Unique tickets:                  {len(tickets):,}")
    print(f"Positive failed-GPU counts:      {len(positive):,}")
    print(f"Zero failed-GPU counts:          {len(zero):,}")
    print(f"Missing failed-GPU counts:       {len(missing):,}")
    print(f"Negative failed-GPU counts:      {len(negative):,}")

    if not negative.empty:
        print("\nWARNING: Negative failed-GPU values found:")
        print(
            negative[
                ["Site", "Status", "Date_Opened", "Failed_GPUs"]
            ].to_string(index=False)
        )

    if positive.empty:
        raise ValueError(
            "No tickets have positive failed-GPU counts."
        )

    # Historical averages use positive counts only.
    site_stats = (
        positive.groupby("Site")["Failed_GPUs"]
        .agg(Tickets_Used="count", Total_Failed_GPUs="sum", Average="mean")
        .reindex(TARGET_SITES)
    )

    overall_average = positive["Failed_GPUs"].mean()

    print_section("HISTORICAL AVERAGES")

    for site in sites:
        row = site_stats.loc[site]

        if pd.isna(row["Average"]):
            print(f"{site}: no usable history")
        else:
            print(
                f"{site}: {int(row['Tickets_Used'])} tickets, "
                f"{int(row['Total_Failed_GPUs'])} failed GPUs, "
                f"average = {row['Average']:.2f} GPUs/ticket"
            )

    print(f"\nOverall average: {overall_average:.2f} GPUs/ticket")

    # User enters available GPU supply.
    while True:
        try:
            gpu_supply = int(
                input("\nEnter total available replacement GPUs: ")
            )

            if gpu_supply < 0:
                print("Enter zero or a positive whole number.")
                continue

            break

        except ValueError:
            print("Enter a whole number.")

    # Distribute GPUs as evenly as possible.
    base = gpu_supply // len(TARGET_SITES)
    remainder = gpu_supply % len(TARGET_SITES)

    allocations = {
        site: base + (1 if i < remainder else 0)
        for i, site in enumerate(TARGET_SITES)
    }

    print_section("FORECAST CALCULATION")

    results = []

    for site in TARGET_SITES:
        allocated = allocations[site]
        average = site_stats.loc[site, "Average"]

        if pd.isna(average) or average <= 0:
            projected = 0
            estimated_use = 0.0
            remaining = float(allocated)
            average_value = None
        else:
            projected = math.floor(allocated / average)
            estimated_use = projected * average
            remaining = allocated - estimated_use
            average_value = float(average)

        results.append({
            "Site": site,
            "GPUs_Allocated": allocated,
            "Average_GPUs_Per_Ticket": average_value,
            "Projected_Tickets": projected,
            "Estimated_GPU_Use": estimated_use,
            "Estimated_GPUs_Remaining": remaining,
        })

        print(f"\n{site}")
        print(f"  GPUs allocated:          {allocated}")
        print(
            "  Calculation:             "
            f"{allocated} / "
            f"{average_value:.2f}"
            if average_value is not None
            else "  Calculation:             No usable average"
        )
        print(f"  Projected tickets:       {projected}")
        print(f"  Estimated GPU use:       {estimated_use:.2f}")
        print(f"  Estimated GPUs remaining: {remaining:.2f}")

    results_df = pd.DataFrame(results)

    print_section("OVERALL RESULT")

    total_projected = int(results_df["Projected_Tickets"].sum())
    total_used = results_df["Estimated_GPU_Use"].sum()
    total_remaining = gpu_supply - total_used

    # Data for the separate chart class
    forecast_data = []

    for result in results:
        site = result["Site"]
        stats = site_stats.loc[site]

        forecast_data.append({
            "Site": site,
            "Historical_Tickets": int(stats["Tickets_Used"]),
            "Historical_Failed_GPUs": int(stats["Total_Failed_GPUs"]),
            "Average_GPUs_Per_Ticket": result["Average_GPUs_Per_Ticket"],
            "GPUs_Allocated": result["GPUs_Allocated"],
            "Projected_Tickets": result["Projected_Tickets"],
            "Estimated_GPU_Use": result["Estimated_GPU_Use"],
            "GPUs_Remaining": result["Estimated_GPUs_Remaining"],
        })

    forecast_data.append({
        "Site": "TOTAL",
        "Historical_Tickets": int(site_stats["Tickets_Used"].sum()),
        "Historical_Failed_GPUs": int(
            site_stats["Total_Failed_GPUs"].sum()
        ),
        "Average_GPUs_Per_Ticket": overall_average,
        "GPUs_Allocated": gpu_supply,
        "Projected_Tickets": total_projected,
        "Estimated_GPU_Use": total_used,
        "GPUs_Remaining": total_remaining,
    })

    forecast_data_df = pd.DataFrame(forecast_data)

    # --------------------------------------------------
    # Export audit/debug workbook
    # --------------------------------------------------

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    debug_tickets = tickets.sort_values(
        ["Site", "Date_Opened", "_ticket_id"]
        if "_ticket_id" in tickets.columns
        else ["Site", "Date_Opened"]
    )

    data_quality = pd.DataFrame([
        {"Metric": "Input workbook", "Value": str(INPUT_FILE.resolve())},
        {"Metric": "File last modified", "Value": str(
            pd.Timestamp(INPUT_FILE.stat().st_mtime, unit="s")
        )},
        {"Metric": "Sites included", "Value": ", ".join(TARGET_SITES)},
        {"Metric": "GPU rows after status/site filters", "Value": len(raw_rows)},
        {"Metric": "Unique tickets", "Value": len(tickets)},
        {"Metric": "Unique tickets", "Value": len(tickets)},
        {"Metric": "Positive failed-GPU tickets", "Value": len(positive)},
        {"Metric": "Zero failed-GPU tickets", "Value": len(zero)},
        {"Metric": "Missing failed-GPU tickets", "Value": len(missing)},
        {"Metric": "Negative failed-GPU tickets", "Value": len(negative)},
        {"Metric": "Earliest ticket open date", "Value": str(
            tickets["Date_Opened"].min()
        )},
        {"Metric": "Latest ticket open date", "Value": str(
            tickets["Date_Opened"].max()
        )},
        {"Metric": "Overall average GPUs per ticket", "Value": overall_average},
        {"Metric": "GPU supply entered", "Value": gpu_supply},
        {"Metric": "Projected tickets", "Value": total_projected},
    ])

    # --------------------------------------------------
    # Export inputs for the weekly chart script
    # --------------------------------------------------

    forecast_inputs = pd.DataFrame([
        {"Metric": "Total_Tickets", "Value": len(tickets)},
        {
            "Metric": "Tickets_With_Valid_GPU_Count",
            "Value": len(positive),
        },
        {
            "Metric": "Tickets_Without_Valid_GPU_Count",
            "Value": len(tickets) - len(positive),
        },
        {
            "Metric": "Average_GPUs_Per_Ticket",
            "Value": overall_average,
        },
        {"Metric": "Available_GPUs", "Value": gpu_supply},
        {"Metric": "Weekly_Ticket_Target", "Value": 300},
    ])

    with pd.ExcelWriter(DEBUG_FILE, engine="openpyxl") as writer:
        debug_tickets.to_excel(
            writer,
            sheet_name="Deduplicated Tickets",
            index=False,
        )

        status_summary.to_excel(
            writer,
            sheet_name="Status Summary",
        )

        site_stats.to_excel(
            writer,
            sheet_name="Historical Averages",
        )

        results_df.to_excel(
            writer,
            sheet_name="Forecast",
            index=False,
        )

        data_quality.to_excel(
            writer,
            sheet_name="Debug Summary",
            index=False,
        )

        forecast_data_df.to_excel(
            writer,
            sheet_name="Forecast Data",
            index=False,
        )

        forecast_inputs.to_excel(
            writer,
            sheet_name="Forecast Inputs",
            index=False,
        )

    print_section("DEBUG EXPORT")
    print(f"Debug workbook saved to:\n{DEBUG_FILE.resolve()}")
    print("Sheets exported:")
    print(
        ", ".join(
            [
                "Deduplicated Tickets",
                "Status Summary",
                "Historical Averages",
                "Forecast",
                "Debug Summary",
                "Forecast Data",
                "Forecast Inputs",
            ]
        )
    )
    print("\nThe source workbook was not modified.")

if __name__ == "__main__":
    main()
