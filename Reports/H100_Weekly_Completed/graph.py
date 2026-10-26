
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


# --------------------------------------------------
# Configuration
# --------------------------------------------------

SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "Output"
DEBUG_FILE = OUTPUT_DIR / "H100_Forecast_Debug.xlsx"

TARGET_SITES = ["SBN1", "PDX2", "ORD3"]


# --------------------------------------------------
# Helpers
# --------------------------------------------------

def section(title):
    print("\n" + "=" * 65)
    print(title)
    print("=" * 65)


def save_chart(fig, filename):
    path = OUTPUT_DIR / filename
    fig.tight_layout()
    fig.savefig(path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {path}")


def normalize_columns(df):
    df.columns = [
        str(column).strip()
        for column in df.columns
    ]
    return df


def find_column(df, candidates, description):
    """Find a column without depending on capitalization or underscores."""
    normalize = lambda value: "".join(
        char.lower()
        for char in str(value)
        if char.isalnum()
    )

    lookup = {
        normalize(column): column
        for column in df.columns
    }

    for candidate in candidates:
        match = lookup.get(normalize(candidate))
        if match is not None:
            return match

    raise KeyError(
        f"Could not find the {description} column.\n"
        f"Available columns: {list(df.columns)}\n"
        f"Accepted names: {candidates}"
    )


def clean_site_names(df):
    df["Site"] = (
        df["Site"]
        .astype("string")
        .str.strip()
        .str.upper()
    )
    return df


# --------------------------------------------------
# Main
# --------------------------------------------------

def main():
    section("H100 FORECAST ANALYTICS")

    if not DEBUG_FILE.exists():
        raise FileNotFoundError(
            f"Debug workbook not found:\n{DEBUG_FILE}\n"
            "Run generator.py first."
        )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    tickets = normalize_columns(
        pd.read_excel(
            DEBUG_FILE,
            sheet_name="Deduplicated Tickets",
        )
    )

    forecast = normalize_columns(
        pd.read_excel(
            DEBUG_FILE,
            sheet_name="Forecast",
        )
    )

    tickets["Date_Opened"] = pd.to_datetime(
        tickets["Date_Opened"],
        errors="coerce",
    )

    tickets["Failed_GPUs"] = pd.to_numeric(
        tickets["Failed_GPUs"],
        errors="coerce",
    )

    tickets = clean_site_names(tickets)
    forecast = clean_site_names(forecast)

    tickets["Status"] = (
        tickets["Status"]
        .fillna("Unknown")
        .astype(str)
        .str.strip()
    )

    tickets = tickets[
        tickets["Site"].isin(TARGET_SITES)
    ].copy()

    forecast = forecast[
        forecast["Site"].isin(TARGET_SITES)
    ].copy()

    print(f"Input workbook: {DEBUG_FILE}")
    print(f"Unique tickets: {len(tickets)}")
    print("\nTickets by site:")
    print(
        tickets.groupby("Site")
        .size()
        .reindex(TARGET_SITES, fill_value=0)
        .to_string()
    )

    # Forecast column names can vary slightly.
    allocation_col = find_column(
        forecast,
        [
            "Allocated_GPUs",
            "GPU_Allocation",
            "Allocated GPUs",
            "GPUs_Allocated",
        ],
        "allocated GPU count",
    )

    projected_col = find_column(
        forecast,
        [
            "Projected_Tickets",
            "Projected Tickets",
        ],
        "projected ticket count",
    )

    used_col = find_column(
        forecast,
        [
            "Estimated_GPUs_Used",
            "Estimated_GPU_Use",
            "Estimated GPUs Used",
            "Projected_GPUs_Used",
            "GPUs_Used",
        ],
        "estimated GPU usage",
    )

    forecast = forecast.set_index("Site").reindex(TARGET_SITES)

    allocated = pd.to_numeric(
        forecast[allocation_col], errors="coerce"
    ).fillna(0)

    projected = pd.to_numeric(
        forecast[projected_col], errors="coerce"
    ).fillna(0)

    estimated_used = pd.to_numeric(
        forecast[used_col], errors="coerce"
    ).fillna(0)

    remaining = (allocated - estimated_used).clip(lower=0)

    open_counts = (
        tickets.groupby("Site")
        .size()
        .reindex(TARGET_SITES, fill_value=0)
    )

    # --------------------------------------------------
    # Chart 1: Current workload vs. forecast capacity
    # --------------------------------------------------

    section("CHART 1: WORKLOAD VS. PROJECTED CAPACITY")

    fig, ax = plt.subplots(figsize=(10, 6))
    x = list(range(len(TARGET_SITES)))
    width = 0.35

    bars_current = ax.bar(
        [i - width / 2 for i in x],
        open_counts.values,
        width,
        label="Current At-Supermicro Tickets",
    )

    bars_projected = ax.bar(
        [i + width / 2 for i in x],
        projected.values,
        width,
        label="Projected Ticket Capacity",
    )

    ax.set_title(
        "H100 Current Workload vs. Projected Capacity"
    )
    ax.set_xlabel("Site")
    ax.set_ylabel("Number of Tickets")
    ax.set_xticks(x)
    ax.set_xticklabels(TARGET_SITES)
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    ax.set_axisbelow(True)

    ax.bar_label(bars_current, padding=3)
    ax.bar_label(bars_projected, padding=3)

    save_chart(fig, "H100_01_Workload_vs_Capacity.png")

    # --------------------------------------------------
    # Chart 2: GPU allocation, usage, and remaining supply
    # --------------------------------------------------

    section("CHART 2: GPU ALLOCATION AND USAGE")

    fig, ax = plt.subplots(figsize=(10, 6))

    bars_used = ax.bar(
        TARGET_SITES,
        estimated_used.values,
        label="Estimated GPU Usage",
    )

    bars_remaining = ax.bar(
        TARGET_SITES,
        remaining.values,
        bottom=estimated_used.values,
        label="Remaining GPU Supply",
    )

    ax.set_title("H100 GPU Allocation: Usage vs. Remaining Supply")
    ax.set_xlabel("Site")
    ax.set_ylabel("Number of GPUs")
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    ax.set_axisbelow(True)

    for i, site in enumerate(TARGET_SITES):
        ax.text(
            i,
            allocated.loc[site] + max(allocated.max() * 0.02, 1),
            f"Allocated: {allocated.loc[site]:.0f}",
            ha="center",
            fontsize=9,
        )

    ax.bar_label(
        bars_used,
        labels=[f"{value:.1f}" for value in estimated_used],
        label_type="center",
    )

    ax.bar_label(
        bars_remaining,
        labels=[f"{value:.1f}" for value in remaining],
        label_type="center",
    )

    save_chart(fig, "H100_02_GPU_Allocation_Usage.png")

    # --------------------------------------------------
    # Chart 3: Average failed GPUs per ticket
    # --------------------------------------------------

    section("CHART 3: FAILED GPUS PER TICKET")

    positive = tickets[
        tickets["Failed_GPUs"].notna()
        & (tickets["Failed_GPUs"] > 0)
    ].copy()

    stats = (
        positive.groupby("Site")["Failed_GPUs"]
        .agg(["count", "mean", "median", "sum"])
        .reindex(TARGET_SITES)
    )

    print("\nPositive failed-GPU ticket statistics:")
    print(stats.round(2).to_string())

    fig, ax = plt.subplots(figsize=(10, 6))

    means = stats["mean"].fillna(0)
    bars = ax.bar(TARGET_SITES, means.values)

    ax.set_title("Average Failed GPUs per Ticket by Site")
    ax.set_xlabel("Site")
    ax.set_ylabel("Average Failed GPUs per Ticket")
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    ax.set_axisbelow(True)

    labels = [
        f"{mean:.2f}\n(Number of Tickets ={int(count)})"
        if pd.notna(count)
        else "N/A"
        for mean, count in zip(stats["mean"], stats["count"])
    ]

    ax.bar_label(bars, labels=labels, padding=4)

    save_chart(fig, "H100_03_Average_GPUs_Per_Ticket.png")

    # --------------------------------------------------
    # Chart 4: Total recorded failed GPUs
    # --------------------------------------------------

    section("CHART 4: RECORDED FAILED-GPU WORKLOAD")

    totals = (
        positive.groupby("Site")["Failed_GPUs"]
        .sum()
        .reindex(TARGET_SITES, fill_value=0)
    )

    fig, ax = plt.subplots(figsize=(10, 6))

    bars = ax.bar(TARGET_SITES, totals.values)

    ax.set_title("Total Recorded Failed GPUs by Site")
    ax.set_xlabel("Site")
    ax.set_ylabel("Recorded Failed GPUs")
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    ax.set_axisbelow(True)

    ax.bar_label(
        bars,
        labels=[f"{value:.0f}" for value in totals.values],
        padding=4,
    )

    save_chart(fig, "H100_04_Total_Failed_GPUs.png")

    # --------------------------------------------------
    # Chart 5: Ticket opening timeline
    # --------------------------------------------------

    section("CHART 5: TICKET OPENING TIMELINE")

    dated = tickets.dropna(subset=["Date_Opened"]).copy()

    if dated.empty:
        print("Skipping timeline: no valid opening dates.")
    else:
        dated["Month"] = (
            dated["Date_Opened"]
            .dt.to_period("M")
            .dt.to_timestamp()
        )

        monthly = (
            dated.groupby(["Month", "Site"])
            .size()
            .unstack(fill_value=0)
            .reindex(columns=TARGET_SITES, fill_value=0)
            .sort_index()
        )

        print("\nMonthly tickets opened:")
        print(monthly.to_string())

        fig, ax = plt.subplots(figsize=(11, 6))

        for site in TARGET_SITES:
            ax.plot(
                monthly.index,
                monthly[site],
                marker="o",
                label=site,
            )

        ax.set_title("H100 Tickets by Month Opened")
        ax.set_xlabel("Month Opened")
        ax.set_ylabel("Number of Tickets")
        ax.xaxis.set_major_formatter(
            mdates.DateFormatter("%b %Y")
        )
        ax.legend(title="Site")
        ax.grid(True, linestyle="--", alpha=0.3)

        fig.autofmt_xdate()

        save_chart(fig, "H100_05_Ticket_Opening_Timeline.png")

    # --------------------------------------------------
    # Chart 6: Failed-GPU data quality
    # --------------------------------------------------

    section("CHART 6: FAILED-GPU DATA QUALITY")

    quality = pd.DataFrame({
        "Positive": (
            tickets["Failed_GPUs"].notna()
            & (tickets["Failed_GPUs"] > 0)
        ).groupby(tickets["Site"]).sum(),

        "Zero": (
            tickets["Failed_GPUs"] == 0
        ).groupby(tickets["Site"]).sum(),

        "Missing": (
            tickets["Failed_GPUs"].isna()
        ).groupby(tickets["Site"]).sum(),

        "Negative": (
            tickets["Failed_GPUs"] < 0
        ).groupby(tickets["Site"]).sum(),
    }).reindex(TARGET_SITES, fill_value=0)

    print("\nFailed-GPU data quality:")
    print(quality.to_string())

    fig, ax = plt.subplots(figsize=(10, 6))

    quality.plot(
        kind="bar",
        ax=ax,
    )

    ax.set_title("Failed-GPU Data Quality by Site")
    ax.set_xlabel("Site")
    ax.set_ylabel("Number of Tickets")
    ax.tick_params(axis="x", rotation=0)
    ax.legend(title="GPU Count Category")
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    ax.set_axisbelow(True)

    save_chart(fig, "H100_06_Failed_GPU_Data_Quality.png")

    # --------------------------------------------------
    # Complete
    # --------------------------------------------------

    section("COMPLETE")
    print("Generated up to six H100 forecast charts.")
    print(f"Output directory:\n{OUTPUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
