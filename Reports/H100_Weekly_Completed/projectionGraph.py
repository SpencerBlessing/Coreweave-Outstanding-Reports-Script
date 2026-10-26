
from pathlib import Path
from datetime import datetime, timedelta

import pandas as pd
import matplotlib.pyplot as plt


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = (
    BASE_DIR / "Output" / "H100_Forecast_Debug.xlsx"
)

OUTPUT_DIR = BASE_DIR / "Output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

CHART_FILE = OUTPUT_DIR / "H100_Replacement_Projection.png"
DATA_FILE = OUTPUT_DIR / "H100_Weekly_Projection.xlsx"

# Customize these settings.
START_DATE = "2026-10-26"  # YYYY-MM-DD
WEEKLY_TICKET_TARGET = 300

SITES = ["SBN1", "PDX2", "ORD3"]

MAX_WEEKS = 104
EPSILON = 1e-9


# ============================================================
# LOAD FORECAST DATA
# ============================================================

def load_forecast():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Forecast workbook not found:\n{INPUT_FILE}\n"
            "Run generator.py first."
        )

    df = pd.read_excel(
        INPUT_FILE,
        sheet_name="Forecast Data"
    )

    required_columns = {
        "Site",
        "Historical_Tickets",
        "Average_GPUs_Per_Ticket",
    }

    missing = required_columns - set(df.columns)

    if missing:
        raise ValueError(
            f"Missing columns in Forecast Data: {sorted(missing)}"
        )

    df["Site"] = df["Site"].astype(str).str.strip().str.upper()

    # Exclude the TOTAL summary row.
    df = df[df["Site"].isin(SITES)].copy()

    if set(df["Site"]) != set(SITES):
        raise ValueError(
            f"Expected sites {SITES}; found {df['Site'].tolist()}"
        )

    df["Historical_Tickets"] = pd.to_numeric(
        df["Historical_Tickets"], errors="raise"
    )

    df["Average_GPUs_Per_Ticket"] = pd.to_numeric(
        df["Average_GPUs_Per_Ticket"], errors="raise"
    )

    if (df["Historical_Tickets"] < 0).any():
        raise ValueError("Ticket counts cannot be negative.")

    if (df["Average_GPUs_Per_Ticket"] <= 0).any():
        raise ValueError("Average GPUs per ticket must be positive.")

    return df.set_index("Site").loc[SITES].reset_index()



def load_gpu_supply():
    xlsx = pd.ExcelFile(INPUT_FILE)

    print("\nWorkbook sheets:", xlsx.sheet_names)

    inputs = pd.read_excel(INPUT_FILE, sheet_name="Forecast Inputs")
    print("\nForecast Inputs columns:", inputs.columns.tolist())
    print(inputs.head().to_string(index=False))

    # Find the total GPU supply from Forecast Inputs.
    for _, row in inputs.iterrows():
        for column in inputs.columns:
            normalized = str(column).strip().lower().replace(" ", "_")

            if normalized in {
                "available_gpus",
                "gpu_supply",
                "total_gpu_supply",
                "total_gpus",
            }:
                value = pd.to_numeric(row[column], errors="coerce")
                if pd.notna(value) and value >= 0:
                    return None, float(value)

    # If the supply is not stored in Forecast Inputs,
    # inspect the Forecast sheet for the allocation totals.
    forecast = pd.read_excel(INPUT_FILE, sheet_name="Forecast")
    print("\nForecast columns:", forecast.columns.tolist())

    for column in forecast.columns:
        normalized = str(column).strip().lower().replace(" ", "_")

        if normalized in {
            "gpus_allocated",
            "available_gpus",
            "total_gpus",
        }:
            values = pd.to_numeric(forecast[column], errors="coerce")
            site_rows = forecast[
                forecast["Site"].astype(str).str.upper().isin(SITES)
            ]

            if not site_rows.empty:
                return None, float(values.loc[site_rows.index].sum())

    raise ValueError(
        "Could not identify the available GPU supply. "
        "Review the printed workbook columns above."
    )

# ============================================================
# BUILD WEEKLY PROJECTION
# ============================================================

def build_projection(df, available_gpus):
    start_date = datetime.strptime(START_DATE, "%Y-%m-%d").date()

    # The report contains the open ticket backlog for these sites.
    remaining_tickets = {
        row["Site"]: float(row["Historical_Tickets"])
        for _, row in df.iterrows()
    }

    averages = {
        row["Site"]: float(row["Average_GPUs_Per_Ticket"])
        for _, row in df.iterrows()
    }

    # Split available GPUs as evenly as possible.
    base_allocation = available_gpus // len(SITES)
    remainder = available_gpus - base_allocation * len(SITES)

    remaining_gpus = {
        site: base_allocation + (1 if i < remainder else 0)
        for i, site in enumerate(SITES)
    }

    rows = []
    week = 0

    while week < MAX_WEEKS:
        total_backlog = sum(remaining_tickets.values())
        total_supply = sum(remaining_gpus.values())

        # Stop when all tickets are complete or no supply remains.
        if total_backlog <= EPSILON or total_supply <= EPSILON:
            break

        # Estimate each site's share of the weekly ticket target
        # based on its proportion of the remaining backlog.
        total_target = min(WEEKLY_TICKET_TARGET, total_backlog)

        weekly_target = {
            site: (
                total_target * remaining_tickets[site] / total_backlog
            )
            for site in SITES
        }

        completed = {}
        gpu_used = {}

        for site in SITES:
            average = averages[site]

            # Estimated completions can be fractional because this
            # is a forecast based on averages, not exact ticket data.
            tickets = min(
                remaining_tickets[site],
                weekly_target[site],
                remaining_gpus[site] / average,
            )

            tickets = max(0.0, tickets)
            used = tickets * average

            completed[site] = tickets
            gpu_used[site] = used

        # If no site can make further progress, stop.
        if sum(completed.values()) <= EPSILON:
            break

        week += 1
        week_start = start_date + timedelta(weeks=week - 1)

        for site in SITES:
            remaining_tickets[site] = max(
                0.0,
                remaining_tickets[site] - completed[site]
            )

            remaining_gpus[site] = max(
                0.0,
                remaining_gpus[site] - gpu_used[site]
            )

        row = {
            "Week": week,
            "Week_Start": week_start,
        }

        for site in SITES:
            row[f"{site}_Tickets_Completed"] = completed[site]
            row[f"{site}_GPUs_Used"] = gpu_used[site]
            row[f"{site}_Tickets_Remaining"] = remaining_tickets[site]
            row[f"{site}_GPUs_Remaining"] = remaining_gpus[site]

        row["Total_Tickets_Completed"] = sum(completed.values())
        row["Total_Tickets_Remaining"] = sum(remaining_tickets.values())
        row["Total_GPUs_Used"] = sum(gpu_used.values())
        row["Total_GPUs_Remaining"] = sum(remaining_gpus.values())

        rows.append(row)

    if not rows:
        return pd.DataFrame()

    return pd.DataFrame(rows)


# ============================================================
# GENERATE CHART
# ============================================================


def generate_chart(projection, df, available_gpus):
    if projection.empty:
        print("No projection generated: no tickets or GPUs available.")
        return

    import numpy as np
    import matplotlib.dates as mdates
    from matplotlib.ticker import MaxNLocator

    # --------------------------------------------------------
    # SUMMARY VALUES
    # --------------------------------------------------------
    total_tickets = int(df["Historical_Tickets"].sum())
    total_gpus_used = float(projection["Total_GPUs_Used"].sum())
    total_gpus_remaining = float(
        projection.iloc[-1]["Total_GPUs_Remaining"]
    )

    # Read total tickets from the generator's input metrics.
    invalid_tickets = 0
    total_report_tickets = total_tickets

    try:
        inputs = pd.read_excel(
            INPUT_FILE,
            sheet_name="Forecast Inputs"
        )
        metrics = dict(zip(inputs["Metric"], inputs["Value"]))
        total_report_tickets = int(
            metrics.get("Total_Tickets", total_tickets)
        )
        invalid_tickets = max(
            0, total_report_tickets - total_tickets
        )
    except Exception:
        pass

    # Week zero represents the opening backlog.
    dates = [
        pd.to_datetime(START_DATE) - pd.Timedelta(days=1)
    ] + list(pd.to_datetime(projection["Week_Start"]))

    total_remaining = [
        float(total_tickets)
    ] + list(projection["Total_Tickets_Remaining"].astype(float))

    site_remaining = {}

    for site in SITES:
        site_remaining[site] = [
            float(
                df.loc[df["Site"] == site, "Historical_Tickets"].iloc[0]
            )
        ] + list(
            projection[f"{site}_Tickets_Remaining"].astype(float)
        )

    # --------------------------------------------------------
    # CLEAN DASHBOARD LAYOUT
    # --------------------------------------------------------
    fig = plt.figure(figsize=(12, 7.5), facecolor="white")

    gs = fig.add_gridspec(
        3, 2,
        height_ratios=[0.75, 2.2, 1.5],
        hspace=0.55,
        wspace=0.25,
    )

    fig.suptitle(
        "H100 Replacement Forecast",
        fontsize=17,
        fontweight="bold",
        x=0.06,
        ha="left",
        y=0.97,
    )

    # --------------------------------------------------------
    # KPI ROW
    # --------------------------------------------------------
    kpi_ax = fig.add_subplot(gs[0, :])
    kpi_ax.axis("off")

    kpis = [
        ("OPEN TICKETS", f"{total_tickets:,}"),
        ("GPU SUPPLY", f"{available_gpus:,.0f}"),
        ("EST. GPU USE", f"{total_gpus_used:,.0f}"),
        ("GPUs REMAINING", f"{total_gpus_remaining:,.0f}"),
    ]

    for x, (label, value) in zip(
        [0.12, 0.37, 0.62, 0.87], kpis
    ):
        kpi_ax.text(
            x, 0.67, value,
            transform=kpi_ax.transAxes,
            ha="center", va="center",
            fontsize=19,
            fontweight="bold",
        )
        kpi_ax.text(
            x, 0.20, label,
            transform=kpi_ax.transAxes,
            ha="center", va="center",
            fontsize=8,
            color="dimgray",
        )

    # --------------------------------------------------------
    # CHART 1: TOTAL BACKLOG BURN-DOWN
    # --------------------------------------------------------
    ax1 = fig.add_subplot(gs[1, 0])

    ax1.plot(
        dates,
        total_remaining,
        marker="o",
        linewidth=2.5,
        markersize=5,
        label="Open tickets",
    )

    ax1.fill_between(
        dates,
        total_remaining,
        alpha=0.08,
    )

    ax1.set_title(
        "Total Backlog",
        loc="left",
        fontsize=11,
        fontweight="bold",
    )
    ax1.set_ylabel("Tickets remaining")
    ax1.set_ylim(bottom=0)
    ax1.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax1.grid(axis="y", alpha=0.2)
    ax1.set_axisbelow(True)

    ax1.annotate(
        f"{total_remaining[0]:.0f}",
        (dates[0], total_remaining[0]),
        xytext=(4, 6),
        textcoords="offset points",
        fontsize=9,
        fontweight="bold",
    )

    ax1.annotate(
        f"{total_remaining[-1]:.0f}",
        (dates[-1], total_remaining[-1]),
        xytext=(-4, 6),
        textcoords="offset points",
        ha="right",
        fontsize=9,
        fontweight="bold",
    )

    # --------------------------------------------------------
    # CHART 2: SITE BACKLOG BURN-DOWN
    # --------------------------------------------------------
    ax2 = fig.add_subplot(gs[1, 1])

    for site in SITES:
        ax2.plot(
            dates,
            site_remaining[site],
            marker="o",
            linewidth=2,
            markersize=4,
            label=site,
        )

    ax2.set_title(
        "Remaining Tickets by Site",
        loc="left",
        fontsize=11,
        fontweight="bold",
    )
    ax2.set_ylabel("Tickets remaining")
    ax2.set_ylim(bottom=0)
    ax2.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax2.grid(axis="y", alpha=0.2)
    ax2.set_axisbelow(True)
    ax2.legend(
        frameon=False,
        ncol=3,
        loc="upper right",
        fontsize=8,
    )

    # Use concise date labels on both charts.
    for ax in (ax1, ax2):
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %d"))
        ax.tick_params(axis="x", labelrotation=25, labelsize=8)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_alpha(0.25)
        ax.spines["bottom"].set_alpha(0.25)

    # --------------------------------------------------------
    # BOTTOM LEFT: GPU SUPPLY
    # --------------------------------------------------------
    ax3 = fig.add_subplot(gs[2, 0])

    ax3.set_title(
        "GPU Supply",
        loc="left",
        fontsize=11,
        fontweight="bold",
    )

    if available_gpus > 0:
        used_pct = total_gpus_used / available_gpus * 100
        remaining_pct = total_gpus_remaining / available_gpus * 100

        ax3.barh(
            ["Available GPUs"],
            [total_gpus_used],
            label="Estimated use",
        )
        ax3.barh(
            ["Available GPUs"],
            [total_gpus_remaining],
            left=[total_gpus_used],
            label="Remaining",
        )

        ax3.text(
            total_gpus_used / 2,
            0,
            f"{total_gpus_used:.0f} used",
            ha="center",
            va="center",
            fontsize=9,
        )

        if total_gpus_remaining > 0:
            ax3.text(
                total_gpus_used + total_gpus_remaining / 2,
                0,
                f"{total_gpus_remaining:.0f} left",
                ha="center",
                va="center",
                fontsize=9,
            )

        ax3.set_xlim(0, available_gpus * 1.05)
        ax3.legend(
            frameon=False,
            loc="lower center",
            bbox_to_anchor=(0.5, -0.40),
            ncol=2,
            fontsize=8,
        )
    else:
        ax3.text(
            0.5, 0.5,
            "No GPU supply available",
            transform=ax3.transAxes,
            ha="center",
            va="center",
        )

    ax3.set_xlabel("GPUs")
    ax3.grid(axis="x", alpha=0.2)
    ax3.set_axisbelow(True)
    ax3.spines["top"].set_visible(False)
    ax3.spines["right"].set_visible(False)

    # --------------------------------------------------------
    # BOTTOM RIGHT: SITE SUMMARY TABLE
    # --------------------------------------------------------
    ax4 = fig.add_subplot(gs[2, 1])
    ax4.axis("off")

    ax4.set_title(
        "Site Summary",
        loc="left",
        fontsize=11,
        fontweight="bold",
        pad=10,
    )

    table_data = []

    for site in SITES:
        row = df.loc[df["Site"] == site].iloc[0]

        table_data.append([
            site,
            f"{row['Historical_Tickets']:.0f}",
            f"{projection[f'{site}_Tickets_Completed'].sum():.1f}",
            f"{projection.iloc[-1][f'{site}_Tickets_Remaining']:.1f}",
            f"{row['Average_GPUs_Per_Ticket']:.2f}",
        ])

    table = ax4.table(
        cellText=table_data,
        colLabels=[
            "Site",
            "Opening",
            "Done",
            "Left",
            "GPU/Ticket",
        ],
        cellLoc="center",
        loc="center",
        colWidths=[0.16, 0.17, 0.17, 0.17, 0.25],
    )

    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1.0, 1.7)

    for (row, col), cell in table.get_celld().items():
        cell.set_linewidth(0.4)
        cell.set_edgecolor("#DDDDDD")

        if row == 0:
            cell.set_text_props(fontweight="bold")
            cell.set_facecolor("#F1F3F5")
        else:
            cell.set_facecolor("white")

    # --------------------------------------------------------
    # FOOTNOTE
    # --------------------------------------------------------
    note = (
        f"Projection starts {START_DATE} | "
        "Based on historical average GPUs per ticket."
    )

    if invalid_tickets:
        note += (
            f" Excludes {invalid_tickets} tickets with invalid GPU counts."
        )

    fig.text(
        0.06, 0.025,
        note,
        ha="left",
        fontsize=8,
        color="dimgray",
    )

    fig.subplots_adjust(
        top=0.90,
        bottom=0.13,
        left=0.08,
        right=0.97,
    )

    fig.savefig(
        CHART_FILE,
        dpi=220,
        bbox_inches="tight",
        facecolor="white",
    )
    plt.close(fig)

    print(f"Chart saved to:\n{CHART_FILE}")


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 65)
    print("H100 WEEKLY REPLACEMENT PROJECTION")
    print("=" * 65)

    df = load_forecast()
    _, available_gpus = load_gpu_supply()

    # Use the actual site ticket counts from the report as backlog.
    total_open_tickets = int(df["Historical_Tickets"].sum())

    print(f"Start date:          {START_DATE}")
    print(f"Open tickets:        {total_open_tickets}")
    print(f"Available GPUs:      {available_gpus:,.0f}")
    print(f"Weekly ticket target:{WEEKLY_TICKET_TARGET:>8,}")
    print("\nSite assumptions:")

    for _, row in df.iterrows():
        site = row["Site"]
        print(
            f"  {site}: {row['Historical_Tickets']:.0f} tickets, "
            f"{row['Average_GPUs_Per_Ticket']:.2f} GPUs/ticket"
        )

    projection = build_projection(df, available_gpus)

    if projection.empty:
        print("\nNo work can be projected.")
        return

    projection.to_excel(DATA_FILE, index=False)
    generate_chart(projection, df, available_gpus)

    final = projection.iloc[-1]

    print("\n" + "=" * 65)
    print("PROJECTION SUMMARY")
    print("=" * 65)
    print(f"Weeks projected:       {len(projection)}")
    print(f"Final projected week:  {final['Week_Start']}")
    print(f"Tickets completed:     "
          f"{projection['Total_Tickets_Completed'].sum():.2f}")
    print(f"Tickets remaining:     "
          f"{final['Total_Tickets_Remaining']:.2f}")
    print(f"GPUs consumed:         "
          f"{projection['Total_GPUs_Used'].sum():.2f}")
    print(f"GPUs remaining:        "
          f"{final['Total_GPUs_Remaining']:.2f}")
    print(f"\nWeekly data saved to:\n{DATA_FILE}")


if __name__ == "__main__":
    main()
