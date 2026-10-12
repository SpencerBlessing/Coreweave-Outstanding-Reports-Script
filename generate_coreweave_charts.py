"""
Generate charts from the CoreWeave Outstanding Cases Excel workbook.

Usage:
  py generate_coreweave_charts.py
  py generate_coreweave_charts.py "CoreWeave Outstanding Cases 10092026.xlsx"

Requirements:
  py -m pip install pandas matplotlib openpyxl

The charts are saved to a "Charts" folder next to the Excel workbook.
"""
from pathlib import Path
import sys
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime

PALETTE = [
    "#0072B2",  # Blue
    "#E69F00",  # Orange
    "#009E73",  # Green
    "#D55E00",  # Vermilion
    "#CC79A7",  # Pink
    "#56B4E9",  # Sky blue
    "#F0E442",  # Yellow
    "#332288",  # Indigo
    "#882255",  # Wine
    "#44AA99",  # Teal
    "#AA4499",  # Purple
    "#999933",  # Olive
    "#CBFFCE",  # Black
    "#A6761D",  # Brown
    "#1B9E77",  # Dark green
    "#7570B3",  # Lavender
    "#E7298A",  # Magenta
    "#66A61E",  # Lime green
    "#E6AB02",  # Gold
    "#A6761D",  # Brown
]

def pie_chart(series, title, output_path):
    series = series.dropna()

    if series.empty:
        print(f"Skipping empty chart: {title}")
        return

    fig, ax = plt.subplots(figsize=(8, 6))

    colors=[PALETTE[i % len(PALETTE)] for i in range(len(series))]

    wedges, texts, autotexts = ax.pie(
        series.values,
        labels=series.index,
        autopct=lambda pct: f"{int(round(pct * sum(series.values) / 100))}",
        startangle=90,
        colors=colors,
        labeldistance=1.25,
        pctdistance=0.65
    )

    # Draw a connector from each slice to its label
    for wedge, text in zip(wedges, texts):
        angle = np.deg2rad((wedge.theta1 + wedge.theta2) / 2)
        x, y = np.cos(angle), np.sin(angle)

        ax.plot(
            [0.95 * x, 1.2 * x],
            [0.95 * y, 1.2 * y],
            color="gray",
            linewidth=1
        )

    ax.legend(
        wedges,
        series.index,
        title=(
            "Status" if "Status Breakdown" in title
            else "Region" if "Region" in title
            else "Case Age"
        ),
        loc="upper left",
        bbox_to_anchor=(1.3, 1.0),
        frameon=False,
        labelspacing=0.5,
        handletextpad=0.5,
        borderaxespad=0.25
    )

    ax.set_title(title, pad=20)
    ax.axis("equal")

    # Reserve space for the title and legend
    fig.subplots_adjust(
        top=0.82,
        bottom=0.05,
        left=0.05,
        right=0.75
    )

    fig.savefig(output_path, dpi=160, bbox_inches="tight")
    plt.close(fig)

    print(f"Saved: {output_path}")

def save_bar(series, title, xlabel, ylabel, output_path,
             horizontal=False, limit=None):

    series = series.dropna()

    if limit:
        series = series.head(limit)

    if series.empty:
        print(f"Skipping empty chart: {title}")
        return

    fig, ax = plt.subplots(figsize=(11, 6))

    if horizontal:
        series = series.sort_values()
        series.plot(kind="barh", ax=ax, color=PALETTE[0])
    else:
        series.plot(kind="bar", ax=ax, color=PALETTE[0])
        ax.tick_params(axis="x", labelrotation=35)

    # Add numbers inside the bars
    for container in ax.containers:
       ax.bar_label(
            container,
            label_type="edge",
            padding=4,
            color="black",
            fontsize=10,
            fmt="%.0f"
        )

    ax.margins(x=0.15)

    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)

    if "Status" in title:
        total_cases = int(series.sum())
        ax.text(
            0.98, 0.02,
            f"Total Cases: {total_cases}",
            transform=ax.transAxes,
            ha="right",
            va="bottom",
            fontsize=12,
            fontweight="bold"
        )

    fig.tight_layout()
    fig.savefig(output_path, dpi=160)
    plt.close(fig)

    print(f"Saved: {output_path}")

def age_status_pies(main_df, output_dir):
    age_order = [
        "A · Open ≤ 9 days",
        "B · Open 10–15 days",
        "C · Open 16–30 days",
        "D · Open > 30 days",
    ]

    for age in age_order:
        status_counts = (
    main_df.loc[main_df["Case Age"] == age, "Status"]
    .value_counts()
    )

        # Keep the top 4 statuses; combine the rest into Other
        if len(status_counts) > 4:
            other_count = status_counts.iloc[4:].sum()
            status_counts = status_counts.iloc[:4].copy()
            status_counts["Other"] = other_count
            
        if status_counts.empty:
            continue

        pie_chart(
            status_counts,
            f"Status Breakdown - {age}",
            output_dir / f"status_by_{age[0].lower()}.png"
        )

def part_number_chart(main_df, output_path):
    # Split part numbers and keep their associated statuses.
    part_data = main_df[["Faulty Part Number", "Status"]].dropna(
        subset=["Faulty Part Number", "Status"]
    ).copy()

    part_data["Faulty Part Number"] = (
        part_data["Faulty Part Number"].astype(str).str.split(",")
    )
    part_data = part_data.explode("Faulty Part Number")
    part_data = part_data.reset_index(drop=True)
    part_data["Faulty Part Number"] = (
        part_data["Faulty Part Number"].str.strip()
    )
    part_data = part_data[part_data["Faulty Part Number"] != ""]

    # Find the top 10 part numbers by total cases.
    top_parts = part_data["Faulty Part Number"].value_counts().head(10).index

    chart_data = pd.crosstab(
        part_data["Faulty Part Number"], part_data["Status"]
    ).reindex(top_parts)

    # Sort by total count, smallest at bottom and largest at top.
    chart_data = chart_data.loc[
        chart_data.sum(axis=1).sort_values().index
    ]

    fig, ax = plt.subplots(figsize=(12, 7))

    chart_data.plot(
        kind="barh",
        stacked=True,
        ax=ax,
        colormap="tab20"
    )

    # Show total case count at the end of each bar.
    totals = chart_data.sum(axis=1)
    for y, total in enumerate(totals):
        ax.text(total + max(totals) * 0.01, y, str(int(total)),
                va="center", fontsize=9)

    ax.set_title("Parts Summary",
                 fontsize=15, fontweight="bold", pad=15)
    ax.set_xlabel("Number of Cases")
    ax.set_ylabel("Faulty Part Number")
    ax.set_axisbelow(True)
    ax.grid(axis="x", linestyle="--", alpha=0.25)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.legend(title="Status", bbox_to_anchor=(1.02, 1), loc="upper left")
    ax.margins(x=0.15)

    fig.tight_layout()
    fig.savefig(output_path, dpi=180, bbox_inches="tight")
    plt.close(fig)

    print(f"Saved: {output_path}")

def age_by_status_chart(main_df, output_path):
    age_order = [
        "Open ≤ 9 days",
        "Open 10–15 days",
        "Open 16–30 days",
        "Open > 30 days",
    ]

    chart_data = pd.crosstab(main_df["Case Age"], main_df["Status"])
    chart_data = chart_data.reindex(age_order).fillna(0)

    colors = [
        PALETTE[i % len(PALETTE)]
        for i in range(len(chart_data.columns))
    ]

    fig, ax = plt.subplots(figsize=(12, 7))

    chart_data.plot(
        kind="bar",
        stacked=True,
        ax=ax,
        color=colors
    )

    # Label only segments with more than 100 cases
    for i, age_group in enumerate(chart_data.index):
        bottom = 0

        for status in chart_data.columns:
            value = chart_data.loc[age_group, status]
            midpoint = bottom + value / 2

            if value > 100:
                ax.annotate(
                    f"{status}: {int(value)}",
                    xy=(i, midpoint),
                    xytext=(5, 0),
                    textcoords="offset points",
                    ha="center",
                    va="center",
                    fontsize=8,
                    arrowprops=dict(
                        arrowstyle="-",
                        color="gray",
                        linewidth=0.7
                    )
                )

            bottom += value

    ax.set_title(
        "Outstanding Cases by Age and Status",
        fontsize=16,
        fontweight="bold",
        color="#183247",
        pad=18
    )
    ax.set_xlabel("Case Age", fontsize=11, labelpad=10)
    ax.set_ylabel("Number of Cases", fontsize=11, labelpad=10)
    ax.tick_params(axis="x", rotation=0)

    ax.set_axisbelow(True)
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    ax.legend(
        title="Status",
        loc="center left",
        bbox_to_anchor=(1.05, 0.5),
        frameon=False,
        labelspacing=1.5,
        handletextpad=0.8,
        borderaxespad=0.5
    )

    # Reserve space on the right for labels and legend
    fig.subplots_adjust(right=1.1)

    fig.savefig(
        output_path,
        dpi=180,
        bbox_inches="tight"
    )
    plt.close(fig)

    print(f"Saved: {output_path}")

def open_cases_over_time_chart(main_df, output_path):
    # Count currently outstanding cases by creation date.
    df = main_df.copy()

    df["Case Creation Date"] = pd.to_datetime(
        df["Case Creation Date"], errors="coerce"
    )

    df = df.dropna(subset=["Case Creation Date"])
    df = df[df["Case Creation Date"] >= "2023-01-01"]

    daily_counts = (
        df.groupby(df["Case Creation Date"].dt.date)
        .size()
        .sort_index()
    )

    if daily_counts.empty:
        print("Skipping empty chart: Open Cases Over Time")
        return

    fig, ax = plt.subplots(figsize=(12, 6))

    ax.plot(
        daily_counts.index,
        daily_counts.values,
        marker="o",
        markersize=3,
        linewidth=2,
        color=PALETTE[0]
    )

    ax.set_title(
        "Currently Outstanding Cases by Creation Date",
        fontsize=15,
        fontweight="bold",
        pad=15
    )
    ax.set_xlabel("Case Creation Date")
    ax.set_ylabel("Number of Outstanding Cases")
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(output_path, dpi=180, bbox_inches="tight")
    plt.close(fig)

    print(f"Saved: {output_path}")

def newly_added_cases_chart(current_df, previous_df, output_path):
    # Identify cases present in the current report but not the previous one.
    current_cases = set(current_df["Case Number"].dropna())
    previous_cases = set(previous_df["Case Number"].dropna())

    new_cases = current_df[
        current_df["Case Number"].isin(current_cases - previous_cases)
    ]

    if new_cases.empty:
        print("No newly added outstanding cases found.")
        return

    region_counts = new_cases["Region"].fillna("Unknown").value_counts()

    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.barh(
        region_counts.sort_values().index,
        region_counts.sort_values().values,
        color=PALETTE[0]
    )

    ax.bar_label(bars, padding=4, fmt="%.0f")
    ax.set_title("Newly Added Outstanding Cases by Region")
    ax.set_xlabel("Number of New Cases")
    ax.set_ylabel("Region")
    ax.grid(axis="x", linestyle="--", alpha=0.25)
    ax.set_axisbelow(True)

    fig.tight_layout()
    fig.savefig(output_path, dpi=180, bbox_inches="tight")
    plt.close(fig)

    print(f"Saved: {output_path}")

def get_two_latest_reports(script_dir):
    import re
    from datetime import datetime

    pattern = re.compile(
        r"CoreWeave Outstanding Cases (\d{8})\.xlsx$",
        re.IGNORECASE
    )

    dated_files = []

    for file in script_dir.glob("CoreWeave Outstanding Cases *.xlsx"):
        match = pattern.fullmatch(file.name)

        if match:
            try:
                report_date = datetime.strptime(
                    match.group(1), "%m%d%Y"
                )
                dated_files.append((report_date, file))
            except ValueError:
                continue

    dated_files.sort(key=lambda item: item[0])

    if len(dated_files) < 2:
        raise FileNotFoundError(
            f"Need at least two dated report files in {script_dir}"
        )

    previous_report = dated_files[-2][1]
    current_report = dated_files[-1][1]

    print(f"Older report: {previous_report.name}")
    print(f"Newer report: {current_report.name}")

    return previous_report, current_report

def main():
    script_dir = Path(__file__).resolve().parent

    reports_dir = script_dir / "Reports"

    if len(sys.argv) > 1:
        workbook = Path(sys.argv[1])
        if not workbook.is_absolute():
            workbook = script_dir / workbook
    else:
        reports_dir = script_dir / "Reports"
        workbooks = list(
            reports_dir.glob("CoreWeave Outstanding Cases *.xlsx")
        )

        if not workbooks:
            raise FileNotFoundError(
                f"No matching Excel workbooks found in {reports_dir}"
            )

        workbook = max(workbooks, key=lambda p: p.stat().st_mtime)
        
    if not workbook.exists():
        raise FileNotFoundError(
            f"Could not find workbook: {workbook}\n"
            "Put the Excel file beside this script or pass its full path as an argument."
        )

    output_dir = workbook.parent / "Charts"
    output_dir.mkdir(exist_ok=True)

    main_df = pd.read_excel(workbook, sheet_name="CoreWeave Outstanding Cases")

    # Age chart
    age_order = [
        "A · Open ≤ 9 days",
        "B · Open 10–15 days",
        "C · Open 16–30 days",
        "D · Open > 30 days",
    ]

    age_counts = main_df["Case Age"].value_counts().reindex(age_order).dropna()

    pie_chart(
        age_counts,
        "Outstanding Cases by Age",
        output_dir / "cases_by_age.png"
    )

    # Status chart
    save_bar(
        main_df["Status"].value_counts(),
        "Outstanding Cases by Status",
        "Number of Cases",
        "Status",
        output_dir / "cases_by_status.png",
        horizontal=True,
        limit=12
    )

    age_status_pies(main_df, output_dir)

    age_by_status_chart(
        main_df,
        output_dir / "cases_by_age_and_status.png"
    )

    # Region chart
    region_counts = main_df["Region"].value_counts()

    pie_chart(
        region_counts,
        "Outstanding Cases by Region",
        output_dir / "cases_by_region_pie.png"
    )


    # Faulty part number chart
    part_number_chart(
        main_df,
        output_dir / "cases_by_part_number.png"
    )

    #  Cases over time chard
    open_cases_over_time_chart(
    main_df,
    output_dir / "open_cases_over_time.png"
)

if __name__ == "__main__":
    main()