
"""CoreWeave H100 GPU burndown.

Weekly stacked columns show GPU requirements by site.
A line overlays the total GPU burndown trajectory.
Numbers inside each segment show GPUs needed by site.

Run:
    python Reports/H100/ReportH100Burndown.py

Requires openpyxl.
"""

from __future__ import annotations

import argparse
import math
from openpyxl.drawing.colors import ColorChoice, SchemeColor
import re
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    Paragraph,
    ParagraphProperties,
    CharacterProperties,
    Font as DrawingFont,
)
from openpyxl.chart.data_source import (
    AxDataSource,
    NumData,
    NumVal,
    StrData,
    StrRef,
    StrVal,
)
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties


SOURCE_SHEET = "H100 Thermal Issues"
OUTPUT_SHEET = "Weekly Burndown"

WEEKLY_GPU_BUDGET = 400
PRIORITY_SITES = ("SBN1", "PDX2", "ORD3")
PRIORITY_SHARE = 100

INPUT_MARKER = "__H100_400_GPU_ALLOCATION_V1__"

NAVY = "17324D"
INPUT_BLUE = "E7F2FC"

SITE_COLORS = (
    "1D4ED8",
    "D97706",
    "0F766E",
    "7C3AED",
    "BE123C",
    "16A34A",
    "0E7490",
    "A16207",
    "C2410C",
    "4F46E5",
    "475569",
    "DB2777",
    "65A30D",
    "0284C7",
    "9333EA",
)

PART_SPLIT = re.compile(r"\s*\|\s*|[\r\n,]+")
PART_COUNT = re.compile(r"^(\d+)\s*[xX×]\s*(.+)$")


# ============================================================
# PROJECT PATHS
# ============================================================

def project_root() -> Path:

    for folder in (
        *Path(__file__).resolve().parents,
        Path.cwd().resolve(),
    ):
        if (
            folder / "Reports" / "Coreweave_Outstanding"
        ).is_dir():
            return folder

    return Path.cwd().resolve()


def latest_report(root: Path) -> Path:

    folder = (
        root
        / "Reports"
        / "Coreweave_Outstanding"
        / "Output"
    )

    candidates = [
        path
        for path in folder.glob(
            "CoreWeave Outstanding Cases *.xlsx"
        )
        if path.is_file()
        and not path.name.startswith("~$")
    ]

    if not candidates:
        raise FileNotFoundError(
            f"No finalized outstanding report in {folder}"
        )

    return max(
        candidates,
        key=lambda path: path.stat().st_mtime,
    )


# ============================================================
# GPU QUANTITY CALCULATION
# ============================================================

def gpu_equivalents(parts) -> int:
    total = 0

    for item in PART_SPLIT.split(str(parts or "")):
        item = item.strip()

        if not item:
            continue

        match = PART_COUNT.match(item)

        quantity = int(match.group(1)) if match else 1
        part = (match.group(2) if match else item).upper()

        if "FRU-H100" in part or "SVC-H100" in part:
            total += quantity

    return total


# ============================================================
# READ H100 REQUIREMENTS
# ============================================================

def read_gpu_requirements(report: Path):

    book = load_workbook(
        report,
        read_only=True,
        data_only=True,
    )

    try:

        if SOURCE_SHEET not in book.sheetnames:
            raise ValueError(
                f"Missing worksheet: {SOURCE_SHEET}"
            )

        rows = book[SOURCE_SHEET].iter_rows(
            values_only=True
        )

        headers = {
            str(value).strip(): index
            for index, value in enumerate(next(rows))
            if value is not None
        }

        required = (
            "Case Number",
            "Site",
            "Faulty Part Number",
        )

        missing = [
            name
            for name in required
            if name not in headers
        ]

        if missing:
            raise ValueError(
                f"Missing columns: {missing}"
            )

        case_map = {}
        conflicts = set()

        for cells in rows:

            case = str(
                cells[headers["Case Number"]] or ""
            ).strip().upper()

            if (
                not case
                or case == "TOTAL"
                or case in conflicts
            ):
                continue

            site = str(
                cells[headers["Site"]] or ""
            ).strip()

            if not site:
                continue

            gpus = gpu_equivalents(
                cells[headers["Faulty Part Number"]]
            )

            previous = case_map.get(case)

            if previous and previous[0] != site:

                del case_map[case]
                conflicts.add(case)
                continue

            case_map[case] = (
                site,
                max(
                    gpus,
                    previous[1] if previous else 0,
                ),
            )

        totals = defaultdict(int)

        for site, gpus in case_map.values():
            totals[site] += gpus

        if not totals:
            raise ValueError(
                "No H100 GPU requirements with an "
                "assigned Site found"
            )

        return sorted(
            totals.items(),
            key=lambda entry: (
                -entry[1],
                entry[0],
            ),
        )

    finally:
        book.close()


# ============================================================
# PRESERVE MANUAL INPUTS
# ============================================================

def normalize_week(value):

    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")

    if not value:
        return None

    for fmt in (
        "%Y-%m-%d",
        "%b %d, %Y",
        "%m/%d/%Y",
    ):

        try:
            return datetime.strptime(
                str(value).strip(),
                fmt,
            ).date().isoformat()

        except ValueError:
            continue

    return None


def load_inputs(path: Path):
    """Preserve GPUs Available from previous reports."""

    if not path.is_file():
        return {}

    book = load_workbook(
        path,
        read_only=True,
        data_only=False,
    )

    try:

        if OUTPUT_SHEET not in book.sheetnames:
            return {}

        if book[OUTPUT_SHEET]["Q1"].value != INPUT_MARKER:
            return {}

        saved = {}
        heading = None

        for values in book[OUTPUT_SHEET].iter_rows(
            values_only=True
        ):

            labels = {
                str(value).strip().casefold(): index
                for index, value in enumerate(values)
                if isinstance(value, str)
            }

            keys = (
                "site",
                "week starting",
                "gpus available",
            )

            if all(key in labels for key in keys):

                heading = {
                    key: labels[key]
                    for key in keys
                }

                continue

            if heading is None:
                continue

            try:

                site = str(
                    values[heading["site"]] or ""
                ).strip()

                week = normalize_week(
                    values[heading["week starting"]]
                )

                gpu = float(
                    values[heading["gpus available"]]
                )

                if (
                    site
                    and week
                    and gpu >= 0
                    and gpu.is_integer()
                ):
                    baseline = (
                        values[16]
                        if len(values) > 16
                        else None
                    )

                    manual_flag = (
                        values[17]
                        if len(values) > 17
                        else None
                    )

                    if (
                        baseline is not None
                        and (
                            int(gpu) != int(baseline)
                            or manual_flag == 1
                        )
                    ):
                        saved[(site, week)] = int(gpu)

            except (
                ValueError,
                TypeError,
                IndexError,
            ):
                pass

        return saved

    finally:
        book.close()


# ============================================================
# WEEKLY GPU ALLOCATION
# ============================================================

def distribute_evenly(
    pool,
    sites,
    remaining,
    allocated,
    week,
):
    active = sorted(
        site for site in sites
        if remaining[site] > allocated[site]
    )

    if active:
        offset = week % len(active)
        active = active[offset:] + active[:offset]

    while pool > 0 and active:
        each, extra = divmod(pool, len(active))
        used = 0

        for index, site in enumerate(active):
            amount = min(
                remaining[site] - allocated[site],
                each + int(index < extra),
            )

            allocated[site] += amount
            used += amount

        if used == 0:
            break

        pool -= used

        active = [
            site for site in active
            if remaining[site] > allocated[site]
        ]

    return pool


def make_allocation_plan(
    summary,
    start,
    weeks,
    budget,
    saved,
):
    remaining = dict(summary)
    sites = list(remaining)

    priority = [
        site for site in PRIORITY_SITES
        if site in remaining
    ]

    secondary = [
        site for site in sites
        if site not in priority
    ]

    allocations = {
        site: [] for site in sites
    }

    projected = {
        site: [] for site in sites
    }

    for week in range(weeks):
        week_date = (
            start + timedelta(weeks=week)
        ).isoformat()

        allocated = {
            site: 0 for site in sites
        }

        # Preserve explicitly edited allocations.
        for site in sites:
            projected[site].append(
                remaining[site]
            )

            key = (site, week_date)

            if key in saved:
                allocated[site] = min(
                    saved[key],
                    remaining[site],
                )

        pool = budget - sum(allocated.values())

        if pool < 0:
            raise ValueError(
                f"Manual allocations exceed the "
                f"{budget}-GPU budget for {week_date}"
            )

        # Allocate up to 100 GPUs to each priority site.
        for site in priority:
            if (site, week_date) in saved:
                continue

            amount = min(
                PRIORITY_SHARE,
                remaining[site],
                pool,
            )

            allocated[site] += amount
            pool -= amount

        # Redistribute available capacity evenly
        # among remaining secondary sites.
        pool = distribute_evenly(
            pool,
            [
                site for site in secondary
                if (site, week_date) not in saved
            ],
            remaining,
            allocated,
            week,
        )

        # If all secondary demand is satisfied,
        # remaining capacity can finish priority sites.
        pool = distribute_evenly(
            pool,
            [
                site for site in priority
                if (site, week_date) not in saved
            ],
            remaining,
            allocated,
            week,
        )

        for site in sites:
            allocations[site].append(
                allocated[site]
            )

            remaining[site] -= allocated[site]

    return allocations, projected

# ============================================================
# CREATE WORKBOOK
# ============================================================

def create_workbook(
    summary,
    start: date,
    weeks: int,
    weekly_budget: int,
    saved,
):

    allocations, planned_remaining = make_allocation_plan(
        summary,
        start,
        weeks,
        weekly_budget,
        saved,
    )

    book = Workbook()

    ws = book.active
    ws.title = OUTPUT_SHEET
    ws["Q1"] = INPUT_MARKER
    ws["R1"] = "Manual Override"

    ws.column_dimensions["Q"].hidden = True
    ws.column_dimensions["R"].hidden = True

    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 80

    # Allow normal worksheet scrolling.
    ws.freeze_panes = None

    widths = {
        "A": 21,
        "B": 22,
        "C": 22,
        "D": 23,
        "E": 22,
        "F": 23,
    }

    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    for column in "GHIJKLMNOP":
        ws.column_dimensions[column].width = 12

    # ========================================================
    # TITLE
    # ========================================================

    ws.merge_cells("A1:F2")

    ws["A1"] = (
        "COREWEAVE  |  H100 WEEKLY GPU BURNDOWN"
    )

    ws["A1"].font = Font(
        name="Aptos Display",
        size=18,
        bold=True,
        color="FFFFFF",
    )

    ws["A1"].alignment = Alignment(
        vertical="center",
        indent=1,
    )

    for row in ws["A1:F2"]:
        for cell in row:
            cell.fill = PatternFill(
                "solid",
                fgColor=NAVY,
            )

    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 26

    # ========================================================
    # SITE COLORS / LINE LEGEND
    # ========================================================

    ws["A35"] = "SITE COLORS"

    ws["A35"].font = Font(
        size=11,
        bold=True,
        color=NAVY,
    )

    ws["D35"] = "DARK LINE = TOTAL GPUs NEEDED"

    ws["D35"].font = Font(
        size=11,
        bold=True,
        color=NAVY,
    )

    ws["A39"] = (
    "DISCLAIMER: Burndown projections assume receipt of 400 GPUs "
    "per week and no new H100 thermal breakdown tickets."
    )

    ws["A39"].font = Font(
        size=10,
        italic=True,
        color="666666",
    )

    for index, (site, _) in enumerate(summary):

        row = 36 + index // 6
        column = 1 + index % 6

        color = SITE_COLORS[
            index % len(SITE_COLORS)
        ]

        cell = ws.cell(row, column, site)

        cell.fill = PatternFill(
            "solid",
            fgColor=color,
        )

        cell.font = Font(
            color="FFFFFF",
            bold=True,
            size=11,
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

        ws.row_dimensions[row].height = 26

    # ========================================================
    # GPU REQUIREMENTS SUMMARY
    # ========================================================

    ws.merge_cells("A40:F40")

    ws["A40"] = (
        "OUTSTANDING H100 GPU REQUIREMENTS BY SITE"
    )

    ws["A40"].font = Font(
        size=12,
        color=NAVY,
        bold=True,
    )

    ws["A40"].fill = PatternFill(
        "solid",
        fgColor="DCE7F0",
    )

    for col, title in enumerate(
        (
            "Site",
            "GPUs Needed",
        ),
        1,
    ):

        cell = ws.cell(42, col, title)

        cell.fill = PatternFill(
            "solid",
            fgColor=NAVY,
        )

        cell.font = Font(
            color="FFFFFF",
            bold=True,
        )

        cell.alignment = Alignment(
            vertical="center"
        )

    ws.row_dimensions[42].height = 26

    summary_rows = {}

    for index, (site, total) in enumerate(summary):

        row = 43 + index

        summary_rows[site] = row

        ws.cell(row, 1, site)

        ws.cell(
            row,
            2,
            total,
        ).number_format = "#,##0"

        ws.cell(row, 1).font = Font(
            color=SITE_COLORS[
                index % len(SITE_COLORS)
            ],
            bold=True,
        )

        if index % 2:

            for col in (1, 2):

                ws.cell(row, col).fill = PatternFill(
                    "solid",
                    fgColor="F2F6FA",
                )

    total_row = 43 + len(summary)

    ws.cell(
        total_row,
        1,
        "ALL SITES",
    )

    ws.cell(
        total_row,
        2,
        f"=SUM(B43:B{total_row - 1})",
    ).number_format = "#,##0"

    for col in (1, 2):

        ws.cell(
            total_row,
            col,
        ).fill = PatternFill(
            "solid",
            fgColor="DCE7F0",
        )

        ws.cell(
            total_row,
            col,
        ).font = Font(
            color=NAVY,
            bold=True,
        )

    # ========================================================
    # WEEKLY EDITABLE GPU TABLE
    # ========================================================

    help_row = total_row + 3

    ws.merge_cells(
        start_row=help_row,
        start_column=1,
        end_row=help_row,
        end_column=6,
    )

    ws.cell(
        help_row,
        1,
        "BLUE CELLS: edit GPUs Available per site "
        "and week. The chart shows GPUs still "
        "needed at the start of each week.",
    ).font = Font(
        size=10,
        color="526779",
    )

    ws.row_dimensions[help_row].height = 26

    header_row = help_row + 2

    headers = (
        "Site",
        "Week Starting",
        "GPUs Needed",
        "GPUs Available",
        "GPUs Applied",
        "GPUs Remaining",
    )

    for col, label in enumerate(headers, 1):

        cell = ws.cell(
            header_row,
            col,
            label,
        )

        cell.fill = PatternFill(
            "solid",
            fgColor=NAVY,
        )

        cell.font = Font(
            color="FFFFFF",
            bold=True,
        )

        cell.alignment = Alignment(
            vertical="center",
            wrap_text=True,
        )

    ws.row_dimensions[header_row].height = 31

    # ========================================================
    # WEEKLY GPU PROJECTIONS
    # ========================================================

    projected_by_site = {}
    first_rows = {}

    for index, (
        site,
        original_gpus,
    ) in enumerate(summary):

        first = (
            header_row
            + 1
            + index * weeks
        )

        first_rows[site] = first

        row_summary = summary_rows[site]

        color = SITE_COLORS[
            index % len(SITE_COLORS)
        ]

        needed_series = []

        for week in range(weeks):

            row = first + week

            week_date = (
                start
                + timedelta(weeks=week)
            )

            available = allocations[site][week]

            ws.cell(row, 1, site)

            ws.cell(
                row,
                2,
                week_date,
            ).number_format = "mmm dd, yyyy"

            # GPUs needed at the start of the week.
            ws.cell(
                row,
                3,
                (
                    f"=$B${row_summary}"
                    if week == 0
                    else f"=F{row - 1}"
                ),
            ).number_format = "#,##0"

            # Editable inventory.

            ws.cell(
                row,
                4,
                available,
            ).number_format = "#,##0"

            # Store generated allocation and manual override status.

            ws.cell(
                row,
                17,
                available,
            )

            ws.cell(
                row,
                18,
                int(
                    (site, week_date.isoformat()) in saved
                ),
            )

            # GPUs applied this week.

            ws.cell(
                row,
                5,
                f"=MIN(C{row},D{row})",
            ).number_format = "#,##0"

            # GPUs remaining after this week.
            ws.cell(
                row,
                6,
                f"=MAX(0,C{row}-E{row})",
            ).number_format = "#,##0"

            ws.cell(
                row,
                4,
            ).fill = PatternFill(
                "solid",
                fgColor=INPUT_BLUE,
            )

            if index % 2:

                for col in (
                    1,
                    2,
                    3,
                    5,
                    6,
                ):

                    ws.cell(
                        row,
                        col,
                    ).fill = PatternFill(
                        "solid",
                        fgColor="F2F6FA",
                    )

            if week == 0:

                ws.cell(
                    row,
                    1,
                ).fill = PatternFill(
                    "solid",
                    fgColor=color,
                )

                ws.cell(
                    row,
                    1,
                ).font = Font(
                    color="FFFFFF",
                    bold=True,
                )

            # Initial chart values before Excel recalculates.
            needed_series.append(
                planned_remaining[site][week]
            )

        projected_by_site[site] = needed_series

    # ========================================================
    # STACKED GPU COLUMN CHART
    # ========================================================

    chart = BarChart()

    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.gapWidth = 35

    chart.title = (
        "H100 GPUs Still Needed by Week and Site"
    )

    chart.style = 13

    # Overall chart dimensions in centimeters.
    chart.width = 31.5
    chart.height = 15.0

    # Custom site key below the chart.
    chart.legend = None

    # --------------------------------------------------------
    # GPU COUNTS INSIDE EACH COLORED SEGMENT
    # --------------------------------------------------------

    segment_font = CharacterProperties(
        latin=DrawingFont(
            typeface="Aptos"
        ),
        sz=1000,
        b=True,
        solidFill="FFFFFF",
    )

    # --------------------------------------------------------
    # X AXIS: WEEKLY DATES
    # --------------------------------------------------------

    chart.x_axis.title = "Week Starting"
    chart.x_axis.axPos = "b"

    chart.x_axis.tickLblPos = "nextTo"
    chart.x_axis.majorTickMark = "out"
    chart.x_axis.delete = False

    # --------------------------------------------------------
    # Y AXIS: TOTAL GPU REQUIREMENTS
    # --------------------------------------------------------

    chart.y_axis.title = "GPUs Needed"
    chart.y_axis.axPos = "l"

    chart.y_axis.numFmt = "#,##0"
    chart.y_axis.scaling.min = 0

    chart.y_axis.tickLblPos = "nextTo"
    chart.y_axis.majorTickMark = "out"
    chart.y_axis.delete = False

    total_gpus = sum(
        gpus
        for _, gpus in summary
    )

    if total_gpus <= 1000:
        tick = 100
    elif total_gpus <= 5000:
        tick = 500
    else:
        tick = 1000

    chart.y_axis.majorUnit = tick

    chart.y_axis.scaling.max = max(
        tick,
        math.ceil(
            total_gpus * 1.10 / tick
        ) * tick,
    )

    chart.y_axis.majorGridlines = ChartLines(
        spPr=GraphicalProperties(
            ln=LineProperties(
                solidFill="DCE3ED",
                w=6350,
            )
        )
    )

    # ========================================================
    # HIDDEN CHART DATA (SAME WORKSHEET)
    # ========================================================

    helper_col = 27  # AA

    label_col = get_column_letter(
        helper_col
    )

    labels = [
        (start + timedelta(weeks=index)).strftime("%m/%d/%Y")
        for index in range(weeks)
    ]

    ws.cell(
        1,
        helper_col,
        "Week",
    )

    for week, label in enumerate(labels):

        ws.cell(
            week + 2,
            helper_col,
            label,
        )

    # --------------------------------------------------------
    # ADD SITE SERIES TO STACKED COLUMNS
    # --------------------------------------------------------

    for index, (site, _) in enumerate(summary):

        col = helper_col + 1 + index

        ws.cell(1, col, site)

        first = first_rows[site]

        for week in range(weeks):

            ws.cell(
                week + 2,
                col,
                f"=C{first + week}",
            )

        data_ref = Reference(
            ws,
            min_col=col,
            min_row=1,
            max_row=weeks + 1,
        )

        chart.add_data(
            data_ref,
            titles_from_data=True,
        )

        series = chart.series[-1]

        color = SITE_COLORS[
            index % len(SITE_COLORS)
        ]

        series.graphicalProperties.solidFill = color

        series.graphicalProperties.line.solidFill = color

        # Show ONLY site code and GPU count for segments >= 100.
        # Never show the week or date in a bar label.
        point_labels = [
            DataLabel(
                idx=week,
                showVal=count >= 100,
                showSerName=count >= 100,
                showCatName=False,
                showLegendKey=False,
            )
            for week, count in enumerate(projected_by_site[site])
        ]

        # Full site and count labels on larger segments.
        series.dLbls = DataLabelList(
            dLbl=point_labels,
            dLblPos="ctr",
            showVal=True,
            showSerName=True,
            showCatName=False,
            showLegendKey=False,
            separator=" - ",
            numFmt="#,##0;(#,##0);",

            txPr=RichText(
                p=[
                    Paragraph(
                        pPr=ParagraphProperties(
                            defRPr=segment_font
                        ),
                        endParaRPr=segment_font,
                    )
                ]
            ),
        )

        # Cached values allow immediate first rendering.
        series.val.numRef.numCache = NumData(
            formatCode="#,##0",
            ptCount=weeks,
            pt=[
                NumVal(
                    idx=week,
                    v=str(value),
                )
                for week, value in enumerate(
                    projected_by_site[site]
                )
            ],
        )

        # Use explicit weekly date labels.
        series.cat = AxDataSource(
            strRef=StrRef(
                f=(
                    f"'{OUTPUT_SHEET}'!"
                    f"${label_col}$2:"
                    f"${label_col}${weeks + 1}"
                ),
                strCache=StrData(
                    ptCount=weeks,
                    pt=[
                        StrVal(
                            idx=week,
                            v=label,
                        )
                        for week, label in enumerate(
                            labels
                        )
                    ],
                ),
            )
        )

    # ========================================================
    # OVERLAID TOTAL GPU BURNDOWN LINE
    # ========================================================

    # The total is calculated from the same site
    # requirements used by the stacked columns.
    total_col = (
        helper_col
        + 1
        + len(summary)
    )

    ws.cell(
        1,
        total_col,
        "Total GPUs Needed",
    )

    first_site_letter = get_column_letter(
        helper_col + 1
    )

    last_site_letter = get_column_letter(
        total_col - 1
    )

    for week in range(weeks):

        row = week + 2

        ws.cell(
            row,
            total_col,
            (
                f"=SUM("
                f"{first_site_letter}{row}:"
                f"{last_site_letter}{row}"
                f")"
            ),
        )

    total_line = LineChart()

    total_line.grouping = "standard"

    total_line.add_data(
        Reference(
            ws,
            min_col=total_col,
            min_row=1,
            max_row=weeks + 1,
        ),
        titles_from_data=True,
    )

    total_series = total_line.series[0]

    line_color = "102A43"

    total_series.graphicalProperties.line.solidFill = ColorChoice(
        schemeClr=SchemeColor(
            val="dk2",
            alpha=80000,
            red=6275,
            green=16471,
            blue=26275,
        )
    )

    total_series.graphicalProperties.line.width = 25400

    total_series.marker.symbol = "none"

    total_series.marker.graphicalProperties.solidFill = (
        line_color
    )

    total_series.marker.graphicalProperties.line.solidFill = (
        line_color
    )

    # Cache combined GPU totals for immediate rendering.
    total_series.val.numRef.numCache = NumData(
        formatCode="#,##0",
        ptCount=weeks,
        pt=[
            NumVal(
                idx=week,
                v=str(
                    sum(
                        projected_by_site[site][week]
                        for site, _ in summary
                    )
                ),
            )
            for week in range(weeks)
        ],
    )

    # Match the weekly categories used by the bars.
    total_series.cat = AxDataSource(
        strRef=StrRef(
            f=(
                f"'{OUTPUT_SHEET}'!"
                f"${label_col}$2:"
                f"${label_col}${weeks + 1}"
            ),
            strCache=StrData(
                ptCount=weeks,
                pt=[
                    StrVal(
                        idx=week,
                        v=label,
                    )
                    for week, label in enumerate(
                        labels
                    )
                ],
            ),
        )
    )

    # --------------------------------------------------------
    # TOTAL GPU LABELS ABOVE THE BURNDOWN LINE
    # --------------------------------------------------------

    total_font = CharacterProperties(
        latin=DrawingFont(
            typeface="Aptos"
        ),
        sz=1000,
        b=True,
        solidFill=line_color,
    )

    total_line.dataLabels = DataLabelList(
        dLblPos="t",
        showVal=True,
        showSerName=False,
        showCatName=False,
        showLegendKey=False,
        showPercent=False,
        numFmt="#,##0",

        txPr=RichText(
            p=[
                Paragraph(
                    pPr=ParagraphProperties(
                        defRPr=total_font
                    ),
                    endParaRPr=total_font,
                )
            ]
        ),
    )

    # --------------------------------------------------------
    # USE THE SAME AXES FOR BARS AND LINE
    # --------------------------------------------------------

    total_line.x_axis.axId = chart.x_axis.axId
    total_line.y_axis.axId = chart.y_axis.axId

    # Combine into one Excel chart.
    chart += total_line

    chart.visible_cells_only = False

    ws.add_chart(
        chart,
        "A4",
    )

    # Keep helper data hidden on the existing worksheet.
    for col in range(
        helper_col,
        total_col + 1,
    ):

        ws.column_dimensions[
            get_column_letter(col)
        ].hidden = True

    book.calculation = CalcProperties(
        calcMode="auto",
        fullCalcOnLoad=True,
        forceFullCalc=True,
    )

    return book


# ============================================================
# MAIN
# ============================================================

def main(argv=None):

    parser = argparse.ArgumentParser(
        description=(
            "Weekly stacked H100 GPU burndown "
            "with total trajectory line"
        )
    )

    parser.add_argument(
        "--input",
        type=Path,
    )

    parser.add_argument(
        "--output",
        type=Path,
    )

    parser.add_argument(
        "--first-week",
        type=date.fromisoformat,
    )

    parser.add_argument(
        "--weeks",
        type=int,
        default=8,
    )

    parser.add_argument(
        "--weekly-budget",
        type=int,
        default=WEEKLY_GPU_BUDGET,
        help="Total GPUs available across all sites per week",
    )

    # Old argument accepted but no longer used.
    parser.add_argument(
        "--gpus",
        type=int,
        help=argparse.SUPPRESS,
    )

    parser.add_argument(
        "--reset-inputs",
        action="store_true",
    )

    # Accept old switches for compatibility.
    parser.add_argument(
        "--tech",
        "--capacity",
        type=int,
        help=argparse.SUPPRESS,
    )

    args = parser.parse_args(argv)

    if (
        args.weeks < 1
        or args.weekly_budget < 0
    ):
        parser.error(
            "Weeks must be positive and "
            "weekly GPU budget cannot be negative"
        )

    root = project_root()

    source = (
        args.input
        or latest_report(root)
    )

    if not source.is_file():
        parser.error(
            f"Input report not found: {source}"
        )

    today = date.today()

    first_week = (
        args.first_week
        or date(2026, 11, 9)
    )

    if first_week.weekday() != 0:
        parser.error(
            "First week must be a Monday"
        )

    output = (
        args.output
        or (
            root
            / "Reports"
            / "H100"
            / "Output"
            / (
                f"H100_Site_Burndown_"
                f"{today:%Y%m%d}.xlsx"
            )
        )
    )

    output.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # summary = read_gpu_requirements(
    #     source
    # )

    summary = sorted(
        {
            "SBN1": 633,
            "PDX3": 296,
            "PDX2": 277,
            "ATL2": 243,
            "ATL4": 140,
            "AUS1": 131,
            "ORD3": 119,
            "PDX1": 87,
            "RNO2": 84,
            "LYF": 66,
            "LAS1": 23,
            "PHX01": 22,
            "DFW1": 11,
        }.items(),
        key=lambda entry: (-entry[1], entry[0]),
    )

    preserved = {}

    if not args.reset_inputs:

        previous = output

        if not previous.is_file():

            earlier = [
                path
                for path in output.parent.glob(
                    "H100_Site_Burndown_*.xlsx"
                )
                if path.is_file()
                and not path.name.startswith("~$")
            ]

            if earlier:

                previous = max(
                    earlier,
                    key=lambda path: path.stat().st_mtime,
                )

        preserved = load_inputs(previous)

    workbook = create_workbook(
        summary,
        first_week,
        args.weeks,
        args.weekly_budget,
        preserved,
    )

    workbook.save(output)

    print(f"Created: {output}")

    print(
        f"Sites: {len(summary)} | "
        f"H100 GPUs needed: "
        f"{sum(value for _, value in summary):,}"
    )

    print(
        f"Existing weekly GPU inputs restored: "
        f"{len(preserved)}"
    )


if __name__ == "__main__":
    main()
