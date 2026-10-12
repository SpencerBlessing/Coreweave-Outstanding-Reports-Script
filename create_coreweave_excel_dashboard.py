"""
Combine existing chart PNGs into one organized Excel worksheet.
This script does NOT compare reports or regenerate charts.

Run from the project folder:
    py create_coreweave_excel_dashboard.py

Expected chart folder:
    Reports\Charts

Requirements:
    py -m pip install openpyxl pillow
"""
from pathlib import Path
from datetime import datetime

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import pandas as pd

CHARTS = [
    ("cases_by_age.png", "Outstanding Cases by Age"),
    ("cases_by_status.png", "Outstanding Cases by Status"),
    ("cases_by_age_and_status.png", "Outstanding Cases by Age and Status"),
    ("cases_by_region_pie.png", "Outstanding Cases by Region"),
    ("cases_by_part_number.png", "Parts Summary"),
    ("open_cases_over_time.png", "Outstanding Cases by Creation Date"),
    ("status_by_a.png", "Open less than 9 days."),
    ("status_by_b.png", "Open 10 - 15 days."),
    ("status_by_c.png", "Open 16 - 30 days."),
    ("status_by_d.png", "Open more than 30 days."),
]

def main():
    project_dir = Path(__file__).resolve().parent

    possible_chart_dirs = [
        project_dir / "Reports" / "Charts",
        project_dir / "Charts",
    ]
    charts_dir = next((p for p in possible_chart_dirs if p.is_dir()), None)

    if charts_dir is None:
        raise FileNotFoundError(
            "Couldn't find a Charts folder. Expected either:\n"
            f"  {possible_chart_dirs[0]}\n"
            f"  {possible_chart_dirs[1]}"
        )

    available = [
        (filename, title, charts_dir / filename)
        for filename, title in CHARTS
        if (charts_dir / filename).is_file()
    ]
    if not available:
        raise FileNotFoundError(
            f"No recognized chart PNGs were found in {charts_dir}."
        )

    wb = Workbook()
    ws = wb.active
    ws.title = "Charts Dashboard"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"

    # Header
    ws.merge_cells("A1:Z1")
    ws["A1"] = "COREWEAVE OUTSTANDING CASES — CHARTS"
    ws["A1"].font = Font(name="Aptos Display", size=20, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="183247")
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:Z2")
    ws["A2"] = f"Compiled {datetime.now().strftime('%B %d, %Y')}  |  {len(available)} charts"
    ws["A2"].font = Font(size=10, color="183247")
    ws["A2"].fill = PatternFill("solid", fgColor="EAF1F5")
    ws["A2"].alignment = Alignment(vertical="center")
    ws.row_dimensions[2].height = 23

    ws.row_dimensions[3].hidden = True

    for col in range(1, 27):
        ws.column_dimensions[get_column_letter(col)].width = 11

    # Two charts per row band, keeping each image's aspect ratio.
    first_image_row = 5
    row_band = 29
    target_width_px = 760

    for i, (_, title, image_path) in enumerate(available):
        img = XLImage(str(image_path))
        scale = target_width_px / img.width
        img.width = target_width_px
        img.height = int(img.height * scale)

        band_row = first_image_row + (i // 2) * row_band
        is_left = (i % 2 == 0)
        anchor_col = "A" if is_left else "N"
        caption_col = 1 if is_left else 14

        caption = ws.cell(row=band_row - 1, column=caption_col, value=title)
        caption.font = Font(size=12, bold=True, color="183247")
        caption.alignment = Alignment(vertical="center")
        ws.row_dimensions[band_row - 1].height = 23

        img.anchor = f"{anchor_col}{band_row}"
        ws.add_image(img)

       
    # Compact report comparison summary on the dashboard
    import re

    report_pattern = re.compile(
        r"CoreWeave Outstanding Cases (\d{8})\.xlsx",
        re.IGNORECASE
    )

    reports_dir = project_dir / "Reports"
    reports = []

    for file in reports_dir.glob("CoreWeave Outstanding Cases *.xlsx"):
        match = report_pattern.fullmatch(file.name)
        if match:
            try:
                date = datetime.strptime(match.group(1), "%m%d%Y")
                reports.append((date, file))
            except ValueError:
                pass

    reports.sort(key=lambda x: x[0])

    if len(reports) >= 2:
        previous_date, previous_path = reports[-2]
        current_date, current_path = reports[-1]

        sheet_name = "CoreWeave Outstanding Cases"
        key = "Case Number"
        status_col = "Status"

        old_df = pd.read_excel(previous_path, sheet_name=sheet_name)
        new_df = pd.read_excel(current_path, sheet_name=sheet_name)

        old_df = old_df.dropna(subset=[key]).set_index(key)
        new_df = new_df.dropna(subset=[key]).set_index(key)

        old_cases = set(old_df.index)
        new_cases = set(new_df.index)
        shared = old_cases & new_cases

        opened = len(new_cases - old_cases)
        closed = len(old_cases - new_cases)

        status_changed = int(
            old_df.loc[list(shared), status_col]
            .ne(new_df.loc[list(shared), status_col])
            .sum()
        )

        ws.merge_cells("A4:Z4")
        ws["A4"] = (
            f"Report Changes | "
            f"{previous_date:%m/%d/%Y} → {current_date:%m/%d/%Y}"
            f"    •    Opened: {opened}"
            f"    •    Closed: {closed}"
            f"    •    Status Changes: {status_changed}"
        )
        ws["A4"].font = Font(
            name="Aptos", size=11, bold=True, color="183247"
        )
        ws["A4"].fill = PatternFill("solid", fgColor="EAF1F5")
        ws["A4"].alignment = Alignment(
            vertical="center", horizontal="left", indent=1
        )
        ws.row_dimensions[4].height = 28

    else:
        print("Comparison skipped: need at least two dated source reports.")


    output_dir = project_dir / "Reports"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "CoreWeave Charts Dashboard.xlsx"

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = "landscape"
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5

    wb.save(output_path)
    print(f"Excel charts workbook created: {output_path}")
    print(f"Charts embedded: {len(available)}")
    if len(available) < len(CHARTS):
        missing = [name for name, _, _ in CHARTS if not (charts_dir / name).is_file()]
        print("Charts not found (skipped): " + ", ".join(missing))

if __name__ == "__main__":
    main()
