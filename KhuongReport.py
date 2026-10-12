
from pathlib import Path
import pandas as pd

# File locations
project_dir = Path(__file__).resolve().parent
reports_dir = project_dir / "Reports"

input_file = reports_dir / "CoreWeave Outstanding Cases 10122026.xlsx"
output_file = reports_dir / "H100 Pending Customer Verification 10042026-10122026.xlsx"

# Load the main outstanding-cases worksheet
df = pd.read_excel(
    input_file,
    sheet_name="CoreWeave Outstanding Cases"
)

# Convert Case Creation Date to dates
df["Case Creation Date"] = pd.to_datetime(
    df["Case Creation Date"],
    format="mixed",
    errors="coerce"
)

start_date = pd.Timestamp("2026-10-04")
end_date = pd.Timestamp("2026-10-12 23:59:59")

# Apply all three filters
filtered = df[
    df["Faulty Part Number"].astype("string").str.contains(
        "H100", case=False, na=False
    )
    & df["Status"].eq("Pending Customer Verification")
    & df["Case Creation Date"].between(start_date, end_date)
].copy()

# Export matching cases
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    filtered.to_excel(writer, sheet_name="Filtered Cases", index=False)

    # Summary of matching faulty part numbers
    summary = (
        filtered["Faulty Part Number"]
        .value_counts()
        .rename_axis("Faulty Part Number")
        .reset_index(name="Case Count")
    )
    summary.to_excel(writer, sheet_name="Part Number Summary", index=False)

print(f"Matching cases: {len(filtered)}")
print(f"Output saved to: {output_file}")
