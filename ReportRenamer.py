from pathlib import Path
import shutil
import pandas as pd


# ============================================================
# Source and destination folders
# ============================================================

downloads = Path(r"C:\Users\spencerb\Downloads")

report_folder = Path(
    r"C:\Users\spencerb\OneDrive - Super Micro Computer, Inc"
    r"\Documents\Code\Coreweave Outstanding Reports\Reports"
)

report_folder.mkdir(parents=True, exist_ok=True)


# ============================================================
# Salesforce report names
# ============================================================

report_types = {
    "CoreWeave All Cases with Addresses": "new_address.xlsx",
    "CoreWeave All Outstanding Cases": "case_input.xlsx",
    "CoreWeave Outstanding Cases with RMA": "faulty_part_input.xlsx",
}


# ============================================================
# Find downloaded Excel files
# ============================================================

xlsx_files = list(downloads.glob("*.xlsx"))

if not xlsx_files:
    print("ERROR: No .xlsx files found in Downloads.")
    input("Press Enter to close...")
    raise SystemExit


# ============================================================
# Match files to Salesforce reports
# ============================================================

found_reports = {}

for file in xlsx_files:

    filename = file.stem.lower()

    for identifier, new_name in report_types.items():

        if identifier.lower() in filename:

            found_reports[identifier] = (
                file,
                new_name
            )

            break


# ============================================================
# Check that all three reports were found
# ============================================================

missing = [
    identifier
    for identifier in report_types
    if identifier not in found_reports
]

if missing:

    print("ERROR: Could not find all three Salesforce reports.\n")

    for identifier in missing:
        print(f"Missing: {identifier}")

    print("\nExcel files currently in Downloads:")

    for file in xlsx_files:
        print(f"  - {file.name}")

    input("\nPress Enter to close...")
    raise SystemExit


# ============================================================
# Move and normalize reports
# ============================================================

print("Moving Salesforce reports...\n")


for identifier, (source, new_name) in found_reports.items():

    destination = report_folder / new_name

    # Remove old version if it exists
    if destination.exists():
        destination.unlink()

    # Move downloaded file
    shutil.move(
        str(source),
        str(destination)
    )

    print(f"Moved:")
    print(f"  {source.name}")
    print(f"  -> {new_name}")


    # ========================================================
    # Find the real Salesforce header row
    # ========================================================

    print("  Searching for Salesforce header...")

    raw_df = pd.read_excel(
        destination,
        header=None
    )

    header_row = None

    for row_index in range(len(raw_df)):

        row_values = (
            raw_df.iloc[row_index]
            .astype(str)
            .str.strip()
            .tolist()
        )

        if "Case Number" in row_values:

            header_row = row_index
            break


    # ========================================================
    # Make sure we found the header
    # ========================================================

    if header_row is None:

        print(
            "  ERROR: Could not find 'Case Number' "
            "in this Salesforce export."
        )

        continue


    print(
        f"  Found header on Excel row "
        f"{header_row + 1}"
    )


    # ========================================================
    # Read the file using the correct header
    # ========================================================

    clean_df = pd.read_excel(
        destination,
        header=header_row
    )


    # ========================================================
    # Remove completely empty columns
    # ========================================================

    clean_df = clean_df.dropna(
        axis=1,
        how="all"
    )


    # ========================================================
    # Clean column names
    # ========================================================

    clean_df.columns = (
        clean_df.columns
        .astype(str)
        .str.strip()
    )


    # ========================================================
    # Verify Case Number exists
    # ========================================================

    if "Case Number" not in clean_df.columns:

        print(
            "  ERROR: Header was found, but "
            "'Case Number' is missing after reading."
        )

        continue


    # ========================================================
    # Rewrite the Excel file with the correct headers
    # ========================================================

    clean_df.to_excel(
        destination,
        index=False
    )


    print(
        "  SUCCESS: Salesforce export normalized."
    )

    print(
        f"  Columns found: {', '.join(clean_df.columns)}"
    )

    print()


# ============================================================
# Verify the final files
# ============================================================

print("Verifying reports...\n")

success = True


for new_name in report_types.values():

    destination = report_folder / new_name

    if not destination.exists():

        print(
            f"ERROR: {new_name} was not found!"
        )

        success = False

        continue


    # Verify Case Number is actually the header now
    try:

        test_df = pd.read_excel(
            destination,
            nrows=1
        )

        if "Case Number" in test_df.columns:

            print(
                f"OK: {new_name}"
            )

        else:

            print(
                f"ERROR: {new_name} exists but "
                f"'Case Number' is still missing!"
            )

            success = False

    except Exception as e:

        print(
            f"ERROR reading {new_name}: {e}"
        )

        success = False


print()


if success:

    print(
        "All three Salesforce reports were successfully "
        "moved, renamed, and normalized."
    )

else:

    print(
        "Something went wrong. Please check the Reports folder."
    )


input("\nPress Enter to close...")