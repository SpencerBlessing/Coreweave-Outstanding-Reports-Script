from pathlib import Path
import pandas as pd


BASE_DIR = Path(__file__).resolve().parent

OLD_FILE = BASE_DIR / "old.xlsx"
NEW_FILE = BASE_DIR / "new.xlsx"


def clean_value(value):
    if pd.isna(value):
        return ""

    return str(value).strip()


def load_report(path):
    df = pd.read_excel(path)

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    if "Case Number" not in df.columns:
        raise ValueError(
            f"'Case Number' column not found in {path.name}"
        )

    df["Case Number"] = (
        df["Case Number"]
        .astype(str)
        .str.strip()
    )

    return df


print("=" * 80)
print("COREWEAVE REPORT COMPARISON")
print("=" * 80)

print()
print(f"OLD: {OLD_FILE}")
print(f"NEW: {NEW_FILE}")

if not OLD_FILE.exists():
    print(f"\nERROR: {OLD_FILE.name} was not found.")
    input("Press Enter to exit...")
    raise SystemExit

if not NEW_FILE.exists():
    print(f"\nERROR: {NEW_FILE.name} was not found.")
    input("Press Enter to exit...")
    raise SystemExit


old_df = load_report(OLD_FILE)
new_df = load_report(NEW_FILE)


# ------------------------------------------------------------
# Basic information
# ------------------------------------------------------------

print()
print("-" * 80)
print("REPORT SIZE")
print("-" * 80)

print(f"Old rows: {len(old_df)}")
print(f"New rows: {len(new_df)}")


# ------------------------------------------------------------
# Case numbers
# ------------------------------------------------------------

old_cases = set(old_df["Case Number"])
new_cases = set(new_df["Case Number"])

missing_from_new = sorted(old_cases - new_cases)
only_in_new = sorted(new_cases - old_cases)

print()
print("-" * 80)
print("CASE NUMBER DIFFERENCES")
print("-" * 80)

print(f"Missing from new: {len(missing_from_new)}")

for case in missing_from_new:
    print(f"  {case}")

print()
print(f"Only in new: {len(only_in_new)}")

for case in only_in_new:
    print(f"  {case}")


# ------------------------------------------------------------
# Compare shared cases
# ------------------------------------------------------------

shared_cases = old_cases & new_cases

old_indexed = (
    old_df
    .drop_duplicates("Case Number")
    .set_index("Case Number")
)

new_indexed = (
    new_df
    .drop_duplicates("Case Number")
    .set_index("Case Number")
)

shared_columns = (
    set(old_df.columns)
    & set(new_df.columns)
)

shared_columns.discard("Case Number")
differences = []


for case_number in sorted(shared_cases):

    old_row = old_indexed.loc[case_number]
    new_row = new_indexed.loc[case_number]

    for column in sorted(shared_columns):

        old_value = clean_value(old_row[column])
        new_value = clean_value(new_row[column])

        if old_value != new_value:

            differences.append(
                {
                    "Case Number": case_number,
                    "Column": column,
                    "Old": old_value,
                    "New": new_value,
                }
            )


# ------------------------------------------------------------
# Display differences
# ------------------------------------------------------------

print()
print("=" * 80)
print("FIELD DISCREPANCIES")
print("=" * 80)

if not differences:

    print()
    print("NO DISCREPANCIES FOUND.")
    print()
    print("The old and new reports match.")

else:

    print()
    print(f"Found {len(differences)} discrepancies.")

    current_case = None

    for difference in differences:

        case = difference["Case Number"]

        if case != current_case:

            print()
            print("-" * 80)
            print(f"Case Number: {case}")
            print("-" * 80)

            current_case = case

        print()
        print(f"Column: {difference['Column']}")
        print(f"  OLD: {difference['Old']}")
        print(f"  NEW: {difference['New']}")


# ------------------------------------------------------------
# Column differences
# ------------------------------------------------------------

old_only_columns = sorted(
    set(old_df.columns) - set(new_df.columns)
)

new_only_columns = sorted(
    set(new_df.columns) - set(old_df.columns)
)

print()
print("=" * 80)
print("COLUMN DIFFERENCES")
print("=" * 80)

if not old_only_columns and not new_only_columns:

    print()
    print("Column structures match.")

else:

    if old_only_columns:

        print()
        print("Only in OLD:")

        for column in old_only_columns:
            print(f"  {column}")

    if new_only_columns:

        print()
        print("Only in NEW:")

        for column in new_only_columns:
            print(f"  {column}")


# ------------------------------------------------------------
# Finish
# ------------------------------------------------------------

print()
print("=" * 80)
print("COMPARISON COMPLETE")
print("=" * 80)

input("\nPress Enter to exit...")