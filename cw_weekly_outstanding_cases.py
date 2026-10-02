from pathlib import Path

import pandas as pd

# --------------------------------------------------------
# Paths
# --------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent

REPORT_FOLDER = BASE_DIR / "Reports"
ADDRESS_CACHE_FOLDER = BASE_DIR / "Address Cache"

REPORT_FOLDER.mkdir(exist_ok=True)
ADDRESS_CACHE_FOLDER.mkdir(exist_ok=True)

ADDRESS_FILE = ADDRESS_CACHE_FOLDER / "address_input.xlsx"

NEW_ADDRESS_FILE = REPORT_FOLDER / "new_address.xlsx"
CASE_FILE = REPORT_FOLDER / "case_input.xlsx"
FAULTY_PART_FILE = REPORT_FOLDER / "faulty_part_input.xlsx"
OUTPUT_FILE = REPORT_FOLDER / "outstanding_case_output.xlsx"


# --------------------------------------------------------
# Salesforce Excel reader
# --------------------------------------------------------

def read_salesforce_excel(path):

    raw_df = pd.read_excel(
        path,
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

    if header_row is None:

        raise ValueError(
            f"Could not find 'Case Number' header in {path}"
        )

    df = pd.read_excel(
        path,
        header=header_row,
        dtype={"Case Number": str}
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    df = df.dropna(
        axis=1,
        how="all"
    )

    return df


# --------------------------------------------------------
# Start
# --------------------------------------------------------

print(
    "Starting CoreWeave report...",
    flush=True
)


# --------------------------------------------------------
# Address processing
# --------------------------------------------------------

print(
    "Loading address data...",
    flush=True
)

old_address_df = pd.read_excel(
    ADDRESS_FILE,
    dtype={"Case Number": str}
)

print(
    "Loading new address data...",
    flush=True
)

new_address_df = read_salesforce_excel(
    NEW_ADDRESS_FILE
)

print(
    "Merging address information...",
    flush=True
)

new_address_df = pd.merge(
    new_address_df,
    old_address_df,
    how="left",
    on="Case Number"
)

new_address_df = new_address_df[
    [
        "Case Number",
        "Asset Location_x",
        "City",
        "State",
        "Region"
    ]
]

new_address_df = new_address_df.rename(
    columns={
        "Asset Location_x": "Asset Location"
    }
)

unique_address_df = old_address_df[
    [
        "Asset Location",
        "City",
        "State",
        "Region"
    ]
].drop_duplicates(
    subset=["Asset Location"]
)

new_address_df = pd.merge(
    new_address_df,
    unique_address_df,
    how="left",
    on="Asset Location"
)

new_address_df["City_x"] = (
    new_address_df["City_x"]
    .fillna(new_address_df["City_y"])
)

new_address_df["State_x"] = (
    new_address_df["State_x"]
    .fillna(new_address_df["State_y"])
)

new_address_df["Region_x"] = (
    new_address_df["Region_x"]
    .fillna(new_address_df["Region_y"])
)

new_address_df = new_address_df[
    [
        "Case Number",
        "Asset Location",
        "City_x",
        "State_x",
        "Region_x"
    ]
]

new_address_df = new_address_df.rename(
    columns={
        "City_x": "City",
        "State_x": "State",
        "Region_x": "Region"
    }
)

print(
    "Updating address_input.xlsx...",
    flush=True
)

new_address_df.to_excel(
    ADDRESS_FILE,
    index=False
)


# --------------------------------------------------------
# Case processing
# --------------------------------------------------------

print(
    "Loading case data...",
    flush=True
)

case_df = read_salesforce_excel(
    CASE_FILE
)

print(
    "Adding address information to case data...",
    flush=True
)

address_df = new_address_df[
    [
        "Case Number",
        "City",
        "State",
        "Region"
    ]
]

case_df = pd.merge(
    case_df,
    address_df,
    how="left",
    on="Case Number"
)


# --------------------------------------------------------
# Faulty part processing
# --------------------------------------------------------

print(
    "Loading faulty part data...",
    flush=True
)

faulty_part_df = read_salesforce_excel(
    FAULTY_PART_FILE
)

print(
    "Filtering GPU parts...",
    flush=True
)

faulty_part_df = faulty_part_df[
    [
        "Case Number",
        "Faulty Part Number"
    ]
]

faulty_part_df = faulty_part_df[
    faulty_part_df["Faulty Part Number"]
    .str.contains(
        "GPU",
        na=False
    )
]

print(
    "Grouping GPU parts by case...",
    flush=True
)

faulty_part_df = (
    faulty_part_df
    .groupby("Case Number")["Faulty Part Number"]
    .agg(
        lambda x: ", ".join(
            x.dropna()
            .astype(str)
            .unique()
        )
    )
    .reset_index()
)

case_df = pd.merge(
    case_df,
    faulty_part_df,
    how="left",
    on="Case Number"
)


# --------------------------------------------------------
# Split outstanding / self-maintenance cases
# --------------------------------------------------------

print(
    "Building outstanding case report...",
    flush=True
)

outstanding_case_df = case_df[
    ~case_df["Subject"]
    .str.upper()
    .str.startswith(
        "SELF",
        na=False
    )
]

self_maint_case_df = case_df[
    case_df["Subject"]
    .str.upper()
    .str.startswith(
        "SELF",
        na=False
    )
]


# --------------------------------------------------------
# Create output
# --------------------------------------------------------

print(
    "Creating output workbook...",
    flush=True
)

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    outstanding_case_df.to_excel(
        writer,
        sheet_name="CoreWeave Outstanding Cases",
        index=False
    )

    self_maint_case_df.to_excel(
        writer,
        sheet_name="Self-Maintenance",
        index=False
    )


# --------------------------------------------------------
# Complete
# --------------------------------------------------------

print(
    "CoreWeave report completed successfully.",
    flush=True
)