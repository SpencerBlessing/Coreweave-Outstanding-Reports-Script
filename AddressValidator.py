from pathlib import Path

import pandas as pd


class AddressValidator:

    REQUIRED_COLUMNS = [
        "Case Number",
        "Asset Location",
        "City",
        "State",
        "Region",
    ]

    LOCATION_COLUMNS = [
        "City",
        "State",
        "Region",
    ]

    def __init__(self, address_file):
        self.address_file = Path(address_file)

    # ========================================================
    # LOAD ADDRESS DATA
    # ========================================================

    def load_data(self):

        if not self.address_file.exists():

            raise FileNotFoundError(
                f"Address file not found:\n{self.address_file}"
            )

        df = pd.read_excel(
            self.address_file,
            dtype={
                "Case Number": str,
                "Asset Location": str,
                "City": str,
                "State": str,
                "Region": str,
            },
        )

        # Clean column names
        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        # Make sure the expected columns exist
        missing_columns = [
            column
            for column in self.REQUIRED_COLUMNS
            if column not in df.columns
        ]

        if missing_columns:

            raise ValueError(
                "address_input.xlsx is missing required columns:\n"
                + ", ".join(missing_columns)
            )

        return df

    # ========================================================
    # CLEAN VALUES
    # ========================================================

    @staticmethod
    def clean_value(value):

        if pd.isna(value):
            return ""

        return str(value).strip()

    # ========================================================
    # VALID LOCATION MASK
    # ========================================================

    def valid_location_mask(self, df):
        """
        Identifies real Asset Locations.

        Salesforce footer/total rows are ignored.
        """

        location = (
            df["Asset Location"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        return (
            location.ne("")
            & location.str.lower().ne("total")
        )

    # ========================================================
    # FIND MISSING LOCATIONS
    # ========================================================

    def find_missing_locations(self, df):

        valid_mask = self.valid_location_mask(df)

        missing_mask = (
            df[self.LOCATION_COLUMNS]
            .isna()
            .any(axis=1)
        )

        # Also treat blank strings as missing
        for column in self.LOCATION_COLUMNS:

            missing_mask |= (
                df[column]
                .fillna("")
                .astype(str)
                .str.strip()
                .eq("")
            )

        missing_df = df[
            valid_mask & missing_mask
        ].copy()

        return missing_df

    # ========================================================
    # FIND INCONSISTENT LOCATIONS
    # ========================================================

    def find_inconsistent_locations(self, df):
        """
        Finds Asset Locations that have conflicting
        City / State / Region values.

        Instead of returning every row for the location,
        this method identifies the majority value and
        returns only the rows that disagree with it.

        Example:

        Asset Location:
        1650 East Higgins Road...

        Majority:
        City: Elk Grove Village
        State: IL
        Region: Central (US)

        Outlier:
        Excel Row: 305
        City: Bensenville
        State: IL
        Region: Central (US)
        """

        valid_df = df[
            self.valid_location_mask(df)
        ].copy()

        if valid_df.empty:

            return pd.DataFrame(
                columns=[
                    "Asset Location",
                    "Majority",
                    "Rows",
                ]
            )

        # Normalize values
        for column in self.LOCATION_COLUMNS:

            valid_df[column] = (
                valid_df[column]
                .fillna("")
                .astype(str)
                .str.strip()
            )

        inconsistent_locations = []

        grouped = valid_df.groupby(
            "Asset Location",
            dropna=False
        )

        for location, group in grouped:

            # ------------------------------------------------
            # FIND MAJORITY VALUE FOR EACH FIELD
            # ------------------------------------------------

            majority_values = {}

            for column in self.LOCATION_COLUMNS:

                counts = (
                    group[column]
                    .value_counts()
                )

                if counts.empty:

                    majority_values[column] = ""

                else:

                    majority_values[column] = (
                        counts.index[0]
                    )

            # ------------------------------------------------
            # FIND ROWS THAT DISAGREE WITH THE MAJORITY
            # ------------------------------------------------

            outlier_rows = []

            for index, row in group.iterrows():

                is_outlier = False

                for column in self.LOCATION_COLUMNS:

                    if (
                        row[column]
                        != majority_values[column]
                    ):

                        is_outlier = True
                        break

                if not is_outlier:
                    continue

                # Excel row number
                excel_row = index + 2

                outlier_rows.append(
                    {
                        "Excel Row": excel_row,

                        "Case Number": (
                            self.clean_value(
                                row["Case Number"]
                            )
                        ),

                        "City": (
                            row["City"]
                            if row["City"]
                            else "Missing"
                        ),

                        "State": (
                            row["State"]
                            if row["State"]
                            else "Missing"
                        ),

                        "Region": (
                            row["Region"]
                            if row["Region"]
                            else "Missing"
                        ),
                    }
                )

            # ------------------------------------------------
            # ONLY REPORT LOCATIONS WITH OUTLIERS
            # ------------------------------------------------

            if outlier_rows:

                inconsistent_locations.append(
                    {
                        "Asset Location": location,

                        "Majority": {
                            "City": (
                                majority_values["City"]
                                if majority_values["City"]
                                else "Missing"
                            ),

                            "State": (
                                majority_values["State"]
                                if majority_values["State"]
                                else "Missing"
                            ),

                            "Region": (
                                majority_values["Region"]
                                if majority_values["Region"]
                                else "Missing"
                            ),
                        },

                        "Rows": outlier_rows,
                    }
                )

        return pd.DataFrame(
            inconsistent_locations
        )

    # ========================================================
    # FIND NEW LOCATIONS
    # ========================================================

    def find_new_locations(self, df):
        """
        Identifies Asset Locations that exist in the
        current address file.

        This is primarily useful for reporting/diagnostics.
        """

        valid_df = df[
            self.valid_location_mask(df)
        ].copy()

        if valid_df.empty:

            return pd.DataFrame(
                columns=[
                    "Asset Location"
                ]
            )

        return (
            valid_df[
                ["Asset Location"]
            ]
            .drop_duplicates()
            .reset_index(drop=True)
        )

    # ========================================================
    # VALIDATE
    # ========================================================

    def validate(self):

        df = self.load_data()

        valid_mask = self.valid_location_mask(df)

        total_rows = int(
            valid_mask.sum()
        )

        missing_df = (
            self.find_missing_locations(df)
        )

        inconsistent_df = (
            self.find_inconsistent_locations(df)
        )

        unique_locations = int(
            df.loc[
                valid_mask,
                "Asset Location"
            ]
            .drop_duplicates()
            .shape[0]
        )

        complete_locations = (
            unique_locations
            - self._count_unique_missing_locations(
                missing_df
            )
        )

        return {
            "dataframe": df,

            "total_rows": total_rows,

            "unique_locations": unique_locations,

            "complete_locations": max(
                0,
                complete_locations
            ),

            "missing": missing_df,

            "inconsistent": inconsistent_df,

            "missing_count": len(
                missing_df
            ),

            "inconsistent_count": len(
                inconsistent_df
            ),
        }

    # ========================================================
    # COUNT UNIQUE MISSING LOCATIONS
    # ========================================================

    @staticmethod
    def _count_unique_missing_locations(df):

        if df.empty:
            return 0

        return int(
            df["Asset Location"]
            .dropna()
            .astype(str)
            .str.strip()
            .nunique()
        )