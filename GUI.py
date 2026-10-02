from pathlib import Path
from datetime import datetime
import os
import shutil
import threading
import webbrowser

import customtkinter as ctk
import pandas as pd

from AddressValidator import AddressValidator


# ============================================================
# CONFIGURATION
# ============================================================

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

BASE_DIR = Path(__file__).resolve().parent

REPORT_FOLDER = BASE_DIR / "Reports"
ADDRESS_CACHE_FOLDER = BASE_DIR / "Address Cache"
DOWNLOADS_FOLDER = Path.home() / "Downloads"

REPORT_FOLDER.mkdir(exist_ok=True)
ADDRESS_CACHE_FOLDER.mkdir(exist_ok=True)

ADDRESS_FILE = ADDRESS_CACHE_FOLDER / "address_input.xlsx"

NEW_ADDRESS_FILE = REPORT_FOLDER / "new_address.xlsx"
CASE_FILE = REPORT_FOLDER / "case_input.xlsx"
FAULTY_PART_FILE = REPORT_FOLDER / "faulty_part_input.xlsx"
OUTPUT_FILE = REPORT_FOLDER / "outstanding_case_output.xlsx"

REPORT_URLS = [
    "https://supermicrocomputer.lightning.force.com/lightning/r/Report/00OPm00000KI2sfMAD/view?queryScope=userFolders",
    "https://supermicrocomputer.lightning.force.com/lightning/r/Report/00OPm00000KI3APMA1/view?queryScope=userFolders",
    "https://supermicrocomputer.lightning.force.com/lightning/r/Report/00OPm00000KIDGDMA5/view?queryScope=userFolders",
]


REPORT_TYPES = {
    "CoreWeave All Cases with Addresses": "new_address.xlsx",
    "CoreWeave All Outstanding Cases": "case_input.xlsx",
    "CoreWeave Outstanding Cases with RMA": "faulty_part_input.xlsx",
}


# ============================================================
# MAIN APPLICATION
# ============================================================

class CoreWeaveTool(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.address_validator = AddressValidator(
            ADDRESS_FILE
        )

        self.title(
            "CoreWeave Outstanding Reports"
        )

        self.geometry("1000x850")
        self.minsize(900, 700)

        self.build_ui()
        self.refresh()

    # ========================================================
    # UI
    # ========================================================

    def build_ui(self):

        self.grid_columnconfigure(
            0,
            weight=1
        )

        self.grid_rowconfigure(
            1,
            weight=1
        )

        # ----------------------------------------------------
        # HEADER
        # ----------------------------------------------------

        header = ctk.CTkFrame(
            self
        )

        header.grid(
            row=0,
            column=0,
            padx=20,
            pady=(20, 10),
            sticky="ew"
        )

        header.grid_columnconfigure(
            0,
            weight=1
        )

        title = ctk.CTkLabel(
            header,
            text="CoreWeave Outstanding Reports",
            font=ctk.CTkFont(
                size=26,
                weight="bold"
            )
        )

        title.grid(
            row=0,
            column=0,
            padx=20,
            pady=(15, 2),
            sticky="w"
        )

        subtitle = ctk.CTkLabel(
            header,
            text="Salesforce → CoreWeave Report",
            font=ctk.CTkFont(
                size=14
            )
        )

        subtitle.grid(
            row=1,
            column=0,
            padx=20,
            pady=(0, 15),
            sticky="w"
        )

        self.refresh_button = ctk.CTkButton(
            header,
            text="Refresh",
            width=100,
            command=self.refresh
        )

        self.refresh_button.grid(
            row=0,
            column=1,
            rowspan=2,
            padx=20
        )

        # ----------------------------------------------------
        # MAIN CONTENT
        # ----------------------------------------------------

        self.main_frame = ctk.CTkScrollableFrame(
            self
        )

        self.main_frame.grid(
            row=1,
            column=0,
            padx=20,
            pady=(0, 20),
            sticky="nsew"
        )

        self.main_frame.grid_columnconfigure(
            0,
            weight=1
        )

        self.build_workflow()

        # ----------------------------------------------------
        # STATUS BAR
        # ----------------------------------------------------

        self.status_label = ctk.CTkLabel(
            self,
            text="Ready",
            anchor="w"
        )

        self.status_label.grid(
            row=2,
            column=0,
            padx=20,
            pady=(0, 15),
            sticky="ew"
        )

    # ========================================================
    # WORKFLOW
    # ========================================================

    def build_workflow(self):

        # ----------------------------------------------------
        # STEP 1
        # ----------------------------------------------------

        self.step1_frame = self.create_step(
            0,
            "1. Open Salesforce Reports",
            "Open the three required Salesforce reports in your browser."
        )

        self.salesforce_button = ctk.CTkButton(
            self.step1_frame,
            text="Open Salesforce Reports",
            command=self.open_salesforce_reports
        )

        self.salesforce_button.pack(
            anchor="w",
            padx=20,
            pady=(0, 15)
        )

        # ----------------------------------------------------
        # STEP 2
        # ----------------------------------------------------

        self.step2_frame = self.create_step(
            1,
            "2. Organize Salesforce Downloads",
            "Copy the downloaded Salesforce reports into the Reports folder."
        )

        self.organize_button = ctk.CTkButton(
            self.step2_frame,
            text="Organize Reports",
            command=self.organize_reports
        )

        self.organize_button.pack(
            anchor="w",
            padx=20,
            pady=(0, 15)
        )

        # ----------------------------------------------------
        # STEP 3
        # ----------------------------------------------------

        self.step3_frame = self.create_step(
            2,
            "3. Generate CoreWeave Report",
            "Combine the Salesforce reports and generate the working report."
        )

        self.generate_button = ctk.CTkButton(
            self.step3_frame,
            text="Generate Report",
            command=self.generate_report
        )

        self.generate_button.pack(
            anchor="w",
            padx=20,
            pady=(0, 15)
        )

        # ----------------------------------------------------
        # STEP 4
        # ----------------------------------------------------

        self.step4_frame = self.create_step(
            3,
            "4. Validate Addresses",
            "Check for missing or inconsistent City, State, and Region information."
        )

        self.address_textbox = ctk.CTkTextbox(
            self.step4_frame,
            height=180,
            font=ctk.CTkFont(
                family="Consolas",
                size=12
            )
        )

        self.address_textbox.pack(
            fill="x",
            padx=20,
            pady=(0, 10)
        )

        address_buttons = ctk.CTkFrame(
            self.step4_frame,
            fg_color="transparent"
        )

        address_buttons.pack(
            fill="x",
            padx=20,
            pady=(0, 15)
        )

        self.check_address_button = ctk.CTkButton(
            address_buttons,
            text="Validate Addresses",
            command=self.check_address_file
        )

        self.check_address_button.pack(
            side="left",
            padx=(0, 10)
        )

        self.open_address_button = ctk.CTkButton(
            address_buttons,
            text="Open Excel",
            command=self.open_address_file
        )

        self.open_address_button.pack(
            side="left"
        )

        # ----------------------------------------------------
        # STEP 5
        # ----------------------------------------------------

        self.step5_frame = self.create_step(
            4,
            "5. Finalize Report",
            "Verify addresses, regenerate the report, and create the final dated report."
        )

        self.finalize_button = ctk.CTkButton(
            self.step5_frame,
            text="Finalize Report",
            command=self.finalize_report
        )

        self.finalize_button.pack(
            anchor="w",
            padx=20,
            pady=(0, 15)
        )

        # ----------------------------------------------------
        # REPORT FOLDER
        # ----------------------------------------------------

        folder_frame = ctk.CTkFrame(
            self.main_frame
        )

        folder_frame.grid(
            row=5,
            column=0,
            padx=5,
            pady=(15, 5),
            sticky="ew"
        )

        folder_buttons = ctk.CTkFrame(
            folder_frame,
            fg_color="transparent"
        )

        folder_buttons.pack(
            padx=20,
            pady=15
        )

        folder_button = ctk.CTkButton(
            folder_buttons,
            text="Open Reports Folder",
            command=self.open_reports_folder
        )

        folder_button.pack(
            side="left",
            padx=(0, 10)
        )

        clear_reports_button = ctk.CTkButton(
            folder_buttons,
            text="Clear Reports Folder",
            command=self.clear_reports_folder
        )

        clear_reports_button.pack(
            side="left"
        )

    # ========================================================
    # STEP CREATION
    # ========================================================

    def create_step(
        self,
        row,
        title,
        description
    ):

        frame = ctk.CTkFrame(
            self.main_frame
        )

        frame.grid(
            row=row,
            column=0,
            padx=5,
            pady=7,
            sticky="ew"
        )

        frame.grid_columnconfigure(
            0,
            weight=1
        )

        title_label = ctk.CTkLabel(
            frame,
            text=title,
            font=ctk.CTkFont(
                size=18,
                weight="bold"
            ),
            anchor="w"
        )

        title_label.pack(
            fill="x",
            padx=20,
            pady=(15, 3)
        )

        description_label = ctk.CTkLabel(
            frame,
            text=description,
            anchor="w",
            justify="left"
        )

        description_label.pack(
            fill="x",
            padx=20,
            pady=(0, 10)
        )

        status_label = ctk.CTkLabel(
            frame,
            text="Checking...",
            anchor="w"
        )

        status_label.pack(
            fill="x",
            padx=20,
            pady=(0, 10)
        )

        frame.status_label = status_label

        return frame

    # ========================================================
    # STATUS
    # ========================================================

    def set_status(
        self,
        message
    ):

        self.status_label.configure(
            text=message
        )

    def set_worker_status(
        self,
        message
    ):

        self.after(
            0,
            lambda message=message:
                self.set_status(message)
        )

    # ========================================================
    # WORKFLOW STATUS
    # ========================================================

    def update_workflow_status(self):

        # ----------------------------------------------------
        # STEP 1
        # ----------------------------------------------------

        self.step1_frame.status_label.configure(
            text="Ready"
        )

        # ----------------------------------------------------
        # STEP 2
        # ----------------------------------------------------

        required_files = [
            NEW_ADDRESS_FILE,
            CASE_FILE,
            FAULTY_PART_FILE
        ]

        if all(
            path.exists()
            for path in required_files
        ):

            self.step2_frame.status_label.configure(
                text="✓ Salesforce reports are organized"
            )

        else:

            missing = [
                path.name
                for path in required_files
                if not path.exists()
            ]

            self.step2_frame.status_label.configure(
                text=f"Missing: {', '.join(missing)}"
            )

        # ----------------------------------------------------
        # STEP 3
        # ----------------------------------------------------

        if OUTPUT_FILE.exists():

            self.step3_frame.status_label.configure(
                text="✓ CoreWeave report generated"
            )

        else:

            self.step3_frame.status_label.configure(
                text="Not generated"
            )

        # ----------------------------------------------------
        # STEP 4
        # ----------------------------------------------------

        if not ADDRESS_FILE.exists():

            self.step4_frame.status_label.configure(
                text="Address file not found"
            )

        else:

            try:

                results = (
                    self.address_validator.validate()
                )

                missing_count = (
                    results["missing_count"]
                )

                inconsistent_count = (
                    results["inconsistent_count"]
                )

                if (
                    missing_count == 0
                    and inconsistent_count == 0
                ):

                    self.step4_frame.status_label.configure(
                        text="✓ Address validation passed"
                    )

                else:

                    self.step4_frame.status_label.configure(
                        text=(
                            f"⚠ {missing_count} missing, "
                            f"{inconsistent_count} inconsistent"
                        )
                    )

            except Exception as e:

                self.step4_frame.status_label.configure(
                    text=f"⚠ Validation error: {e}"
                )

        # ----------------------------------------------------
        # STEP 5
        # ----------------------------------------------------

        final_files = list(
            REPORT_FOLDER.glob(
                "CoreWeave Outstanding Cases *.xlsx"
            )
        )

        if final_files:

            latest = max(
                final_files,
                key=lambda p: p.stat().st_mtime
            )

            self.step5_frame.status_label.configure(
                text=(
                    f"✓ Latest final report: "
                    f"{latest.name}"
                )
            )

        else:

            self.step5_frame.status_label.configure(
                text="No finalized report yet"
            )

    # ========================================================
    # REFRESH
    # ========================================================

    def refresh(self):

        self.update_workflow_status()

        if ADDRESS_FILE.exists():

            self.check_address_file(
                update_status=False
            )

        else:

            self.address_textbox.delete(
                "1.0",
                "end"
            )

            self.address_textbox.insert(
                "end",
                "Generate the report to create address_input.xlsx."
            )

        self.update_workflow_status()

        self.set_status(
            "Workflow refreshed"
        )

    # ========================================================
    # SALESFORCE
    # ========================================================

    def open_salesforce_reports(self):

        try:

            for url in REPORT_URLS:
                webbrowser.open(url)

            self.set_status(
                "Opened Salesforce reports."
            )

        except Exception as e:

            self.set_status(
                f"Error opening Salesforce: {e}"
            )

    # ========================================================
    # ORGANIZE REPORTS
    # ========================================================

    def organize_reports(self):

        try:

            found = []

            for file in DOWNLOADS_FOLDER.glob(
                "*.xlsx"
            ):

                try:

                    df = pd.read_excel(
                        file,
                        header=None,
                        nrows=10
                    )

                    text = " ".join(
                        df.astype(str)
                        .fillna("")
                        .values
                        .flatten()
                    )

                except Exception:

                    continue

                for (
                    report_name,
                    destination_name
                ) in REPORT_TYPES.items():

                    if report_name.lower() in text.lower():

                        destination = (
                            REPORT_FOLDER
                            / destination_name
                        )

                        shutil.copy2(
                            file,
                            destination
                        )

                        found.append(
                            f"{report_name} → "
                            f"{destination_name}"
                        )

                        break

            if not found:

                self.set_status(
                    "No matching Salesforce reports "
                    "found in Downloads."
                )

                return

            self.set_status(
                f"Organized {len(found)} "
                "Salesforce report(s)."
            )

            self.update_workflow_status()

        except Exception as e:

            self.set_status(
                f"Error organizing reports: {e}"
            )

    # ========================================================
    # EXCEL READING
    # ========================================================

    @staticmethod
    def read_salesforce_excel(path):

        raw_df = pd.read_excel(
            path,
            header=None
        )

        header_row = None

        for row_index in range(
            len(raw_df)
        ):

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
                f"Could not find 'Case Number' "
                f"header in {path.name}"
            )

        df = pd.read_excel(
            path,
            header=header_row,
            dtype={
                "Case Number": str
            }
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

    # ========================================================
    # GENERATE REPORT
    # ========================================================

    def generate_report(self):

        required_files = [
            ADDRESS_FILE,
            NEW_ADDRESS_FILE,
            CASE_FILE,
            FAULTY_PART_FILE
        ]

        missing = [
            path.name
            for path in required_files
            if not path.exists()
        ]

        if missing:

            self.set_status(
                "Missing files: "
                + ", ".join(missing)
            )

            return

        self.generate_button.configure(
            state="disabled",
            text="Generating..."
        )

        self.set_status(
            "Generating CoreWeave report..."
        )

        thread = threading.Thread(
            target=self.generate_report_worker,
            daemon=True
        )

        thread.start()

    def generate_report_worker(self):

        try:

            self.run_coreweave_report()

            self.after(
                0,
                self.report_generation_success
            )

        except Exception as e:

            error_message = str(e)

            print(
                f"REPORT GENERATION ERROR: "
                f"{error_message}"
            )

            self.after(
                0,
                lambda message=error_message:
                    self.report_generation_error(
                        message
                    )
            )

    def report_generation_success(self):

        self.generate_button.configure(
            state="normal",
            text="Generate Report"
        )

        self.update_workflow_status()

        self.check_address_file(
            update_status=False
        )

        self.set_status(
            "CoreWeave report generated successfully."
        )

    def report_generation_error(
        self,
        error
    ):

        self.generate_button.configure(
            state="normal",
            text="Generate Report"
        )

        self.set_status(
            f"Error generating report: {error}"
        )

        self.update_workflow_status()

    # ========================================================
    # COREWEAVE REPORT PROCESSING
    # ========================================================

    def run_coreweave_report(self):

        # ====================================================
        # ADDRESS PROCESSING
        # ====================================================

        self.set_worker_status(
            "Loading address data..."
        )

        old_address_df = pd.read_excel(
            ADDRESS_FILE,
            dtype={
                "Case Number": str
            }
        )

        self.set_worker_status(
            "Loading new address data..."
        )

        new_address_df = (
            self.read_salesforce_excel(
                NEW_ADDRESS_FILE
            )
        )

        self.set_worker_status(
            "Merging address information..."
        )

        new_address_df = pd.merge(
            new_address_df,
            old_address_df,
            how="left",
            on="Case Number"
        )

        # ----------------------------------------------------
        # IGNORE SALESFORCE TOTAL / FOOTER ROWS
        # ----------------------------------------------------

        new_address_df = new_address_df[
            new_address_df["Case Number"].notna()
            & ~new_address_df["Case Number"]
            .astype(str)
            .str.strip()
            .str.lower()
            .eq("total")
        ].copy()

        # ----------------------------------------------------
        # SELECT ADDRESS COLUMNS
        # ----------------------------------------------------

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
                "Asset Location_x":
                    "Asset Location"
            }
        )

        # ----------------------------------------------------
        # BUILD EXISTING LOCATION LOOKUP
        # ----------------------------------------------------

        unique_address_df = old_address_df[
            [
                "Asset Location",
                "City",
                "State",
                "Region"
            ]
        ].drop_duplicates(
            subset=[
                "Asset Location"
            ]
        )

        new_address_df = pd.merge(
            new_address_df,
            unique_address_df,
            how="left",
            on="Asset Location"
        )

        # ----------------------------------------------------
        # FILL CURRENT DATA FROM EXISTING LOCATION DATA
        # ----------------------------------------------------

        new_address_df["City_x"] = (
            new_address_df["City_x"]
            .fillna(
                new_address_df["City_y"]
            )
        )

        new_address_df["State_x"] = (
            new_address_df["State_x"]
            .fillna(
                new_address_df["State_y"]
            )
        )

        new_address_df["Region_x"] = (
            new_address_df["Region_x"]
            .fillna(
                new_address_df["Region_y"]
            )
        )

        # ----------------------------------------------------
        # FINAL ADDRESS DATAFRAME
        # ----------------------------------------------------

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

        self.set_worker_status(
            "Updating address_input.xlsx..."
        )

        new_address_df.to_excel(
            ADDRESS_FILE,
            index=False
        )

        # ====================================================
        # CASE PROCESSING
        # ====================================================

        self.set_worker_status(
            "Loading case data..."
        )

        case_df = self.read_salesforce_excel(
            CASE_FILE
        )

        self.set_worker_status(
            "Adding address information "
            "to case data..."
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

        # ====================================================
        # FAULTY PART PROCESSING
        # ====================================================

        self.set_worker_status(
            "Loading faulty part data..."
        )

        faulty_part_df = (
            self.read_salesforce_excel(
                FAULTY_PART_FILE
            )
        )

        self.set_worker_status(
            "Filtering GPU parts..."
        )

        faulty_part_df = faulty_part_df[
            [
                "Case Number",
                "Faulty Part Number"
            ]
        ]

        faulty_part_df = faulty_part_df[
            faulty_part_df[
                "Faulty Part Number"
            ]
            .astype(str)
            .str.contains(
                "GPU",
                na=False
            )
        ]

        self.set_worker_status(
            "Grouping GPU parts by case..."
        )

        faulty_part_df = (
            faulty_part_df
            .groupby(
                "Case Number"
            )[
                "Faulty Part Number"
            ]
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

        # ====================================================
        # SPLIT OUTSTANDING / SELF-MAINTENANCE
        # ====================================================

        self.set_worker_status(
            "Building outstanding case report..."
        )

        subject_upper = (
            case_df["Subject"]
            .fillna("")
            .astype(str)
            .str.upper()
        )

        outstanding_case_df = case_df[
            ~subject_upper.str.startswith(
                "SELF"
            )
        ]

        self_maint_case_df = case_df[
            subject_upper.str.startswith(
                "SELF"
            )
        ]

        # ====================================================
        # CREATE OUTPUT
        # ====================================================

        self.set_worker_status(
            "Creating output workbook..."
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

        self.set_worker_status(
            "CoreWeave report completed successfully."
        )

    # ========================================================
    # ADDRESS VALIDATION
    # ========================================================

    def check_address_file(
        self,
        update_status=True
    ):

        if not ADDRESS_FILE.exists():

            self.address_textbox.delete(
                "1.0",
                "end"
            )

            self.address_textbox.insert(
                "end",
                "address_input.xlsx does not exist yet."
            )

            if update_status:

                self.set_status(
                    "Address file not found."
                )

            return

        try:

            results = (
                self.address_validator.validate()
            )

            missing_df = results["missing"]
            inconsistent_df = results["inconsistent"]

            self.address_textbox.delete(
                "1.0",
                "end"
            )

            # ------------------------------------------------
            # SUMMARY
            # ------------------------------------------------

            self.address_textbox.insert(
                "end",
                "ADDRESS VALIDATION\n"
                f"{'=' * 70}\n\n"
            )

            self.address_textbox.insert(
                "end",
                f"Unique Asset Locations: "
                f"{results['unique_locations']}\n"
            )

            self.address_textbox.insert(
                "end",
                f"Missing Location Data: "
                f"{results['missing_count']}\n"
            )

            self.address_textbox.insert(
                "end",
                f"Inconsistent Locations: "
                f"{results['inconsistent_count']}\n\n"
            )

            # ------------------------------------------------
            # MISSING LOCATIONS
            # ------------------------------------------------

            if not missing_df.empty:

                self.address_textbox.insert(
                    "end",
                    "MISSING LOCATION INFORMATION\n"
                    f"{'-' * 70}\n\n"
                )

                for index, row in missing_df.iterrows():

                    excel_row = index + 2

                    location = str(
                        row["Asset Location"]
                    )

                    city = (
                        "Missing"
                        if pd.isna(row["City"])
                        or str(row["City"]).strip() == ""
                        else str(row["City"])
                    )

                    state = (
                        "Missing"
                        if pd.isna(row["State"])
                        or str(row["State"]).strip() == ""
                        else str(row["State"])
                    )

                    region = (
                        "Missing"
                        if pd.isna(row["Region"])
                        or str(row["Region"]).strip() == ""
                        else str(row["Region"])
                    )

                    self.address_textbox.insert(
                        "end",
                        f"Excel Row: {excel_row}\n"
                        f"Asset Location: {location}\n"
                        f"City: {city}\n"
                        f"State: {state}\n"
                        f"Region: {region}\n"
                        f"{'-' * 70}\n\n"
                    )

            # ------------------------------------------------
            # INCONSISTENT LOCATIONS
            # ------------------------------------------------

            if not inconsistent_df.empty:

                self.address_textbox.insert(
                    "end",
                    "INCONSISTENT LOCATIONS\n"
                    f"{'-' * 70}\n\n"
                )

                for _, row in inconsistent_df.iterrows():

                    self.address_textbox.insert(
                        "end",
                        "Asset Location:\n"
                        f"{row['Asset Location']}\n\n"
                    )

                    # ----------------------------------------
                    # MAJORITY VALUE
                    # ----------------------------------------

                    majority = row["Majority"]

                    self.address_textbox.insert(
                        "end",
                        "MAJORITY VALUE\n"
                        f"{'-' * 30}\n"
                        f"City: {majority['City']}\n"
                        f"State: {majority['State']}\n"
                        f"Region: {majority['Region']}\n\n"
                    )

                    # ----------------------------------------
                    # CONFLICTING ROWS
                    # ----------------------------------------

                    self.address_textbox.insert(
                        "end",
                        "CONFLICTING ROWS\n"
                        f"{'-' * 30}\n\n"
                    )

                    for detail in row["Rows"]:

                        self.address_textbox.insert(
                            "end",
                            f"Excel Row: "
                            f"{detail['Excel Row']}\n"
                            f"Case Number: "
                            f"{detail['Case Number']}\n"
                            f"City: {detail['City']}\n"
                            f"State: {detail['State']}\n"
                            f"Region: {detail['Region']}\n\n"
                        )

                    self.address_textbox.insert(
                        "end",
                        f"{'-' * 70}\n\n"
                    )

            # ------------------------------------------------
            # EVERYTHING CLEAN
            # ------------------------------------------------

            if (
                missing_df.empty
                and inconsistent_df.empty
            ):

                self.address_textbox.insert(
                    "end",
                    "✓ ADDRESS VALIDATION PASSED\n\n"
                    "All Asset Locations have complete "
                    "and consistent City, State, and Region data."
                )

            # ------------------------------------------------
            # STATUS
            # ------------------------------------------------

            if update_status:

                if (
                    missing_df.empty
                    and inconsistent_df.empty
                ):

                    self.set_status(
                        "Address validation passed."
                    )

                else:

                    self.set_status(
                        f"Address validation found "
                        f"{len(missing_df)} missing "
                        f"and "
                        f"{len(inconsistent_df)} "
                        f"inconsistent location(s)."
                    )

            self.update_workflow_status()

        except Exception as e:

            self.address_textbox.delete(
                "1.0",
                "end"
            )

            self.address_textbox.insert(
                "end",
                f"Error checking address file:\n\n{e}"
            )

            if update_status:

                self.set_status(
                    f"Error checking address file: {e}"
                )

            self.update_workflow_status()

    # ========================================================
    # OPEN ADDRESS FILE
    # ========================================================

    def open_address_file(self):

        if not ADDRESS_FILE.exists():

            self.set_status(
                "address_input.xlsx does not exist yet."
            )

            return

        try:

            os.startfile(
                ADDRESS_FILE
            )

            self.set_status(
                "Opened address_input.xlsx."
            )

        except Exception as e:

            self.set_status(
                f"Error opening Excel: {e}"
            )

    # ========================================================
    # FINALIZE REPORT
    # ========================================================

    def finalize_report(self):

        # ----------------------------------------------------
        # MAKE SURE WORKING REPORT EXISTS
        # ----------------------------------------------------

        if not OUTPUT_FILE.exists():

            self.set_status(
                "Generate the report first."
            )

            return

        # ----------------------------------------------------
        # MAKE SURE ADDRESS FILE EXISTS
        # ----------------------------------------------------

        if not ADDRESS_FILE.exists():

            self.set_status(
                "address_input.xlsx was not found."
            )

            return

        try:

            # ------------------------------------------------
            # VALIDATE CURRENT ADDRESS INFORMATION
            # ------------------------------------------------

            results = (
                self.address_validator.validate()
            )

            missing_df = results["missing"]
            inconsistent_df = results["inconsistent"]

            # ------------------------------------------------
            # STOP IF CURRENT ADDRESSES ARE INVALID
            # ------------------------------------------------

            if (
                not missing_df.empty
                or not inconsistent_df.empty
            ):

                self.check_address_file(
                    update_status=False
                )

                self.set_status(
                    f"Cannot finalize: "
                    f"{len(missing_df)} missing "
                    f"and "
                    f"{len(inconsistent_df)} "
                    f"inconsistent location(s). "
                    "Fix address_input.xlsx, save it, "
                    "then generate the report again."
                )

                return

            # ------------------------------------------------
            # START FINALIZATION
            # ------------------------------------------------

            self.finalize_button.configure(
                state="disabled",
                text="Finalizing..."
            )

            self.generate_button.configure(
                state="disabled"
            )

            self.set_status(
                "All addresses are complete. "
                "Regenerating final report..."
            )

            thread = threading.Thread(
                target=self.finalize_report_worker,
                daemon=True
            )

            thread.start()

        except Exception as e:

            self.set_status(
                f"Error finalizing report: {e}"
            )

    # ========================================================
    # FINALIZE WORKER
    # ========================================================

    def finalize_report_worker(self):

        try:

            # ------------------------------------------------
            # REGENERATE FINAL REPORT
            # ------------------------------------------------

            self.run_coreweave_report()

            # ------------------------------------------------
            # VALIDATE AGAIN AFTER REGENERATION
            # ------------------------------------------------

            self.set_worker_status(
                "Validating regenerated address data..."
            )

            results = (
                self.address_validator.validate()
            )

            missing_df = results["missing"]
            inconsistent_df = results["inconsistent"]

            # ------------------------------------------------
            # STOP IF NEW INVALID ADDRESSES APPEARED
            # ------------------------------------------------

            if (
                not missing_df.empty
                or not inconsistent_df.empty
            ):

                self.after(
                    0,
                    lambda:
                        self.finalize_validation_error(
                            len(missing_df),
                            len(inconsistent_df)
                        )
                )

                return

            # ------------------------------------------------
            # CREATE FINAL FILENAME
            # ------------------------------------------------

            timestamp = datetime.now().strftime(
                "%m%d%Y"
            )

            final_file = (
                REPORT_FOLDER
                / (
                    "CoreWeave Outstanding Cases "
                    f"{timestamp}.xlsx"
                )
            )

            # ------------------------------------------------
            # SEND BACK TO GUI THREAD
            # ------------------------------------------------

            self.after(
                0,
                lambda final_file=final_file:
                    self.finish_finalize(
                        final_file
                    )
            )

        except Exception as e:

            error_message = str(e)

            print(
                f"FINALIZATION ERROR: "
                f"{error_message}"
            )

            self.after(
                0,
                lambda message=error_message:
                    self.finalize_report_error(
                        message
                    )
            )

    # ========================================================
    # FINALIZATION VALIDATION ERROR
    # ========================================================

    def finalize_validation_error(
        self,
        missing_count,
        inconsistent_count
    ):

        self.check_address_file(
            update_status=False
        )

        self.generate_button.configure(
            state="normal"
        )

        self.finalize_button.configure(
            state="normal",
            text="Finalize Report"
        )

        self.update_workflow_status()

        self.set_status(
            f"Finalization stopped: "
            f"{missing_count} missing and "
            f"{inconsistent_count} inconsistent "
            "location(s) were found after regeneration. "
            "Fix address_input.xlsx, save it, "
            "then generate and finalize again."
        )

    # ========================================================
    # FINISH FINALIZATION
    # ========================================================

    def finish_finalize(
        self,
        final_file
    ):

        # ----------------------------------------------------
        # ASK BEFORE OVERWRITING
        # ----------------------------------------------------

        if final_file.exists():

            dialog = ctk.CTkInputDialog(
                text=(
                    f"{final_file.name} already exists.\n\n"
                    "Type YES to replace it:"
                ),
                title="Confirm Overwrite"
            )

            response = dialog.get_input()

            if response is None:

                self.generate_button.configure(
                    state="normal"
                )

                self.finalize_button.configure(
                    state="normal",
                    text="Finalize Report"
                )

                self.set_status(
                    "Finalize cancelled."
                )

                return

            if response.strip().upper() != "YES":

                self.generate_button.configure(
                    state="normal"
                )

                self.finalize_button.configure(
                    state="normal",
                    text="Finalize Report"
                )

                self.set_status(
                    "Finalize cancelled."
                )

                return

        # ----------------------------------------------------
        # COPY FINAL REPORT
        # ----------------------------------------------------

        try:

            shutil.copy2(
                OUTPUT_FILE,
                final_file
            )

            self.generate_button.configure(
                state="normal"
            )

            self.finalize_button.configure(
                state="normal",
                text="Finalize Report"
            )

            self.set_status(
                f"Final report created: "
                f"{final_file.name}"
            )

            self.update_workflow_status()

        except Exception as e:

            self.generate_button.configure(
                state="normal"
            )

            self.finalize_button.configure(
                state="normal",
                text="Finalize Report"
            )

            self.set_status(
                f"Error creating final report: {e}"
            )

    # ========================================================
    # FINALIZATION ERROR
    # ========================================================

    def finalize_report_error(
        self,
        error
    ):

        self.generate_button.configure(
            state="normal"
        )

        self.finalize_button.configure(
            state="normal",
            text="Finalize Report"
        )

        self.update_workflow_status()

        self.set_status(
            f"Error finalizing report: {error}"
        )

    # ========================================================
    # REPORTS FOLDER
    # ========================================================

    def open_reports_folder(self):
        try:
            os.startfile(REPORT_FOLDER)
            self.set_status(
                "Opened Reports folder."
            )
        except Exception as e:
            self.set_status(
                f"Error opening Reports folder: {e}"
            )


    def clear_reports_folder(self):

        # ----------------------------------------------------
        # CONFIRM DELETION
        # ----------------------------------------------------
        dialog = ctk.CTkInputDialog(
            text=(
                "This will delete EVERYTHING inside the Reports folder.\n\n"
                "The Address Cache will NOT be affected.\n\n"
                "Type YES to continue:"
            ),
            title="Clear Reports Folder"
        )

        response = dialog.get_input()

        if response is None:
            self.set_status(
                "Clear Reports cancelled."
            )
            return

        if response.strip().upper() != "YES":
            self.set_status(
                "Clear Reports cancelled."
            )
            return

        # ----------------------------------------------------
        # CLEAR REPORTS FOLDER
        # ----------------------------------------------------
        try:
            deleted_count = 0

            for item in REPORT_FOLDER.iterdir():
                if item.is_dir():
                    shutil.rmtree(item)
                else:
                    item.unlink()

                deleted_count += 1

            self.address_textbox.delete(
                "1.0",
                "end"
            )

            self.address_textbox.insert(
                "end",
                "Reports folder cleared.\n\n"
                "Download fresh Salesforce reports, "
                "then click 'Organize Reports'."
            )

            self.update_workflow_status()

            self.set_status(
                f"Reports folder cleared. "
                f"Removed {deleted_count} item(s)."
            )

        except Exception as e:
            self.set_status(
                f"Error clearing Reports folder: {e}"
            )

# ============================================================
# DIRECT RUN
# ============================================================

if __name__ == "__main__":

    app = CoreWeaveTool()
    app.mainloop()