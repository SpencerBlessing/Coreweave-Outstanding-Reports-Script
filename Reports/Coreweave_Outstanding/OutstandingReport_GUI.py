import sys
import os
import shutil
import threading
import webbrowser
import customtkinter as ctk
import pandas as pd

from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook

# Project root: Coreweave Outstanding Reports/
BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

from Utilities.CaseNumberStandardizer import standardize_case_number
from Utilities.AddressValidator import AddressValidator
from Utilities.AddressAutoCorrect import auto_correct_addresses
from Reports.Coreweave_Outstanding.OutstandingReport import generate_report

# ============================================================
# CONFIGURATION
# ============================================================

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

# Report-specific folder
OUTSTANDING_FOLDER = BASE_DIR / "Reports" / "Coreweave_Outstanding"

# Input reports downloaded from Salesforce
REPORT_FOLDER = OUTSTANDING_FOLDER / "Input"

# Shared address cache stays at project root
ADDRESS_CACHE_FOLDER = BASE_DIR / "Address Cache"
ADDRESS_FILE = ADDRESS_CACHE_FOLDER / "address_input.xlsx"

# Salesforce input reports
NEW_ADDRESS_FILE = REPORT_FOLDER / "new_address.xlsx"
CASE_FILE = REPORT_FOLDER / "case_input.xlsx"
FAULTY_PART_FILE = REPORT_FOLDER / "faulty_part_input.xlsx"
RMA_FILE = REPORT_FOLDER / "rma_input.xlsx"

# Generated reports
OUTPUT_FOLDER = OUTSTANDING_FOLDER / "Output"
OUTPUT_FILE = OUTPUT_FOLDER / "outstanding_case_output.xlsx"

# Ensure folders exist
REPORT_FOLDER.mkdir(parents=True, exist_ok=True)
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

ADDRESS_CACHE_FOLDER.mkdir(parents=True, exist_ok=True)

DOWNLOADS_FOLDER = Path.home() / "Downloads"

REPORT_URLS = [
    "https://supermicrocomputer.lightning.force.com/lightning/r/Report/00OPm00000KI2sfMAD/view?queryScope=userFolders",
    "https://supermicrocomputer.lightning.force.com/lightning/r/Report/00OPm00000KI3APMA1/view?queryScope=userFolders",
    "https://supermicrocomputer.lightning.force.com/lightning/r/Report/00OPm00000KIDGDMA5/view?queryScope=userFolders",
    "https://supermicrocomputer.lightning.force.com/lightning/r/Report/00OPm00000MgLg5MAF/view?queryScope=userFolders",
]


REPORT_TYPES = {
    "CoreWeave All Cases with Addresses": "new_address.xlsx",
    "CoreWeave All Outstanding Cases": "case_input.xlsx",
    "CoreWeave Outstanding Cases with RMA": "faulty_part_input.xlsx",
    "All CoreWeave RMA": "rma_input.xlsx",
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
            "Open the Salesforce reports in your browser."
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

        self.since_label = ctk.CTkLabel(
        self.step3_frame,
        text="Optional start date (YYYY-MM-DD):"
       )

        self.since_label.pack(
            anchor="w",
            padx=20,
            pady=(0, 5)
        )

        self.since_entry = ctk.CTkEntry(
            self.step3_frame,
            placeholder_text="e.g. 2026-10-04",
            width=200
        )

        self.since_entry.pack(
            anchor="w",
            padx=20,
            pady=(0, 10)
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

        self.auto_correct_button = ctk.CTkButton(
            address_buttons,
            text="Auto-Correct Addresses",
            command=self.auto_correct_address_file
        )
        self.auto_correct_button.pack(side="left", padx=(0, 10))

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
            text="Not started",
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
            FAULTY_PART_FILE,
            RMA_FILE
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
            OUTPUT_FOLDER.glob(
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
            REPORT_FOLDER.mkdir(parents=True, exist_ok=True)
            downloads = list(DOWNLOADS_FOLDER.glob("*.xlsx"))
            copied = []
            missing = []

            for report_name, destination_name in REPORT_TYPES.items():
                matches = [
                    path for path in downloads
                    if report_name.casefold() in path.stem.casefold()
                ]
                if not matches:
                    missing.append(report_name)
                    continue

                # Always use the most recently modified matching download.
                source = max(matches, key=lambda path: path.stat().st_mtime)
                destination = REPORT_FOLDER / destination_name

                # Validate the source before replacing a previous good export.
                cleaned = self.read_salesforce_excel(source)
                if "Case Number" not in cleaned.columns:
                    raise ValueError(f"Missing Case Number in {source.name}")
                cleaned.to_excel(destination, index=False)
                copied.append(f"{source.name} -> {destination_name}")

            for item in copied:
                print(item)
            if missing:
                print("Missing downloads: " + ", ".join(missing))
                self.set_status(
                    f"Organized {len(copied)}/{len(REPORT_TYPES)} reports. "
                    f"Missing: {', '.join(missing)}"
                )
            else:
                self.set_status(f"Organized all {len(REPORT_TYPES)} Salesforce reports.")
            self.update_workflow_status()
        except Exception as exc:
            print(f"Error organizing reports: {exc}")
            self.set_status(f"Error organizing reports: {exc}")

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
            NEW_ADDRESS_FILE,
            CASE_FILE,
            FAULTY_PART_FILE,
            RMA_FILE
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

        since_date = self.since_entry.get().strip() or None

        if since_date:
            try:
                datetime.strptime(since_date, "%Y-%m-%d")
            except ValueError:
                self.set_status("Invalid date. Use YYYY-MM-DD.")
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
            args=(since_date,),
            daemon=True
        )
        thread.start()

    def generate_report_worker(self, since_date=None):
        try:
            self.run_coreweave_report(since=since_date)

            self.after(0, self.report_generation_success)

        except Exception as e:
            error_message = str(e)
            print(f"REPORT GENERATION ERROR: {error_message}")

            self.after(
                0,
                lambda message=error_message:
                    self.report_generation_error(message)
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

    def run_coreweave_report(self, since=None):
        return generate_report(since=since, status_callback=self.set_worker_status)

    # ========================================================

    def check_address_file(self, update_status=True):
        if not ADDRESS_FILE.exists():
            self.address_textbox.delete("1.0", "end")
            self.address_textbox.insert("end", "address_input.xlsx does not exist yet.")
            return
        try:
            results = self.address_validator.validate()
            missing = results["missing"]
            inconsistent = results["inconsistent"]
            lines = ["ADDRESS VALIDATION", "=" * 70,
                     f"Unique Asset Locations: {results['unique_locations']}",
                     f"Missing Location Data: {len(missing)}",
                     f"Inconsistent Location Fields: {len(inconsistent)}", ""]
            if not missing.empty:
                lines.append("MISSING LOCATION INFORMATION")
                for idx, row in missing.iterrows():
                    lines.append(f"Row {idx + 2}: {row.get('Case Number', '')} | {row.get('Asset Location', '')}")
            if not inconsistent.empty:
                lines.extend(["", "INCONSISTENT LOCATION INFORMATION"])
                for _, row in inconsistent.iterrows():
                    lines.append(f"{row['Asset Location']} | {row['Field']}: {row['Observed Values']} | Cases: {row['Case Numbers']}")
            if missing.empty and inconsistent.empty:
                lines.append("ADDRESS VALIDATION PASSED")
            self.address_textbox.delete("1.0", "end")
            self.address_textbox.insert("end", "\n".join(lines))
            if update_status:
                self.set_status(f"Address validation: {len(missing)} missing, {len(inconsistent)} inconsistent.")
            self.update_workflow_status()
        except Exception as exc:
            self.address_textbox.delete("1.0", "end")
            self.address_textbox.insert("end", f"Error checking address file:\n\n{exc}")
            if update_status:
                self.set_status(f"Error checking address file: {exc}")

    # ========================================================

    def auto_correct_address_file(self):
        if not ADDRESS_FILE.exists():
            self.set_status("address_input.xlsx does not exist yet.")
            return
        try:
            result = auto_correct_addresses(ADDRESS_FILE)
            self.check_address_file(update_status=False)
            counts = result["fields_filled"]
            self.set_status(
                f"Auto-corrected {result['rows_updated']} rows "
                f"(City: {counts['City']}, State: {counts['State']}, "
                f"Region: {counts['Region']}). "
                "Re-generate the report to update its output."
            )
        except Exception as exc:
            self.set_status(f"Auto-correction failed: {exc}")

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

            since_date = self.since_entry.get().strip() or None

            if since_date:
                try:
                    datetime.strptime(since_date, "%Y-%m-%d")
                except ValueError:
                    self.set_status(
                        "Invalid date. Use YYYY-MM-DD."
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
                args=(since_date,),
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

    def finalize_report_worker(self, since_date=None):

        try:

            # ------------------------------------------------
            # REGENERATE FINAL REPORT
            # ------------------------------------------------

            self.run_coreweave_report(since=since_date)

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
                OUTPUT_FOLDER
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

            # Copy the working report to the final filename
            shutil.copy2(
                OUTPUT_FILE,
                final_file
            )

            # Remove Unnamed columns while preserving worksheets
            wb = load_workbook(final_file)

            try:
                for ws in wb.worksheets:
                    for col in range(ws.max_column, 0, -1):
                        header = ws.cell(row=1, column=col).value

                        if (
                            header is not None
                            and str(header).startswith("Unnamed:")
                        ):
                            ws.delete_cols(col)

                wb.save(final_file)

            finally:
                wb.close()


            self.generate_button.configure(
                state="normal"
            )

            self.finalize_button.configure(
                state="normal",
                text="Finalize Report"
            )

            self.set_status(
                f"Final report created: {final_file.name}"
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
            os.startfile(OUTPUT_FOLDER)
            self.set_status("Opened Outstanding Reports Output folder.")
        except Exception as e:
            self.set_status(f"Error opening Output folder: {e}")

    def clear_reports_folder(self):
        """Clear all contents of the Outstanding Input and Output folders."""

        dialog = ctk.CTkInputDialog(
            text=(
                "This will permanently delete all files and subfolders inside:\n\n"
                "Reports\\Coreweave_Outstanding\\Input\n"
                "Reports\\Coreweave_Outstanding\\Output\n\n"
                "The Address Cache will NOT be affected.\n\n"
                "Type YES to continue:"
            ),
            title="Clear Outstanding Reports"
        )

        response = dialog.get_input()

        if response is None or response.strip().upper() != "YES":
            self.set_status("Clear Reports cancelled.")
            return

        try:
            folders_to_clear = [REPORT_FOLDER, OUTPUT_FOLDER]
            deleted_count = 0
            errors = []

            for folder in folders_to_clear:
                folder.mkdir(parents=True, exist_ok=True)

                for item in list(folder.iterdir()):
                    try:
                        if item.is_dir() and not item.is_symlink():
                            shutil.rmtree(item)
                        else:
                            item.unlink()

                        deleted_count += 1

                    except Exception as e:
                        errors.append(f"{item}: {e}")

            self.address_textbox.delete("1.0", "end")

            if errors:
                self.address_textbox.insert(
                    "end",
                    "Some files could not be deleted:\n\n"
                    + "\n".join(errors)
                )
                self.set_status(
                    f"Cleared {deleted_count} item(s), "
                    f"but {len(errors)} item(s) failed. See details."
                )
            else:
                self.address_textbox.insert(
                    "end",
                    "Input and Output folders cleared.\n\n"
                    "The Address Cache was preserved."
                )
                self.set_status(
                    f"Successfully deleted {deleted_count} file(s)/folder(s) "
                    f"from Input and Output."
                )

            self.update_workflow_status()

        except Exception as e:
            self.set_status(f"Error clearing report folders: {e}")

# ============================================================
# DIRECT RUN
# ============================================================

if __name__ == "__main__":

    app = CoreWeaveTool()
    app.mainloop()