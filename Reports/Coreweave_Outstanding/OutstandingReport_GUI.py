import sys
import os
import shutil
import threading
import tempfile
import webbrowser
import customtkinter as ctk
import pandas as pd
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
# Project root: Coreweave Outstanding Reports/
BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))
from Utilities.AddressValidator import AddressValidator
from Reports.Coreweave_Outstanding.OutstandingReport import (
    generate_report,
    commit_finalized_history,
    DEBUG_FOLDER,
)
# ============================================================
# CONFIGURATION
# ============================================================
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")
# Report-specific folder
OUTSTANDING_FOLDER = BASE_DIR / "Reports" / "Coreweave_Outstanding"
# Input reports downloaded from Salesforce
REPORT_FOLDER = OUTSTANDING_FOLDER / "Input"
# Single authoritative site/address map. No historical address cache.
ADDRESS_CACHE_FOLDER = BASE_DIR / "Address Cache"
SITE_MAP_FILE = ADDRESS_CACHE_FOLDER / "CoreWeave_Map.json"
# Salesforce input reports
CASE_FILE = REPORT_FOLDER / "case_input.xlsx"
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
    "https://supermicrocomputer.lightning.force.com/lightning/r/Report/00OPm00000KI3APMA1/view?queryScope=userFolders",
    "https://supermicrocomputer.lightning.force.com/lightning/r/Report/00OPm00000JGzZFMA1/view?queryScope=userFolders",
]
REPORT_TYPES = {
    "CoreWeave All Outstanding Cases": "case_input.xlsx",
    "All CoreWeave RMA": "rma_input.xlsx",
}

# ============================================================
# MAIN APPLICATION
# ============================================================
class CoreWeaveTool(ctk.CTk):
    def __init__(self):
        super().__init__()
        # Re-read the live Salesforce export and map for every validation.
        # Results are held only for the current GUI session, never saved as a cache.
        self.validation_results = None
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
            "4. Validate Addresses & Sites",
            "Validate current outstanding case addresses against CoreWeave_Map.json. Never auto-update the map."
        )
        self.address_textbox = ctk.CTkTextbox(
            self.step4_frame,
            height=350,
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
            text="Validate Addresses & Sites",
            command=self.check_address_file
        )
        self.check_address_button.pack(
            side="left",
            padx=(0, 10)
        )
        self.export_review_button = ctk.CTkButton(
            address_buttons,
            text="Export Review",
            command=self.export_address_review
        )
        self.export_review_button.pack(side="left", padx=(0, 10))
        self.open_address_button = ctk.CTkButton(
            address_buttons,
            text="Open Outstanding Cases Export",
            command=self.open_address_file
        )
        self.open_address_button.pack(
            side="left"
        )
        self.open_map_button = ctk.CTkButton(
            address_buttons,
            text="Open Map",
            command=self.open_map_file
        )
        self.open_map_button.pack(side="left", padx=(10, 0))
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
        self.step1_frame.status_label.configure(text="Ready")
        required = [CASE_FILE, RMA_FILE]
        missing = [p.name for p in required if not p.is_file()]
        self.step2_frame.status_label.configure(
            text="Salesforce reports organized" if not missing
            else "Missing: " + ", ".join(missing)
        )
        self.step3_frame.status_label.configure(
            text="CoreWeave report generated" if OUTPUT_FILE.is_file()
            else "Not generated"
        )
        if not SITE_MAP_FILE.is_file():
            validation_text = f"Missing CoreWeave map: {SITE_MAP_FILE.name}"
        elif not CASE_FILE.is_file():
            validation_text = f"Missing outstanding cases export: {CASE_FILE.name}"
        elif self.validation_results is None:
            validation_text = "Ready to validate against CoreWeave_Map.json"
        else:
            r = self.validation_results
            matches = r['match_counts'].get('matched', 0)
            unknown = r['new_location_count']
            conflicts = r['match_counts'].get('conflict', 0)
            fallback = r.get('site_resolution', {}).get('locode_fallback', 0)
            shipping_count = r.get('site_resolution', {}).get('shipping', 0)
            validation_text = (
                f"{matches:,} matched ({fallback:,} LoCode, {shipping_count:,} Shipping) | "
                f"{unknown:,} unresolved/conflicting address keys | {conflicts:,} conflicting cases"
            )
        self.step4_frame.status_label.configure(text=validation_text)
        finalized = list(OUTPUT_FOLDER.glob("CoreWeave Outstanding Cases *.xlsx"))
        if finalized:
            latest = max(finalized, key=lambda p: p.stat().st_mtime)
            self.step5_frame.status_label.configure(text=f"Latest final report: {latest.name}")
        else:
            self.step5_frame.status_label.configure(text="No finalized report yet")
    def refresh(self):
        if CASE_FILE.is_file() and SITE_MAP_FILE.is_file():
            self.check_address_file(update_status=False)
        else:
            self.validation_results = None
            self.address_textbox.delete("1.0", "end")
            if not SITE_MAP_FILE.is_file():
                message = f"Missing authoritative map:\n{SITE_MAP_FILE}"
            else:
                message = (
                    "Missing outstanding cases export.\n"
                    "Organize the Salesforce reports to create case_input.xlsx."
                )
            self.address_textbox.insert("end", message)
        self.update_workflow_status()
        self.set_status("Workflow refreshed")
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
            CASE_FILE,
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
    def get_address_validation(self):
        """Compare today's Salesforce export with the sole authoritative map."""
        if not SITE_MAP_FILE.is_file():
            raise FileNotFoundError(f"CoreWeave map is missing: {SITE_MAP_FILE}")
        if not CASE_FILE.is_file():
            raise FileNotFoundError(f"Outstanding cases export is missing: {CASE_FILE}")
        # Create a fresh validator so edits to CoreWeave_Map.json take effect.
        # The current outstanding export contains Asset Location but not the
        # optional City/State/Region columns expected by the legacy validator.
        # Build a temporary validation view from THAT export only, never a
        # second address authority or persistent address cache.
        source = self.read_salesforce_excel(CASE_FILE)
        for column in ("City", "State", "Region"):
            if column not in source.columns:
                source[column] = ""
        with tempfile.TemporaryDirectory(prefix="cw_case_validation_") as folder:
            validation_input = Path(folder) / "outstanding_validation.xlsx"
            source.to_excel(validation_input, index=False)
            validator = AddressValidator(validation_input, site_map_path=SITE_MAP_FILE)
            results = validator.validate()
            results["map_data"] = validator.site_map
        return results
    @staticmethod
    def _clean(value):
        if value is None or pd.isna(value):
            return ""
        value = str(value).strip()
        return "" if value.casefold() in {"nan", "none", "null"} else value
    @classmethod
    def address_review_data(cls, results):
        """Build read-only review tables. Never infer or change a canonical address."""
        raw = results["dataframe"]
        mapped = results["mapped"]
        location_fields = ("City", "State", "Region")
        blank = raw.loc[
            raw["Asset Location"].map(cls._clean).eq("")
            & raw["Case Number"].map(cls._clean).ne("")
            & raw["Case Number"].map(cls._clean).str.casefold().ne("total")
        ].copy()
        unknown = mapped.loc[mapped["Match Status"].isin(["unresolved", "unmatched", "conflict"])].copy()
        unresolved_rows = []
        for (key, status), group in unknown.groupby(["Address Key", "Match Status"], dropna=False):
            unresolved_rows.append({
                "Asset Location": cls._clean(group["Asset Location"].iloc[0]),
                "Match Status": status,
                "Cases": len(group),
                "Case Numbers": ", ".join(dict.fromkeys(group["Case Number"].map(cls._clean))),
                "Address Key": key,
            })
        unresolved = pd.DataFrame(unresolved_rows, columns=[
            "Asset Location", "Match Status", "Cases", "Case Numbers", "Address Key"
        ])
        if not unresolved.empty:
            unresolved = unresolved.sort_values(
                ["Cases", "Asset Location"], ascending=[False, True]
            ).reset_index(drop=True)
        shipping = results.get('shipping_locations', mapped.iloc[0:0]).copy()
        shipping_summary = pd.DataFrame(columns=['Asset Location', 'Cases', 'Report Site'])
        if not shipping.empty:
            shipping_summary = (shipping.groupby('Standardized Address', dropna=False)
                                .size().reset_index(name='Cases')
                                .rename(columns={'Standardized Address': 'Asset Location'}))
            shipping_summary['Report Site'] = 'Shipping'
            shipping_summary = shipping_summary.sort_values(
                ['Cases', 'Asset Location'], ascending=[False, True]
            ).reset_index(drop=True)
        map_missing_rows = []
        mismatched_rows = []
        export_missing_rows = []
        matches = mapped.loc[mapped["Match Status"].eq("matched")]
        for field in location_fields:
            expected = matches[f"Expected {field}"].map(cls._clean)
            observed = matches[field].map(cls._clean)
            for _, row in matches.loc[expected.eq("")].iterrows():
                map_missing_rows.append({
                    "Asset Location": row.get("Standardized Address", ""),
                    "LoCode": row.get("Matched LoCode", ""),
                    "Missing Map Field": field,
                    "Case Number": row.get("Case Number", ""),
                })
            compare = expected.str.casefold().ne(observed.str.casefold())
            for _, row in matches.loc[expected.ne("") & observed.ne("") & compare].iterrows():
                mismatched_rows.append({
                    "Case Number": row.get("Case Number", ""),
                    "Asset Location": row.get("Asset Location", ""),
                    "Field": field,
                    "Salesforce Value": row.get(field, ""),
                    "CoreWeave Map Value": row.get(f"Expected {field}", ""),
                })
        # Show missing Salesforce fields even when the incoming address is unknown.
        for field in location_fields:
            for _, row in mapped.loc[mapped[field].map(cls._clean).eq("")].iterrows():
                export_missing_rows.append({
                    "Case Number": row.get("Case Number", ""),
                    "Asset Location": row.get("Asset Location", ""),
                    "Missing Salesforce Field": field,
                    "Map Value": row.get(f"Expected {field}", ""),
                    "Match Status": row.get("Match Status", ""),
                })
        map_missing = pd.DataFrame(map_missing_rows, columns=[
            "Asset Location", "LoCode", "Missing Map Field", "Case Number"
        ])
        if not map_missing.empty:
            map_missing = map_missing.drop_duplicates(
                ["Asset Location", "Missing Map Field"]
            ).reset_index(drop=True)
        mismatched = pd.DataFrame(mismatched_rows, columns=[
            "Case Number", "Asset Location", "Field", "Salesforce Value", "CoreWeave Map Value"
        ])
        export_missing = pd.DataFrame(export_missing_rows, columns=[
            "Case Number", "Asset Location", "Missing Salesforce Field", "Map Value", "Match Status"
        ])
        supplied_by_map = export_missing.loc[
            export_missing["Match Status"].eq("matched")
            & export_missing["Map Value"].map(cls._clean).ne("")
        ].copy()
        export_missing = export_missing.drop(index=supplied_by_map.index).reset_index(drop=True)
        supplied_by_map = supplied_by_map.reset_index(drop=True)
        # Audit the entire authoritative map, including sites without current cases.
        # These are review findings, never a second source of address values.
        map_data = results.get("map_data") or {}
        map_gaps_rows = []
        for record in map_data.get("records", []):
            # Site/LoCode and optional contact fields are not required for a
            # known shipping destination. Count these separately from gaps.
            if cls._clean(record.get('address_type')).casefold() in {'shipping', 'shipping_only'}:
                continue
            for field in ("site", "locode", "address", "city", "state", "region",
                          "facility", "data_hall", "customer"):
                if not cls._clean(record.get(field)):
                    map_gaps_rows.append({
                        "LoCode": record.get("locode", ""),
                        "Site": record.get("site", ""),
                        "Address": record.get("address", ""),
                        "Missing Field": field,
                        "Priority": (
                            "No identifier" if field in {"site", "locode"}
                            and not cls._clean(record.get("site"))
                            and not cls._clean(record.get("locode"))
                            else "LoCode fallback" if field == "site"
                            and cls._clean(record.get("locode"))
                            else "Site available" if field == "locode"
                            and cls._clean(record.get("site"))
                            else "Required" if field in {"address", "city", "state", "region"}
                            else "Optional"
                        ),
                        "Source": record.get("source", ""),
                    })
        map_gaps = pd.DataFrame(map_gaps_rows, columns=[
            "LoCode", "Site", "Address", "Missing Field", "Priority", "Source"
        ]).drop_duplicates().reset_index(drop=True)
        if not map_gaps.empty:
            map_gaps["_sort"] = map_gaps["Priority"].map({
                "No identifier": 0, "Required": 1,
                "LoCode fallback": 2, "Site available": 3, "Optional": 4,
            }).fillna(5)
            map_gaps = map_gaps.sort_values(
                ["_sort", "Missing Field", "LoCode", "Address"]
            ).drop(columns="_sort").reset_index(drop=True)
        conflict_rows = []
        for kind, index in map_data.get("conflicts", {}).items():
            if not isinstance(index, dict):
                continue
            for key, details in index.items():
                details = details if isinstance(details, dict) else {}
                conflict_rows.append({
                    "Type": kind,
                    "Key": key,
                    "Conflicting Sites": " | ".join(details.get("sites", [])),
                    "Sources": " | ".join(details.get("sources", [])),
                    "Reason": details.get("reason", ""),
                })
        map_conflicts = pd.DataFrame(conflict_rows, columns=[
            "Type", "Key", "Conflicting Sites", "Sources", "Reason"
        ])
        # Actionable location-level summary; detailed case rows are retained in
        # Site Issues and Blank Addresses for the exported audit workbook.
        exceptions = results.get('address_exceptions')
        if exceptions is None:
            exceptions = AddressValidator.summarize_address_exceptions(
                results['site_issues'], blank
            )
        # Preserve all observations in the exported review; separate from map.
        site_issues = results["site_issues"].copy()
        inconsistent = results["inconsistent"].copy()
        return {
            "Address Exceptions": exceptions,
            "Map Field Gaps": map_gaps,
            "Map Conflicts": map_conflicts,
            "Blank Addresses": blank,
            "Shipping Locations": shipping_summary,
            "Unresolved Addresses": unresolved,
            "Missing in Map": map_missing,
            "Missing in Salesforce": export_missing,
            "Supplied by Map": supplied_by_map,
            "Location Mismatches": mismatched,
            "Site Issues": site_issues,
            "Source Inconsistencies": inconsistent,
            "Mapped Records": mapped,
        }
    @classmethod
    def format_address_review(cls, results, tables, limit=30):
        matched = results["match_counts"].get("matched", 0)
        conflict = results["match_counts"].get("conflict", 0)
        map_data = results.get("map_data") or {}
        map_gaps = tables["Map Field Gaps"]
        map_conflicts = tables["Map Conflicts"]
        unres = results["match_counts"].get("unresolved", 0) + results["match_counts"].get("unmatched", 0)
        total = results["total_rows"]
        unresolved = tables["Unresolved Addresses"]
        blank = tables["Blank Addresses"]
        shipping = tables["Shipping Locations"]
        map_missing = tables["Missing in Map"]
        missing_export = tables["Missing in Salesforce"]
        mismatched = tables["Location Mismatches"]
        site_issues = tables["Site Issues"]
        inconsistent = tables["Source Inconsistencies"]
        map_records = results["mapped"]
        matched_without_site = map_records.loc[
            map_records["Site Verification"].eq("No Site or LoCode in Map")
        ]
        unique_site_missing = matched_without_site["Address Key"].nunique()
        lines = [
            "COREWEAVE ADDRESS VALIDATION (READ ONLY)",
            "=" * 68,
            f"Source: {CASE_FILE.name}",
            f"Authority: {SITE_MAP_FILE.name}",
            "",
            "AUTHORITATIVE MAP HEALTH",
            "-" * 68,
            f"Map records: {len(map_data.get('records', [])):,}",
            f"Address lookup keys: {len(map_data.get('by_address', {})):,}",
            f"Known map conflicts: {len(map_conflicts):,}",
            f"Known shipping destinations in map: {sum(str(r.get('address_type', '')).strip().casefold() in ('shipping', 'shipping_only') for r in map_data.get('records', [])):,}",
            f"Locations missing BOTH Site and LoCode: {map_gaps.loc[map_gaps['Priority'].eq('No identifier'), 'Address'].nunique():,}",
            f"Map records with LoCode fallback: {len(map_gaps.loc[map_gaps['Priority'].eq('LoCode fallback')]):,}",
            f"Other required geographic/address gaps: {len(map_gaps.loc[map_gaps['Priority'].eq('Required')]):,}",
            f"Missing optional map fields: {len(map_gaps.loc[map_gaps['Priority'].eq('Optional')]):,} record/field gaps",
            "",
            "CURRENT SALESFORCE EXPORT",
            "-" * 68,
            f"Cases with an address: {total:,}",
            f"Unique incoming addresses: {results['unique_locations']:,}",
            f"Matched addressed cases: {matched:,}",
            f"  With map Site: {results.get('site_resolution', {}).get('map_site', 0):,}",
            f"  With LoCode fallback: {results.get('site_resolution', {}).get('locode_fallback', 0):,}",
            f"  With Shipping classification: {results.get('site_resolution', {}).get('shipping', 0):,}",
            f"  Without either identifier: {results.get('site_resolution', {}).get('no_identifier', 0):,}",
            f"Unmatched cases: {unres:,}",
            f"Conflicting cases: {conflict:,}",
            f"Unique unresolved/conflicting addresses: {len(unresolved):,}",
            "",
            "WHAT NEEDS REVIEW",
            "-" * 68,
            f"Blank Salesforce addresses: {len(blank):,} cases",
            f"Missing location fields AFTER map lookup: {results['missing_count']:,} cases",
            f"Blank Salesforce fields supplied by map: {len(tables['Supplied by Map']):,} fields",
            f"Blank Salesforce fields still unresolved: {len(missing_export):,} fields",
            f"Missing geographic fields in map: {len(map_missing):,} address/field gaps",
            f"Physical locations without Site or LoCode: {unique_site_missing:,} distinct matched address keys",
            f"Salesforce vs map location mismatches: {len(mismatched):,} field differences",
            f"Site verification issues (including unknown addresses): {len(site_issues):,} cases",
            f"Conflicting observed location groups: {len(inconsistent):,}",
        ]
        if "Site" not in results["dataframe"].columns:
            lines.append("Site comparison: no Site column in outstanding cases export; map Site gaps still checked")
        def section(title, rows, formatter):
            lines.extend(["", title, "-" * 68])
            if rows.empty:
                lines.append("None")
                return
            for _, row in rows.head(limit).iterrows():
                lines.append(formatter(row))
            if len(rows) > limit:
                lines.append(f"... {len(rows) - limit:,} more (use Export Review for all)")
        section("ADDRESS EXCEPTIONS (GROUPED BY LOCATION, LARGEST FIRST)",
                tables["Address Exceptions"],
                lambda r: f"{r['Cases']:>5} cases | {r['Review Reason']} | {r['Asset Location']}")
        lines.append("Case Numbers and all case-level details are in Export Review.")
        section("MISSING FIELDS ACROSS ALL MAP RECORDS", map_gaps,
                lambda r: f"{r['Priority']} | {r['Missing Field']} | {r['LoCode']} | {r['Address']}")
        section("KNOWN CONFLICTS IN THE MAP", map_conflicts,
                lambda r: f"{r['Type']} | {r['Key']} | {r['Conflicting Sites']}")
        section("KNOWN SHIPPING LOCATIONS (INFORMATIONAL, NOT ERRORS)", shipping,
                lambda r: f"{r['Cases']:>5} cases | Shipping | {r['Asset Location']}")
        section("MISSING INFORMATION IN COREWEAVE MAP", map_missing,
                lambda r: f"{r['Missing Map Field']} | {r['LoCode']} | {r['Asset Location']}")
        # These tables are field-level details; aggregate them by address/field
        # rather than showing thousands of repeated Salesforce case IDs.
        if not missing_export.empty:
            source_gap_summary = (missing_export.groupby(
                ['Asset Location', 'Missing Salesforce Field', 'Match Status'],
                dropna=False).size().reset_index(name='Cases')
                .sort_values(['Cases', 'Asset Location'], ascending=[False, True]))
        else:
            source_gap_summary = missing_export
        section("SOURCE FIELDS NOT SUPPLIED BY THE MAP (GROUPED)", source_gap_summary,
                lambda r: f"{r['Cases']:>5} cases | {r['Missing Salesforce Field']} | {r['Match Status']} | {r['Asset Location']}")
        lines.extend(["", "BLANK SOURCE FIELDS SUPPLIED BY MAP (INFORMATIONAL)", "-" * 68,
                      f"{len(tables['Supplied by Map']):,} fields supplied; see Export Review for case details."])
        if not mismatched.empty:
            mismatched_summary = (mismatched.groupby(['Asset Location', 'Field'], dropna=False)
                                  .size().reset_index(name='Cases')
                                  .sort_values(['Cases', 'Asset Location'], ascending=[False, True]))
        else:
            mismatched_summary = mismatched
        section("LOCATION MISMATCHES (SALESFORCE VS MAP)", mismatched_summary,
                lambda r: f"{r['Cases']:>5} cases | {r['Field']} | {r['Asset Location']}")
        section("INCONSISTENT VALUES IN SALESFORCE", inconsistent,
                lambda r: f"{r.get('Field', '')} | {r.get('Asset Location', '')} | {r.get('Observed Values', '')}")
        lines.extend(["", "This validates against the map; it does not change any source data."])
        return "\n".join(lines)
    def check_address_file(self, update_status=True):
        try:
            results = self.get_address_validation()
            tables = self.address_review_data(results)
            self.validation_results = results
            self.address_textbox.delete("1.0", "end")
            self.address_textbox.insert("end", self.format_address_review(results, tables))
            if update_status:
                self.set_status(
                    f"Validation: {results['match_counts'].get('matched', 0):,} matched, "
                    f"{len(tables['Address Exceptions']):,} address exception groups, "
                    f"{len(tables['Map Field Gaps']):,} map-wide field gaps."
                )
        except Exception as exc:
            self.validation_results = None
            self.address_textbox.delete("1.0", "end")
            self.address_textbox.insert("end", f"Address validation error:\n\n{exc}")
            if update_status:
                self.set_status(f"Address validation error: {exc}")
        self.update_workflow_status()
    def export_address_review(self):
        try:
            results = self.get_address_validation()  # Fresh read of the current map.
            tables = self.address_review_data(results)
            # Keep address diagnostics outside the shareable Output directory.
            DEBUG_FOLDER.mkdir(parents=True, exist_ok=True)
            output = DEBUG_FOLDER / "CoreWeave_Address_Validation_Review.xlsx"
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                summary = pd.DataFrame([
                    {"Metric": "Authoritative map", "Value": SITE_MAP_FILE.name},
                    {"Metric": "Outstanding cases export", "Value": CASE_FILE.name},
                    {"Metric": "Map records", "Value": len(results.get("map_data", {}).get("records", []))},
                    {"Metric": "Shipping destinations (map records)", "Value": len([r for r in results.get("map_data", {}).get("records", []) if str(r.get('address_type', '')).strip().casefold() in {'shipping', 'shipping_only'}])},
                    {"Metric": "Map missing fields", "Value": len(tables["Map Field Gaps"])},
                    {"Metric": "Map conflicts", "Value": len(tables["Map Conflicts"])},
                    {"Metric": "Valid address rows", "Value": results["total_rows"]},
                    {"Metric": "Matched", "Value": results["match_counts"].get("matched", 0)},
                    {"Metric": "Unmatched", "Value": results["match_counts"].get("unresolved", 0)},
                    {"Metric": "Conflicts", "Value": results["match_counts"].get("conflict", 0)},
                    {"Metric": "Unique unresolved", "Value": len(tables["Unresolved Addresses"])},
                    {"Metric": "Blank Asset Locations", "Value": len(tables["Blank Addresses"])},
                    {"Metric": "Missing in map (address/field)", "Value": len(tables["Missing in Map"])},
                    {"Metric": "Location fields still missing after map lookup (cases)", "Value": results["missing_count"]},
                    {"Metric": "Source fields supplied by map", "Value": len(tables["Supplied by Map"])},
                    {"Metric": "Location mismatches (field)", "Value": len(tables["Location Mismatches"])},
                    {"Metric": "Site issues (case)", "Value": len(tables["Site Issues"])},
                    {"Metric": "Grouped address exceptions (locations/reasons)", "Value": len(tables["Address Exceptions"])},
                    {"Metric": "Address exception cases (including blank addresses)", "Value": int(tables["Address Exceptions"]["Cases"].sum())},
                    {"Metric": "Map Site available (cases)", "Value": results.get('site_resolution', {}).get('map_site', 0)},
                    {"Metric": "LoCode fallback available (cases)", "Value": results.get('site_resolution', {}).get('locode_fallback', 0)},
                    {"Metric": "Known shipping cases (informational)", "Value": results.get('site_resolution', {}).get('shipping', 0)},
                    {"Metric": "Neither Site nor LoCode (cases)", "Value": results.get('site_resolution', {}).get('no_identifier', 0)},
                ])
                summary.to_excel(writer, sheet_name="Summary", index=False)
                for name, data in tables.items():
                    data.to_excel(writer, sheet_name=name[:31], index=False)
                for ws in writer.book.worksheets:
                    ws.freeze_panes = "A2"
                    ws.sheet_view.showGridLines = False
                    for cell in ws[1]:
                        cell.font = cell.font.copy(bold=True)
                    for column in ws.columns:
                        letter = column[0].column_letter
                        width = max((len(str(c.value)) for c in column[:101]
                                     if c.value is not None), default=12)
                        ws.column_dimensions[letter].width = max(15, min(width + 2, 65))
            self.set_status(f"Review exported: {output.name} (map unchanged)")
        except Exception as exc:
            self.set_status(f"Could not export review: {exc}")
    def open_address_file(self):
        if not CASE_FILE.is_file():
            self.set_status(f"Missing outstanding cases export: {CASE_FILE}")
            return
        try:
            os.startfile(CASE_FILE)
            self.set_status(f"Opened {CASE_FILE.name}")
        except Exception as exc:
            self.set_status(f"Could not open outstanding cases export: {exc}")
    def open_map_file(self):
        if not SITE_MAP_FILE.is_file():
            self.set_status(f"Missing authoritative CoreWeave map: {SITE_MAP_FILE}")
            return
        try:
            os.startfile(SITE_MAP_FILE)
            self.set_status(f"Opened {SITE_MAP_FILE.name}")
        except Exception as exc:
            self.set_status(f"Could not open CoreWeave map: {exc}")
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
        # MAKE SURE CURRENT OUTSTANDING CASES EXPORT EXISTS
        # ----------------------------------------------------
        if not CASE_FILE.is_file():
            self.set_status(
                "case_input.xlsx was not found. Organize Salesforce downloads first."
            )
            return
        try:
            # ------------------------------------------------
            # VALIDATE CURRENT SALESFORCE ADDRESS INFORMATION AGAINST MAP
            # ------------------------------------------------
            results = (
                self.get_address_validation()
            )
            tables = self.address_review_data(results)
            unknown_count = len(tables["Unresolved Addresses"])
            blank_count = len(tables["Blank Addresses"])
            map_gap_count = len(tables["Missing in Map"])
            site_gap_count = results["mapped"].loc[
                results["mapped"]["Site Verification"].eq("No Site or LoCode in Map"),
                "Address Key",
            ].nunique()
            if unknown_count or blank_count or map_gap_count or site_gap_count:
                self.check_address_file(update_status=False)
                dialog = ctk.CTkInputDialog(
                    title="Address Validation Warnings",
                    text=(
                        "The CoreWeave map cannot fully verify this Salesforce export:\n\n"
                        f"{unknown_count} unresolved/conflicting addresses\n"
                        f"{blank_count} cases without an address\n"
                        f"{map_gap_count} missing map location fields\n"
                        f"{site_gap_count} matched addresses without Site or LoCode\n\n"
                        "These records will remain flagged, not automatically corrected.\n"
                        "Type YES to finalize with these warnings:"
                    ),
                )
                response = dialog.get_input()
                if not response or response.strip().upper() != "YES":
                    self.set_status("Finalization cancelled. Review the address issues first.")
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
                "Address validation reviewed. Regenerating final report..."
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
                self.get_address_validation()
            )
            # This is a read-only recheck. Unknown or incomplete sites remain
            # flagged in the generated report, but are not auto-corrected.
            self.set_worker_status(
                f"Validated against CoreWeave map: "
                f"{results['match_counts'].get('matched', 0)} matched cases, "
                f"{results['new_location_count']} unresolved address keys."
            )
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
            # Remove legacy embedded diagnostics if finalizing an older preview.
            # The current generator writes them only to the Debug folder.
            wb = load_workbook(final_file)
            try:
                if 'Location Review' in wb.sheetnames:
                    del wb['Location Review']
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
            # Record history only after the finalized workbook is saved.
            commit_finalized_history(
                final_file,
                status_callback=self.set_status,
            )
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
                "Reports/Coreweave_Outstanding/Input\n"
                "Reports/Coreweave_Outstanding/Output\n\n"
                "CoreWeave_Map.json will NOT be affected.\n\n"
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
                    "CoreWeave_Map.json was preserved."
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
