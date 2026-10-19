
from pathlib import Path
import os
import shutil
import threading
import webbrowser
import queue
import subprocess
import sys
from tkinter import messagebox

# Project root: Coreweave Outstanding Reports
BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

import customtkinter as ctk

from Utilities.AddressValidator import AddressValidator

# ============================================================
# CONFIGURATION
# ============================================================

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

DASHBOARD_FOLDER = BASE_DIR / "Reports" / "QBR_Dashboard"

DASHBOARD_SCRIPT = DASHBOARD_FOLDER / "QBR_Dashboard.py"
INSTALLED_BASES_SCRIPT = DASHBOARD_FOLDER / "InstalledBases.py"

ADDRESS_FILE = BASE_DIR / "Address Cache" / "address_input.xlsx"
DOWNLOADS_FOLDER = Path.home() / "Downloads"

INPUT_FOLDER = DASHBOARD_FOLDER / "Input"
OUTPUT_FOLDER = DASHBOARD_FOLDER / "Output"

OUTPUT_FILE = OUTPUT_FOLDER / "Cases.xlsx"
INSTALLED_BASES_OUTPUT = OUTPUT_FOLDER / "Installed Bases.xlsx"

REPORT_URL = (
    "https://supermicrocomputer.lightning.force.com/"
    "lightning/r/Report/00OPm00000MP9hlMAD/view"
    "?queryScope=userFolders"
)

REPORT_NAME_PATTERN = "CoreWeave QBR Dashboard-*.xlsx"

INPUT_FOLDER.mkdir(parents=True, exist_ok=True)
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)


# ============================================================
# MAIN APPLICATION
# ============================================================

class CoreWeaveDashboardGUI(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.address_validator = AddressValidator(ADDRESS_FILE)
        self.busy = False
        self.log_queue = queue.Queue()

        self.title("CoreWeave QBR Dashboard")
        self.geometry("1000x1000")
        self.minsize(900, 750)

        self.build_ui()
        self.refresh()
        self.after(100, self.poll_worker_queue)

    # ========================================================
    # UI
    # ========================================================

    def build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self)
        header.grid(
            row=0, column=0,
            padx=20, pady=(20, 10),
            sticky="ew"
        )
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="CoreWeave QBR Dashboard",
            font=ctk.CTkFont(size=26, weight="bold")
        ).grid(
            row=0, column=0,
            padx=20, pady=(15, 2),
            sticky="w"
        )

        ctk.CTkLabel(
            header,
            text="Salesforce → Cases.xlsx → Installed Bases.xlsx → Power BI",
            font=ctk.CTkFont(size=14)
        ).grid(
            row=1, column=0,
            padx=20, pady=(0, 15),
            sticky="w"
        )

        ctk.CTkButton(
            header,
            text="Refresh",
            width=100,
            command=self.refresh
        ).grid(
            row=0, column=1,
            rowspan=2, padx=20
        )

        self.main_frame = ctk.CTkScrollableFrame(self)
        self.main_frame.grid(
            row=1, column=0,
            padx=20, pady=(0, 10),
            sticky="nsew"
        )
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.build_workflow()

        self.status_label = ctk.CTkLabel(
            self,
            text="Ready",
            anchor="w"
        )
        self.status_label.grid(
            row=2, column=0,
            padx=20, pady=(0, 15),
            sticky="ew"
        )

    # ========================================================
    # WORKFLOW
    # ========================================================

    def build_workflow(self):

        # STEP 1: SALESFORCE
        self.step1_frame = self.create_step(
            0,
            "1. Open Salesforce QBR Report",
            "Open the Salesforce report and export the data to Excel."
        )

        self.open_report_button = ctk.CTkButton(
            self.step1_frame,
            text="Open Salesforce Report",
            command=self.open_salesforce_report
        )
        self.open_report_button.pack(
            anchor="w", padx=20, pady=(0, 15)
        )

        # STEP 2: MOVE EXPORT
        self.step2_frame = self.create_step(
            1,
            "2. Move Salesforce Export",
            "Move the newest matching QBR Excel export from Downloads "
            "into the project's Input folder."
        )

        self.move_button = ctk.CTkButton(
            self.step2_frame,
            text="Move Latest Export to Input",
            command=self.move_latest_export
        )
        self.move_button.pack(
            anchor="w", padx=20, pady=(0, 15)
        )

        # STEP 3: CASES
        self.step3_frame = self.create_step(
            2,
            "3. Generate Cases.xlsx",
            "Run QBR_Dashboard.py to process the Salesforce export, "
            "format case data, and enrich City, State, and Region."
        )

        self.process_button = ctk.CTkButton(
            self.step3_frame,
            text="Generate Cases.xlsx",
            command=self.run_cases_script
        )
        self.process_button.pack(
            anchor="w", padx=20, pady=(0, 15)
        )

        # STEP 4: INSTALLED BASES
        self.step4_frame = self.create_step(
            3,
            "4. Generate Installed Bases.xlsx",
            "Query the GSP database, enrich location data, and generate "
            "the Installed Bases workbook."
        )

        self.installed_bases_button = ctk.CTkButton(
            self.step4_frame,
            text="Generate Installed Bases.xlsx",
            command=self.run_installed_bases_script
        )
        self.installed_bases_button.pack(
            anchor="w", padx=20, pady=(0, 15)
        )

        # STEP 5: ADDRESS VALIDATION
        self.step5_frame = self.create_step(
            4,
            "5. Validate Address Cache",
            "Check address_input.xlsx for missing or inconsistent "
            "City, State, and Region values."
        )

        self.address_textbox = ctk.CTkTextbox(
            self.step5_frame,
            height=200,
            font=ctk.CTkFont(family="Consolas", size=12)
        )
        self.address_textbox.pack(
            fill="x", padx=20, pady=(0, 10)
        )

        address_buttons = ctk.CTkFrame(
            self.step5_frame,
            fg_color="transparent"
        )
        address_buttons.pack(
            fill="x", padx=20, pady=(0, 15)
        )

        self.validate_button = ctk.CTkButton(
            address_buttons,
            text="Validate Addresses",
            command=self.check_address_file
        )
        self.validate_button.pack(
            side="left", padx=(0, 10)
        )

        self.open_address_button = ctk.CTkButton(
            address_buttons,
            text="Open Address Cache",
            command=self.open_address_file
        )
        self.open_address_button.pack(side="left")

        # STEP 6: OUTPUT AND CLEANUP
        self.step6_frame = self.create_step(
            5,
            "6. Open Output Folder",
            "Access Cases.xlsx, Installed Bases.xlsx, and other generated "
            "files used by the Power BI dashboard."
        )

        output_buttons = ctk.CTkFrame(
            self.step6_frame,
            fg_color="transparent"
        )
        output_buttons.pack(
            fill="x", padx=20, pady=(0, 15)
        )

        self.open_output_button = ctk.CTkButton(
            output_buttons,
            text="Open Output Folder",
            command=self.open_output_folder
        )
        self.open_output_button.pack(
            side="left", padx=(0, 10)
        )

        self.clear_button = ctk.CTkButton(
            output_buttons,
            text="Clear Input / Output",
            fg_color="#A83232",
            hover_color="#7D2424",
            command=self.clear_input_output
        )
        self.clear_button.pack(side="left")

    # ========================================================
    # STEP CREATION
    # ========================================================

    def create_step(self, row, title, description):
        frame = ctk.CTkFrame(self.main_frame)
        frame.grid(
            row=row, column=0,
            padx=5, pady=7,
            sticky="ew"
        )
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            frame,
            text=title,
            font=ctk.CTkFont(size=18, weight="bold"),
            anchor="w"
        ).pack(
            fill="x", padx=20, pady=(15, 3)
        )

        ctk.CTkLabel(
            frame,
            text=description,
            anchor="w",
            justify="left",
            wraplength=850
        ).pack(
            fill="x", padx=20, pady=(0, 10)
        )

        status_label = ctk.CTkLabel(
            frame,
            text="Checking...",
            anchor="w"
        )
        status_label.pack(
            fill="x", padx=20, pady=(0, 10)
        )

        frame.status_label = status_label
        return frame

    # ========================================================
    # STATUS AND REFRESH
    # ========================================================

    def set_status(self, message):
        self.status_label.configure(text=message)

    def set_step_status(self, frame, message):
        frame.status_label.configure(text=message)

    def refresh(self):
        self.set_step_status(
            self.step1_frame,
            "Ready to open report"
        )

        # Check for an export already in Input.
        input_exports = list(
            INPUT_FOLDER.glob(REPORT_NAME_PATTERN)
        )

        if input_exports:
            newest_input = max(
                input_exports,
                key=lambda p: p.stat().st_mtime
            )
            self.set_step_status(
                self.step2_frame,
                f"Export in Input: {newest_input.name}"
            )
        else:
            download_exports = list(
                DOWNLOADS_FOLDER.glob(REPORT_NAME_PATTERN)
            )

            if download_exports:
                self.set_step_status(
                    self.step2_frame,
                    f"{len(download_exports)} matching export(s) in Downloads"
                )
            else:
                self.set_step_status(
                    self.step2_frame,
                    "No matching export found in Downloads or Input"
                )

        self.set_step_status(
            self.step3_frame,
            "Cases.xlsx available"
            if OUTPUT_FILE.exists()
            else "Cases.xlsx has not been generated"
        )

        self.set_step_status(
            self.step4_frame,
            "Installed Bases.xlsx available"
            if INSTALLED_BASES_OUTPUT.exists()
            else "Installed Bases.xlsx has not been generated"
        )

        self.set_step_status(
            self.step6_frame,
            "Output folder ready"
        )

        self.check_address_file(update_status=False)

        self.set_status("Workflow refreshed")

    # ========================================================
    # STEP 1: OPEN SALESFORCE
    # ========================================================

    def open_salesforce_report(self):
        try:
            webbrowser.open(REPORT_URL)
            self.set_step_status(
                self.step1_frame,
                "Salesforce report opened in browser"
            )
            self.set_status(
                "Export the report to Excel before continuing."
            )
        except Exception as exc:
            self.set_status(f"Could not open Salesforce: {exc}")

    # ========================================================
    # STEP 2: MOVE EXPORT
    # ========================================================

    def move_latest_export(self):
        if self.busy:
            self.set_status("Wait for the current operation to finish.")
            return

        candidates = list(
            DOWNLOADS_FOLDER.glob(REPORT_NAME_PATTERN)
        )

        if not candidates:
            self.set_status(
                "No matching QBR export found in Downloads."
            )
            return

        newest_export = max(
            candidates,
            key=lambda p: p.stat().st_mtime
        )

        destination = INPUT_FOLDER / newest_export.name

        try:
            INPUT_FOLDER.mkdir(parents=True, exist_ok=True)

            # If the same filename is already in Input, replace it
            # with the newest export from Downloads.
            if destination.exists():
                destination.unlink()

            shutil.move(
                str(newest_export),
                str(destination)
            )

            self.set_step_status(
                self.step2_frame,
                f"✓ Moved {newest_export.name} to Input"
            )
            self.set_status("Salesforce export moved successfully.")

        except Exception as exc:
            self.set_step_status(
                self.step2_frame,
                f"⚠ Move failed: {exc}"
            )
            self.set_status(f"Could not move export: {exc}")

    # ========================================================
    # SCRIPT EXECUTION
    # ========================================================

    def start_script(
        self,
        script_path,
        output_file,
        step_frame,
        step_name,
        success_message
    ):
        if self.busy:
            self.set_status("Another operation is already running.")
            return

        if not script_path.exists():
            self.set_step_status(
                step_frame,
                f"⚠ Script not found: {script_path}"
            )
            self.set_status(f"Script not found: {script_path}")
            return

        self.busy = True
        self.set_step_status(
            step_frame,
            f"Running {script_path.name}..."
        )
        self.set_status(f"Running {step_name}...")

        self.set_action_buttons_enabled(False)

        threading.Thread(
            target=self.script_worker,
            args=(
                script_path,
                output_file,
                step_frame,
                step_name,
                success_message
            ),
            daemon=True
        ).start()

    def script_worker(
        self,
        script_path,
        output_file,
        step_frame,
        step_name,
        success_message
    ):
        try:
            result = subprocess.run(
                [sys.executable, str(script_path)],
                cwd=str(BASE_DIR),
                capture_output=True,
                text=True,
                check=False
            )

            if result.stdout:
                self.log_queue.put(("log", result.stdout))

            if result.stderr:
                self.log_queue.put(("log", result.stderr))

            if result.returncode != 0:
                raise RuntimeError(
                    result.stderr.strip()
                    or f"Script exited with code {result.returncode}"
                )

            if not output_file.exists():
                raise FileNotFoundError(
                    f"Script completed, but output was not found:\n"
                    f"{output_file}"
                )

            self.log_queue.put((
                "job_done",
                {
                    "success": True,
                    "frame": step_frame,
                    "step_name": step_name,
                    "message": success_message
                }
            ))

        except Exception as exc:
            self.log_queue.put((
                "job_done",
                {
                    "success": False,
                    "frame": step_frame,
                    "step_name": step_name,
                    "message": str(exc)
                }
            ))

    def run_cases_script(self):
        input_exports = list(
            INPUT_FOLDER.glob(REPORT_NAME_PATTERN)
        )

        if not input_exports:
            self.set_status(
                "No QBR export in Input. Complete Step 2 first."
            )
            return

        self.start_script(
            script_path=DASHBOARD_SCRIPT,
            output_file=OUTPUT_FILE,
            step_frame=self.step3_frame,
            step_name="Cases.xlsx generation",
            success_message="✓ Cases.xlsx generated successfully"
        )

    def run_installed_bases_script(self):
        self.start_script(
            script_path=INSTALLED_BASES_SCRIPT,
            output_file=INSTALLED_BASES_OUTPUT,
            step_frame=self.step4_frame,
            step_name="Installed Bases generation",
            success_message="✓ Installed Bases.xlsx generated successfully"
        )

    def set_action_buttons_enabled(self, enabled):
        state = "normal" if enabled else "disabled"

        for button in (
            self.open_report_button,
            self.move_button,
            self.process_button,
            self.installed_bases_button,
            self.validate_button,
            self.open_address_button,
            self.open_output_button,
            self.clear_button
        ):
            button.configure(state=state)

        self.process_button.configure(
            text="Generate Cases.xlsx" if enabled else "Processing..."
        )
        self.installed_bases_button.configure(
            text=(
                "Generate Installed Bases.xlsx"
                if enabled else "Processing..."
            )
        )

    # ========================================================
    # WORKER QUEUE
    # ========================================================

    def poll_worker_queue(self):
        try:
            while True:
                event, payload = self.log_queue.get_nowait()

                if event == "log":
                    # Script logs are printed to the terminal.
                    print(payload)

                elif event == "job_done":
                    self.busy = False
                    self.set_action_buttons_enabled(True)

                    frame = payload["frame"]
                    message = payload["message"]

                    if payload["success"]:
                        self.set_step_status(frame, message)
                        self.set_status(message)
                        self.refresh()
                    else:
                        self.set_step_status(
                            frame,
                            f"⚠ {payload['step_name']} failed: {message}"
                        )
                        self.set_status(
                            f"{payload['step_name']} failed. "
                            "See the terminal for details."
                        )
                        messagebox.showerror(
                            "Processing Failed",
                            f"{payload['step_name']} failed:\n\n{message}"
                        )

        except queue.Empty:
            pass

        self.after(100, self.poll_worker_queue)

    # ========================================================
    # STEP 5: ADDRESS VALIDATION
    # ========================================================

    def check_address_file(self, update_status=True):
        self.address_textbox.delete("1.0", "end")

        if not ADDRESS_FILE.exists():
            self.address_textbox.insert(
                "end",
                "address_input.xlsx was not found.\n"
                "Check the Address Cache folder."
            )
            self.set_step_status(
                self.step5_frame,
                "⚠ Address cache not found"
            )

            if update_status:
                self.set_status("Address cache not found.")
            return

        try:
            results = self.address_validator.validate()

            missing_df = results["missing"]
            inconsistent_df = results["inconsistent"]

            self.address_textbox.insert(
                "end",
                "ADDRESS VALIDATION SUMMARY\n"
                f"{'=' * 65}\n"
                f"Total rows: {results['total_rows']}\n"
                f"Unique locations: {results['unique_locations']}\n"
                f"Complete locations: {results['complete_locations']}\n"
                f"Missing locations: {results['missing_count']}\n"
                f"Inconsistent locations: {results['inconsistent_count']}\n\n"
            )

            if not missing_df.empty:
                self.address_textbox.insert(
                    "end",
                    "MISSING LOCATION DATA\n"
                    f"{'-' * 65}\n"
                )

                for _, row in missing_df.iterrows():
                    self.address_textbox.insert(
                        "end",
                        f"Case: {row['Case Number']} | "
                        f"Location: {row['Asset Location']}\n"
                        f"City: {row['City']} | "
                        f"State: {row['State']} | "
                        f"Region: {row['Region']}\n\n"
                    )

            if not inconsistent_df.empty:
                self.address_textbox.insert(
                    "end",
                    "INCONSISTENT LOCATIONS\n"
                    f"{'-' * 65}\n"
                )

                for _, row in inconsistent_df.iterrows():
                    majority = row["Majority"]

                    self.address_textbox.insert(
                        "end",
                        f"Asset Location: {row['Asset Location']}\n"
                        f"Majority: City={majority['City']}, "
                        f"State={majority['State']}, "
                        f"Region={majority['Region']}\n"
                    )

                    for detail in row["Rows"]:
                        self.address_textbox.insert(
                            "end",
                            f"  Excel row {detail['Excel Row']} | "
                            f"Case {detail['Case Number']} | "
                            f"{detail['City']}, {detail['State']}, "
                            f"{detail['Region']}\n"
                        )

                    self.address_textbox.insert("end", "\n")

            if missing_df.empty and inconsistent_df.empty:
                self.address_textbox.insert(
                    "end",
                    "✓ ADDRESS VALIDATION PASSED\n\n"
                    "All valid Asset Locations have complete, "
                    "consistent City, State, and Region values."
                )
                self.set_step_status(
                    self.step5_frame,
                    "✓ Address validation passed"
                )
            else:
                self.set_step_status(
                    self.step5_frame,
                    f"⚠ {len(missing_df)} missing records; "
                    f"{len(inconsistent_df)} inconsistent locations"
                )

            if update_status:
                if missing_df.empty and inconsistent_df.empty:
                    self.set_status("Address validation passed.")
                else:
                    self.set_status(
                        "Address issues found. Review the validation results."
                    )

        except Exception as exc:
            self.address_textbox.insert(
                "end",
                f"Error validating address cache:\n\n{exc}"
            )
            self.set_step_status(
                self.step5_frame,
                f"⚠ Validation error: {exc}"
            )

            if update_status:
                self.set_status(f"Validation error: {exc}")

    def open_address_file(self):
        if not ADDRESS_FILE.exists():
            self.set_status("Address cache file not found.")
            return

        try:
            os.startfile(ADDRESS_FILE)
            self.set_status("Opened address_input.xlsx.")
        except Exception as exc:
            self.set_status(f"Could not open address file: {exc}")

    # ========================================================
    # STEP 6: OUTPUT FOLDER
    # ========================================================

    def open_output_folder(self):
        try:
            OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
            os.startfile(OUTPUT_FOLDER)
            self.set_status("Opened dashboard output folder.")
        except Exception as exc:
            self.set_status(f"Could not open output folder: {exc}")

    # ========================================================
    # CLEANUP: INPUT / OUTPUT
    # ========================================================

    def clear_input_output(self):
        if self.busy:
            self.set_status("Wait for the current operation to finish.")
            return

        confirmed = messagebox.askyesno(
            "Confirm Cleanup",
            "This will permanently delete all files and subfolders inside:\n\n"
            f"Input:\n{INPUT_FOLDER}\n\n"
            f"Output:\n{OUTPUT_FOLDER}\n\n"
            "The folders themselves and your Address Cache will be preserved.\n\n"
            "Continue?"
        )

        if not confirmed:
            self.set_status("Cleanup cancelled.")
            return

        try:
            deleted_count = 0

            for folder in (INPUT_FOLDER, OUTPUT_FOLDER):
                folder.mkdir(parents=True, exist_ok=True)

                for item in folder.iterdir():
                    if item.is_dir():
                        shutil.rmtree(item)
                    else:
                        item.unlink()

                    deleted_count += 1

            self.set_status(
                f"Cleanup complete. Removed {deleted_count} item(s)."
            )
            self.refresh()

        except Exception as exc:
            self.set_status(f"Cleanup failed: {exc}")
            messagebox.showerror(
                "Cleanup Failed",
                str(exc)
            )


# ============================================================
# DIRECT RUN
# ============================================================

if __name__ == "__main__":
    app = CoreWeaveDashboardGUI()
    app.mainloop()
