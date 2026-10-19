import customtkinter as ctk
import subprocess
import sys
from pathlib import Path


class MainMenu(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("CoreWeave Reporting Tool")
        self.geometry("600x450")
        self.resizable(False, False)

        self.project_folder = Path(__file__).resolve().parent.parent

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Title
        ctk.CTkLabel(
            self,
            text="CoreWeave Reporting Tool",
            font=ctk.CTkFont(size=26, weight="bold")
        ).pack(pady=(35, 5))

        ctk.CTkLabel(
            self,
            text="Select a reporting workflow",
            font=ctk.CTkFont(size=14)
        ).pack(pady=(0, 25))

        # Report buttons
        self.create_report_button(
            title="Outstanding Cases Report",
            description="Generate and validate the outstanding cases report",
            command=self.open_outstanding_reports
        )

        self.create_report_button(
            title="Bi-weekly Dashboard",
            description="Prepare data for the CoreWeave Power BI dashboard",
            command=self.open_dashboard
        )

        # Footer
        ctk.CTkLabel(
            self,
            text="CoreWeave Reporting Tool",
            text_color="gray"
        ).pack(side="bottom", pady=15)

    def create_report_button(self, title, description, command):
        button = ctk.CTkButton(
            self,
            text=f"{title}\n{description}",
            command=command,
            height=75,
            font=ctk.CTkFont(size=15),
            corner_radius=10
        )
        button.pack(fill="x", padx=70, pady=8)

    def launch_script(self, script_name):
        script_path = self.project_folder / script_name

        if not script_path.exists():
            ctk.CTkMessagebox if False else None
            from tkinter import messagebox
            messagebox.showerror(
                "File Not Found",
                f"Could not find:\n{script_path}"
            )
            return

        subprocess.Popen(
            [sys.executable, str(script_path)],
            cwd=str(self.project_folder)
        )

    def open_outstanding_reports(self):
        self.launch_script(
            "Reports/Coreweave_Outstanding/OutstandingReport_GUI.py"
        )

    def open_dashboard(self):
        self.launch_script(
            "Reports/QBR_Dashboard/QBR_Dashboard_GUI.py"
        )


if __name__ == "__main__":
    app = MainMenu()
    app.mainloop()
