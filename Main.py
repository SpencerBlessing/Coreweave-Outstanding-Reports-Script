import traceback

try:
    from Reports.MainMenu import MainMenu

    print("Main menu imported successfully.")

    app = MainMenu()

    print("Main menu created successfully.")

    app.mainloop()

except Exception:
    print("\nERROR STARTING APPLICATION:\n")
    traceback.print_exc()
    input("\nPress Enter to close...")
