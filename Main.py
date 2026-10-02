import traceback

try:
    from GUI import CoreWeaveTool

    print("GUI imported successfully.")

    app = CoreWeaveTool()

    print("Application created successfully.")

    app.mainloop()

except Exception:
    print("\nERROR STARTING APPLICATION:\n")
    traceback.print_exc()
    input("\nPress Enter to close...")