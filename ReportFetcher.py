import webbrowser
import time

reports = [
    "https://supermicrocomputer.lightning.force.com/lightning/r/Report/00OPm00000KI2sfMAD/view?queryScope=userFolders",
    "https://supermicrocomputer.lightning.force.com/lightning/r/Report/00OPm00000KI3APMA1/view?queryScope=userFolders",
    "https://supermicrocomputer.lightning.force.com/lightning/r/Report/00OPm00000KIDGDMA5/view?queryScope=userFolders",
]

# Open the first report
webbrowser.open(reports[0])

# Give the browser a moment to open the window
time.sleep(1)

# Open the remaining reports as tabs
for url in reports[1:]:
    webbrowser.open_new_tab(url)