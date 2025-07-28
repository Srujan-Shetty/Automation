from pywinauto import Application
import time

# Start Notepad
app = Application().start('notepad.exe')
time.sleep(1)

# Get the Notepad window
dlg = app.window(title_re=".*Notepad")

# Click Help > About Notepad
dlg.menu_select("Help->About Notepad")
time.sleep(2)  # Give time for the About dialog to appear

# Get the About Notepad dialog
about_dlg = app.window(title="About Notepad")

# Close the About dialog
about_dlg.Close.click()
time.sleep(1)

# Close Notepad
dlg.close()
