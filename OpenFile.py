# 1. Launch and Close Application
# - Launch Notepad
# - Wait 5 seconds
# - Close the application

from pywinauto import Application
import time

app=Application().start('notepad.exe')
time.sleep(5)
Notepad = app.window(title_re='.*Notepad')
# Notepad.print_control_identifiers()

app.kill()

