from pywinauto import Application
import time
 
app = Application().start("Notepad.exe")
time.sleep(5) # Add a delay if needed for the new process to appear

app.kill()
 