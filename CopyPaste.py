from pywinauto import Application
import time
import pyperclip

app = Application().start("notepad.exe")
time.sleep(2)

notepad = app.window(title_re=".*Notepad")
edit = notepad.Edit

# Type "Hello Automation"
edit.type_keys("Hello Automation", with_spaces=True)

# Select all and copy (Ctrl+A Ctrl+C)
edit.type_keys("^a^c")
time.sleep(0.5)

# Move caret to end, insert newline
edit.type_keys("{END}{ENTER}")

# Paste copied text back into Notepad (Ctrl+V)
edit.type_keys("^v")
time.sleep(0.5)

# Now Notepad has original text + pasted copy

# Close Notepad
app.kill()
