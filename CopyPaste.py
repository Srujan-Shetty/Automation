# 4. Copy and Paste Text
# - Use shortcuts to select all, copy, and paste text
# - Use pyperclip to verify copied text

from pywinauto import Application
import time
import pyperclip

app = Application().start("notepad.exe")
time.sleep(2)

notepad = app.window(title_re=".*Notepad")
edit = notepad.Edit

# Type "Hello Automation"
text_to_type= "Hello Automation"
edit.type_keys(text_to_type, with_spaces=True)

# Select all and copy (Ctrl+A Ctrl+C)
edit.type_keys("^a^c")
time.sleep(0.5)

Pasted_text= pyperclip.paste()
if text_to_type== Pasted_text:
    print("Case 4--->Copied text matches expected text!")
else:
    print("Copied text does NOT match") 
    print("Expected:", text_to_type)
    print("Got:", Pasted_text)

app.kill()  # Close Notepad