# 7. File Dialog Handling
# - Open File > Open in Notepad
# - Execute all the task mentioned in Week 2
# - Click Cancel

from pywinauto import Application
import time
import pyperclip
# Case 1
# 1. Launch and Close Application
# - Launch Notepad
# - Wait 5 seconds

app = Application(backend='uia').start('notepad.exe "C:\\MYNOTE\\Note.txt"')
time.sleep(5)
dlg = app.window(title_re=".*Notepad")
# Get the Edit control
edit = dlg.Edit

# Case 2. Enter and Read Text in Notepad
# - Type "Hello Automation"
# - Read back and verify the text

text_to_type = "Hello Automation"
edit.type_keys(text_to_type, with_spaces=True)
time.sleep(1)

# Read text back from the Edit control
read_text = edit.get_value()

# Verify
if read_text == text_to_type:
    print("Case2--->Text verified successfully!")
    print(f"Expected: {text_to_type}")
    print(f"Got: {read_text}")
    
else:
    print("Text verification failed!")
    print(f"Expected: {text_to_type}")
    print(f"Got: {read_text}")

# 3. Click Menu Items
# - Open Notepad
# - Navigate Help > About Notepad
# - Close dialog
time.sleep(2)
dlg.menu_select("Help->About Notepad")
time.sleep(2)  

# Get the About Notepad dialog
about_dlg = dlg.child_window(title="About Notepad", control_type="Window")
about_dlg.close()

# 4. Copy and Paste Text
# - Use shortcuts to select all, copy, and paste text
# - Use pyperclip to verify copied text

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



