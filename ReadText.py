from pywinauto import Application
import time

# Start Notepad
app = Application().start('notepad.exe')
dlg = app.window(title_re=".*Notepad")

# Get the Edit control
edit = dlg.Edit

# Type text
text_to_type = "Hello Automation"
edit.type_keys(text_to_type, with_spaces=True)
time.sleep(1)

# Read text back from the Edit control
read_text = edit.window_text()

# Verify
if read_text == text_to_type:
    print("Text verified successfully!")
else:
    print("Text verification failed!")
    print(f"Expected: {text_to_type}")
    print(f"Got: {read_text}")

# Close Notepad without saving

app.kill()
