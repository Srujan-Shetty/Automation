# 6. Error Popup Handling
# - Trigger a message box
# - Capture and validate the error message

from pywinauto.application import Application
import time

# Connect to the Notepad window
app = Application(backend="uia").start('notepad.exe')
dlg = app.window(title_re=".*Notepad")
time.sleep(2)

# Type the text into the Notepad
dlg.Edit.type_keys("AUTOMATION POP-UP", with_spaces=True, pause=0.03)
time.sleep(1)

# Click on the Edit menu by 
menu=dlg.child_window(title="Edit", control_type="MenuItem")
menu.click_input()
time.sleep(1)

# Click on the Replace... menu item
Replace=dlg.child_window(title="Replace...	Ctrl+H", auto_id="23", control_type="MenuItem")
Replace.click_input()
time.sleep(1)

# Type the text to find and replace
Replace_what = dlg.child_window(title="Find what:", auto_id="1152", control_type="Edit")
Replace_what.set_edit_text("POP-UP")
time.sleep(1)

# replace the text with another string
Replace_with = dlg.child_window(title="Replace with:", auto_id="1153", control_type="Edit")
Replace_with.set_edit_text("POPUP")
time.sleep(1)

#click on the Replace All button
Replace=dlg.child_window(title="Replace", auto_id="1024", control_type="Button")
Replace.click_input()

 #capture the pop-up message
popup = app.window(title_re=".*Notepad")
message =popup.child_window(title="Cannot find \"POP-UP\"", auto_id="65535", control_type="Text").window_text()
print("Pop-up message:", message)
Expected="Cannot find \"POP-UP\""
if Expected in message:
    print("Test Passed")
else:
    print("Test Failed")

popup.child_window(title="OK", control_type="Button").click_input()
time.sleep(1)
Replace.close()  # Close the Replace dialog
time.sleep(1)
app.kill()  # Close Notepad












