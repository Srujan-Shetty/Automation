from pywinauto import Application
import time
import os

app=Application(backend='uia').start("notepad.exe")
time.sleep(1)
dlg = app.window(title_re='.*Notepad')

dlg.child_window(title="System", control_type="MenuItem")
dlg.child_window(title="File", control_type="MenuItem").click_input()
dlg.child_window(title="Open...	Ctrl+O", auto_id="2", control_type="MenuItem").click_input()

open_dlg = app.window(title="Open",control_type="Window")
file_name_combo=open_dlg.child_window(title="File name:", auto_id="1148", control_type="ComboBox")
file_name_edit = file_name_combo.child_window(control_type="Edit")
file_name_edit.set_edit_text("Note.txt")



# dlg.print_control_identifiers()