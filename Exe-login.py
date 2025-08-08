# 9. Norenexe Login Window
# - Automate a login screen
# - Validate success/failure messages

from pywinauto import Application
from pywinauto.keyboard import send_keys
import os
import time

path = r"C:\UFT\AUG\1.0.8_01AUG_Portable\NT_Release_1.0.8_01AUG_Portable\Release"
os.chdir(path)
exe_name = "NorenTrader.exe"
#NorenTraded Exe Path
exe_path = os.path.join(path, exe_name)

#Load the Exe Path 
app=Application(backend='uia').start(exe_path)
dlg = app.window(title_re='.*Login.*')
dlg.wait('visible', timeout=10)

#Enter Value for Username Field
Username_field=dlg.child_window(auto_id="LoginTxt", control_type="Edit")
Username_field.wait('ready', timeout=10)
Username_field.type_keys('^a{BACKSPACE}')
Username_field.type_keys("SRUJAN", pause=0.1)
time.sleep(2)

#Enter Value For Password Field
Password_field=dlg.child_window(auto_id="PasswordTxt", control_type="Edit")
Password_field.wait('ready', timeout=10)
Password_field.type_keys('^a{BACKSPACE}')  # Ctrl+A + Backspace to clear the field
Password_field.type_keys("Srujan@1234", pause=0.1 )
time.sleep(2)

#Click On Login Button
Login=dlg.child_window(title="Log In", auto_id="LoginBtn", control_type="Button")
Login.click_input()
time.sleep(1)

#Information Pop UP appears click on OK
Information = app.window(title_re=".*Information.*", control_type="Window")
Information.child_window(title="OK", control_type="Button").click_input()
time.sleep(2)

#Connect to LoginPrefpage and Click on Continue Button
dlg.child_window(title="LoginPrefPage", auto_id="LoginPrefPage", control_type="Tab")
dlg.child_window(title="Continue...(9)", control_type="Button").click_input()
time.sleep(5)


#Connect to NorenTrader Main Window
main_window = app.window(title_re=".*Noren Trader.*")
main_window.wait('visible', timeout=20)
main_window.wait('ready', timeout=20)

#Login Successful Pop UP Appears and Click on OK
Welcome=main_window.child_window(title="Login Successful", auto_id="popupWin", control_type="Window")
# Extract text From PopUP message
Message=Welcome.window_text()
print(Message)
Login_success=Welcome.child_window(title="OK", auto_id="btnOK", control_type="Button")
Login_success.click_input()

#Validate Login Success Message with Expected Message
Expected="Login Successful"
if Message in Expected:
    print("Logged In Successfully:")
    print("Expected Result:",Expected)
    print("OutPut:",Message)
else:
    print("Login Error")

Logs=main_window.child_window(title="Logs", auto_id="Logs", control_type="Group")
Logs_panel=Logs.child_window(title="DataPanel", auto_id="dataPresenter", control_type="Pane")
Logs_panel.set_focus()
send_keys('^f')

Search=main_window.child_window(title="SearchPanel", auto_id="SearchPanel", control_type="Pane")
Search_box=Search.child_window(auto_id="SearchComboBox", control_type="ComboBox")
Search_box.set_focus()
Search_box.type_keys("Login Successful",with_spaces=True)
send_keys('{ENTER}')

Extract=Search.child_window(title="Login Successful", auto_id="SearchComboBox", control_type="ComboBox")
Msg=Extract.window_text()
print(Msg)

if Msg in Expected:
    print("Log Msg Verified")
    print("Expected:",Expected)
    print("Actual:",Msg)
else:
    print("Log Msg Not Verified")
    print("Expected:",Expected)
    print("Actual:",Msg)

# Search.print_control_identifiers()


