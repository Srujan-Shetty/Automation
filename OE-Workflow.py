# 10. Order Entry Workflow
# - Automate a flow in your app (Normal Order Entry)
# - Validate error popups
# - Log messages

from pywinauto import Application
from pywinauto.keyboard import send_keys
import os
import time
import pyautogui

path = r"C:\UFT\AUG\1.0.8_01AUG_Portable\NT_Release_1.0.8_01AUG_Portable\Release"
os.chdir(path)
exe_name = "NorenTrader.exe"
#NorenTrader ExePath
exe_path = os.path.join(path, exe_name)

#Load the Exe Path 
app=Application(backend='uia').start(exe_path)
dlg = app.window(title_re='.*Login.*')
dlg.wait('visible', timeout=10)

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


#Connect to LoginPrefpage and Click on Continue Button
dlg.child_window(title="LoginPrefPage", auto_id="LoginPrefPage", control_type="Tab")
dlg.child_window(title="Continue...(8)", control_type="Button").click_input()
time.sleep(5)

#Connect to NorenTrader Main Window
main_window = app.window(title_re=".*Noren Trader.*")
main_window.wait('visible', timeout=20)
main_window.wait('ready', timeout=20)
time.sleep(1)

#Login Successful Pop UP Appears and Click on OK
Welcome=main_window.child_window(title="Login Successful", auto_id="popupWin", control_type="Window")
Login_success=Welcome.child_window(title="OK", auto_id="btnOK", control_type="Button")
Login_success.click_input()
time.sleep(1)

#Set Focus to Marketwtach on Login
main_window.child_window(title="NewGroup", auto_id="W_NewGroupTabId", control_type="TabItem").set_focus()
time.sleep(2)

#Connect to Exchange Segment Field and Select Exchange
Exchange=main_window.child_window(auto_id="ExchangeCombo", control_type="ComboBox")
Exchange.click_input()
Exchange.type_keys('BSE{ENTER}', with_spaces=True)
time.sleep(2)

#Connect to Instrument Field and Select Instrument
Instruments=main_window.child_window(auto_id="InstrCombo", control_type="ComboBox")
Instruments.click_input()
Instruments.type_keys('B{ENTER}',with_spaces=True)
time.sleep(2)

#Connect to Symbol Searchbox field and Enter Scrip Name
Symbol_Search=main_window.child_window(auto_id="SymbolSearchBox", control_type="Custom")
Symbol_name=Symbol_Search.child_window(auto_id="PART_Editor", control_type="Edit")
Symbol_name.set_text('YATRA')
time.sleep(1)
send_keys('{ENTER}')# Scrip Added to MW
time.sleep(1)

#Select Scrip from Marketwatch Group
Scrip_Add=main_window.child_window(title="UFTrader.UFT_Stores.MarketWatchItem", control_type="DataItem")
Scrip_select=Scrip_Add.child_window(title="Row 0, Column 1: YATRA", auto_id="SymbolName", control_type="Custom")#Scrip name 
Scrip_select.click_input()
Scrip_select.set_focus()
send_keys('{F1}')#Invoke Buy Order Entry
time.sleep(1)

#Connect to Order Entry Window and Click on Buy Button
Expected_Result="Please Select Account Id."
Buy_OE=main_window.child_window(title="Buy Order Entry", auto_id="OrderEntrywin", control_type="Window")
Buy_OE.child_window(title="Buy", auto_id="OrderPlaceBtn", control_type="Button").click_input()
time.sleep(1)

#Connect Error Message Pop up window
Error=main_window.child_window(title="Error", auto_id="popupWin", control_type="Window")

#Extract text From PopUP message
Error_Msg=Error.child_window(title="Please Select Account Id.  ", auto_id="textMsg", control_type="Text").window_text()
Error.close()
Buy_OE.close()

#Error Pop Validation with Expected Result
if Expected_Result in Error_Msg:
    print("Test Passed")
    print("Expected_Result:",Expected_Result)
    print("Actual Result:",Error_Msg)
else:
    print("Test Failed")
    print("Expected_Result:",Expected_Result)
    print("Actual Result:",Error_Msg)




Logs=main_window.child_window(title="Logs", auto_id="FloatingPaneWindow(42883409)", control_type="Window")

Logs.print_control_identifiers()