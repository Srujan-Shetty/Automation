from UI_Core_Function_TestData import LoginWindow_TestData
from pywinauto import Application
from pywinauto import keyboard
import os
import time

LoginTestData = LoginWindow_TestData()


class General_Prerequisite():                                     #Class General_Prerequisite has methods to Login exe ,and Download exchanges
    #------------------Prerequisite1: Invoke NorenTrader.exe from Specified Path----------------------#

    def Invoke_Exe(self):                                         #Method to Invoke exe from specified Location and focus on Login Page
        self.path = LoginTestData.Get_LoginExe_path()             # Get EXE Path Module
        os.chdir(os.path.dirname(self.path))
        Application().start(self.path)
        time.sleep(2)
        self.app = Application(backend='uia').connect(path=self.path)
        self.LoginWnd = self.app.window(title ="Login " ,auto_id="loginWindow",class_name="Window" )
        self.LoginWnd.wait('visible', timeout=20)
        self.LoginWnd.set_focus()

    # -----------------------------Prerequisite2: Login exe -------------------------------------#
    def Login_Exe(self):                                            # Method to Login exe by input USERNAME and PASSWORD
        self.UsernameEditBox = self.LoginWnd.child_window(auto_id="PART_Editor", class_name="TextBox",framework_id="WPF")  # Locate the username input textbox
        self.UsernameEditBox.wait('enabled', timeout=20)            # Wait until the username box is enabled (max 20 seconds)
        self.UsernameEditBox.set_focus()                            # Set focus on the username textbox
        USERNAME = LoginTestData.Get_Login_UserName()               # Get the username from the test data
        self.UsernameEditBox.set_edit_text(str(USERNAME))           # Enter the username into the textbox
        print("[INFO] Enter USERNMAE")
        self.PasswordEditBox = self.LoginWnd.child_window(auto_id="PART_Editor", class_name="PasswordBox",framework_id="WPF")  # Locate the password input box
        self.PasswordEditBox.set_focus()                            # Set focus on the password box
        PASSWORD = LoginTestData.Get_Login_Password()               # Get the password from the test data
        self.PasswordEditBox.set_edit_text(str(PASSWORD))           # Enter the password into the password box
        print("[INFO] Enter PASSWORD")
        self.LoginButton = self.LoginWnd.child_window(title="Log In", auto_id="LoginBtn",class_name="Button").click_input()  # Locate and click the "Log In" button
        print("[INFO] Click Login Button")
        Information = self.app.window(title_re=".*Information.*", control_type="Window")
        if Information.exists(timeout=1):
            Information.child_window(title="OK", control_type="Button").click_input()
            print("[INFO] Click On Password Expiry Warning PopUp")
        else:
            print("[INFO] No Password Expiry Warning PopUp")

    # -----------------------------Prerequisite3: Download Exch-------------------------------------#
    def Login_Exch_Download(self):  # Method to Download Exchange
        time.sleep(3)
        keyboard.send_keys("{ENTER}")
        keyboard.send_keys("{ENTER}")
        self.Check_Exe_LoggedIn()

    # -----------------------------Check whether user logged in successfully?-------------------------------------#
    def Check_Exe_LoggedIn(self):
        self.path = LoginTestData.Get_LoginExe_path()  # EXE Path
        self.app = Application(backend='uia').connect(path=self.path)
        self.App_MainWnd = self.app.window(title="Noren Trader", framework_id="WPF")
        if self.App_MainWnd.exists(timeout=100):
            self.LoginSucessPopup = self.App_MainWnd.window(title="Login Successful", auto_id="popupWin",class_name="Window")
            print("Noren Trader is open ")
        else:
            print("Noren Trader.exe does not exist")
        self.LoginSucessPopup.set_focus()
        self.ClickOK = self.LoginSucessPopup.child_window(title="OK", auto_id="btnOK",class_name="Button").click_input()

    # -----------------------------Prerequisite4: Set focus on Noren Trader exe-------------------------------------#
    def Focus_On_Exe(self):                                      #Method to Keep Focus on Logged in exe
        self.path = LoginTestData.Get_LoginExe_path()            # EXE Path
        self.app = Application(backend='uia').connect(path=self.path)
        self.App_MainWnd = self.app.window(title="Noren Trader", framework_id="WPF")
        self.App_MainWnd.set_focus()
   # -----------------------------Prerequisite5: Invoke Market watch -------------------------------------#
    def Invoke_MarketWatch(self):                                                                              # Method to Invoke Market Watch
        keyboard.send_keys('{F4}')                                                                             # Invoke Watch Grp Manager using shortcut F4
        keyboard.send_keys('{F4}')                                                                             # Invoke Watch Grp Manager using shortcut F4
        self.GroupManagerWnd = self.App_MainWnd.child_window(title="Watch Group Manager",class_name="Window")  # Get the Watch Group Manager window under the main application window
        self.SelectGrp = self.GroupManagerWnd.child_window(title="AutomationWatchGrp",framework_id="WPF").double_click_input()  # Double-click on the group named "AutomationWatchGrp" to select it
        self.GroupManagerWnd = self.App_MainWnd.child_window(title="Watch Group Manager",class_name="Window").close()  # Close the Watch Group Manager window after selecting the group
        self.AutomationWatchGrp = self.App_MainWnd.child_window(title="AutomationWatchGrp",auto_id="W_AutomationWatchGrp",class_name="MarketWatch")
        self.Watch_datagrid = self.AutomationWatchGrp.child_window(auto_id="uiGrid")
        self.Select_Scrip = self.Watch_datagrid.child_window(title="UFTrader.UFT_Stores.MarketWatchItem")
        if self.Select_Scrip.exists(timeout=1):
            self.Select_Scrip.click_input()
            print("[INFO] Selected scrip from Market Watch")
        else:
            print("[INFO] No Scrips in Market Watch")




#-------------------------------------------------Prerequisite---------------------------------------------------------#