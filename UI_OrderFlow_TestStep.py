from pywinauto import Application
from pywinauto import keyboard
from pywinauto import Desktop
from UI_OrderFlow_TestData import LoginWindow_TestData           #Import TestData's required for testcase from UI_OrderFlow_TestData Module
from UI_OrderFlow_OrderAction import OrderPlacements
import time
import os
from UI_Order_Flow_helper import Get_File_Path

LoginTestData = LoginWindow_TestData()
OP = OrderPlacements()

class General_Prerequisite():                                     #Class General_Prerequisite has methods to Login exe ,and Download exchanges
    # -------------------------------------------------Prerequisite---------------------------------------------------------#
    def Invoke_PuttyExe(self):

        app = Application().start(r"C:\Program Files\PuTTY\putty.exe ")
        desktop = Desktop(backend='uia')
        pt = app.PuTTY
        PuttyConfig_Window = desktop.window(title="PuTTY Configuration", class_name="PuTTYConfigBox")
        PuttyConfig_Window.child_window(title="Matsya").double_click_input()  # Select Server
        desktop.window(title="matsya.kambala.co.in - PuTTY", class_name="PuTTY")  # Login Server
        pt.wait('ready')
        time.sleep(10)
        print("Open Putty.exe")
        row = 2
        pt.type_keys("ssh{ }rama")
        pt.type_keys("{ENTER}")
        time.sleep(2)
        for row in range(row , 12):
            TestInfo_sheet = Get_File_Path("TestCaseList.xlsx", 0)  # Method to Active the TestCaseList.xlsx sheet
            Exch = TestInfo_sheet.cell(row=5, column=2).value
            TOKEN = TestInfo_sheet.cell(row=row, column=2).value
            pt.type_keys("MFEED{ }BrdBK{ }EXCH_SEG:{ }" + str(Exch) + "{ }TOKEN:{ }" + str(TOKEN) + "{ }_marketfeed{ }{{}{ }UPPER_CKT:{ }10000000{ }LOWER_CKT:{ }100{ }{}}")
            pt.type_keys("{ENTER}")
            time.sleep(1)
            pt.type_keys("MFEED{ }BrdBK{ }EXCH_SEG:{ }" + str(Exch) + "{ }TOKEN:{ }" + str(TOKEN) + "{ }_marketfeed{ }{{}{ }REMARKS:{ }Scrip Used for Automation Testing{}}")
            pt.type_keys("{ENTER}")
            time.sleep(1)
            #pt.type_keys("HI")
            pt.type_keys("RFEED{ }BrdBK{ }EXCH_SEG:{ }" + str(Exch) + "{ }TOKEN:{ }" + str(TOKEN) + "{ }_marketfeed{ }{{}{ }QTY_FREEZE:{ }10000{}}")
            pt.type_keys("{ENTER}")
            time.sleep(1)
            print("Circuit Limit and Remark Publish is done ")
        #row1 = 12
        for row in range(12, 16):  # For Loop to publish RMS limits
            TestInfo_sheet = Get_File_Path("TestCaseList.xlsx", 0)  # Method to Active the TestCaseList.xlsx sheet
            ENTITY_ID = TestInfo_sheet.cell(row=row, column=2).value
            pt.type_keys("cozy_command{ }ReqBK{ }ELSE{ }LT{ }ENTITY_ID:{ }" + str(ENTITY_ID) + "{ }CASH:{ }2388888888500{ }TURNOVER_LIMIT:{ }1000000000000{ }PENDING_ORDER_LIMIT:{ }1000000000000")
            pt.type_keys("{ENTER}")
            time.sleep(1)
            print("Red Limit Publish is done for Entity_Id:"+str(ENTITY_ID))

        Putty_Wndow = desktop.window(class_name="PuTTY", framework_id="Win32")
        Putty_Wndow.set_focus()
        Close_PuttyWnd = Putty_Wndow.child_window(title="Close").click_input()
        keyboard.send_keys('{ENTER}')
        print("Putty.exe is closed")




    def Invoke_Exe(self):                                         #Method to Invoke exe from specified Location and focus on Login Page
        self.path = LoginTestData.Get_LoginExe_path()             # Get EXE Path Module
        #from pywinauto.application import Application
        os.chdir(os.path.dirname(self.path))
        Application().start(self.path)
        time.sleep(2)
        self.app = Application(backend='uia').connect(path=self.path)
        self.LoginWnd = self.app.window(title ="Login " ,auto_id="loginWindow",class_name="Window" )
        self.LoginWnd.wait('visible', timeout=20)
        self.LoginWnd.set_focus()

    def Login_Exe(self):  # Method to Login exe by input USERNAME and PASSWORD
        self.UsernameEditBox = self.LoginWnd.child_window(auto_id="PART_Editor", class_name="TextBox",framework_id="WPF")
        self.UsernameEditBox.wait('enabled', timeout=20)
        self.UsernameEditBox.set_focus()
        USERNAME = LoginTestData.Get_Login_UserName()
        self.UsernameEditBox.set_edit_text(str(USERNAME))
        self.PasswordEditBox = self.LoginWnd.child_window(auto_id="PART_Editor", class_name="PasswordBox",framework_id="WPF")
        self.PasswordEditBox.set_focus()
        PASSWORD = LoginTestData.Get_Login_Password()
        self.PasswordEditBox.set_edit_text(str(PASSWORD))
        self.LoginButton = self.LoginWnd.child_window(title="Log In", auto_id="LoginBtn",class_name="Button").click_input()
        self.desktop = Desktop(backend='uia')
        #self.ClickInfo = self.desktop.window(title="Information", auto_id="popupWin", class_name="Window")
        #self.ClickInfo.set_focus()
        #keyboard.send_keys('{ENTER}')

    def Login_Exch_Download(self):                                   #Method to Download Exchange
        self.ExchDownloadWnd = self.LoginWnd.child_window(title="LoginPrefPage", auto_id="LoginPrefPage",class_name="LoginPrefPage")
        # self.ExchDownloadWnd.print_control_identifiers()
        self.ClickHeader = self.ExchDownloadWnd.child_window(title="Exchanges", auto_id="PART_Content" , control_type="Text")
        self.ClickHeader.click_input()
        self.ClickHeader.click_input(button='right')
        self.SelectAllBtn = self.LoginWnd.child_window(title="Select All", class_name="MenuItem",framework_id="WPF").click_input()
        Select_Button=self.LoginWnd.child_window(title="LoginPrefPage", auto_id="LoginPrefPage", class_name="LoginPrefPage", control_type="Tab")
        Select_Button.child_window(title="Continue...(8)", control_type="Button").click_input()
        time.sleep(10)
        self.Check_Exe_LoggedIn()

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

    def Focus_On_Exe(self):                                      #Method to Keep Focus on Logged in exe
        self.path = LoginTestData.Get_LoginExe_path()            # EXE Path
        self.app = Application(backend='uia').connect(path=self.path)
        self.App_MainWnd = self.app.window(title="Noren Trader", framework_id="WPF")
        self.App_MainWnd.set_focus()

    def Invoke_MarketWatch(self):                                 #Method to Invoke Market Watch
        keyboard.send_keys('{F4}')
        self.GroupManagerWnd = self.App_MainWnd.child_window(title="Watch Group Manager", class_name="Window")
        self.SelectGrp = self.GroupManagerWnd.child_window(title="AutomationWatchGrp",framework_id="WPF").double_click_input()  # Select Group Name : AutomationWatchGrp
        self.GroupManagerWnd = self.App_MainWnd.child_window(title="Watch Group Manager", class_name="Window").close()

    def Focus_On_MW(self):                                          #Method to Keep focus on Market Watch Panel
        self.MWGroup = self.App_MainWnd.child_window(title="AutomationWatchGrp", class_name="MarketWatch",auto_id="W_AutomationWatchGrp")
        self.MW_Panel = self.MWGroup.child_window(title="DataPanel",auto_id="dataPresenter").click_input()


#-------------------------------------------------Prerequisite---------------------------------------------------------#
#------------------------------------------Complete Order Flow TestCase------------------------------------------------#
class GeneralTestStep(General_Prerequisite):
    def LMT_Order_TestCases(self,Result_path):
        self.Focus_On_Exe()
        # self.Invoke_MarketWatch()   #Focus On Logged in Exe
        self.Focus_On_MW()                                                     #Keep Focus in Market Watch
        TestInfo_sheet = Get_File_Path("TestData.xlsx" , 1) # Method to Active the TestData.xlsx sheet
        row_value = 2                                                          # Since Testcase starts from 2nd row in NorenTraderShortCutTestcases.xlsx
        column_value = 1                                                       #Colum start with 1
        TC_NO = 0                                                              #Initialize 0 as TestCase Number
        SUB_TC_NO = 0
        Remark = 0
        for row in range(row_value, 105):
            for column in range(column_value, 105):
                Test_Scenario = TestInfo_sheet.cell(row=row_value, column=1).value  # Get TestScenario from Excel
                if Test_Scenario == None :                                           #Test Scenario Column is None ,increment sub TestCase Number and  do not increment TestCase Number
                    print("DONT")
                else:
                    TC_NO = TC_NO+1                                                   #TestScenario present in column then Increment TestCase Number and initialize
                    #SUB_TC_NO = SUB_TC_NO+1
                    print("üpdated Numbr: "+str(TC_NO))

                Test_Case = TestInfo_sheet.cell(row=row_value, column=2).value         # Get Test Case from Excel
                print("TestCaseNumber_" + str(TC_NO) + ": " + "TestScenario is: "+str(Test_Scenario)+"TestCase: "+str(Test_Case))
                Trans_Type = TestInfo_sheet.cell(row=row_value ,column=3).value          #Get TransType from Excel
                Exchange = TestInfo_sheet.cell(row=row_value ,column=4).value            #Get Exchange from Excel
                Order_Type = TestInfo_sheet.cell(row=row_value ,column=5).value          #Get Order Type from Excel
                Order_Duration = TestInfo_sheet.cell(row=row_value ,column=6).value      #Get Order Duration from Excel
                Trading_Symbol = TestInfo_sheet.cell(row=row_value ,column=7).value      #Get Trading Symbol from Excel
                Input_Quantity = TestInfo_sheet.cell(row=row_value, column=8).value      # Get Input Quantity from from Excel
                Input_Price = TestInfo_sheet.cell(row=row_value, column=9).value         # Get Input Price value from Excel
                Input_TriggerPrice = TestInfo_sheet.cell(row=row_value, column=10).value  # Get Input Trigger Price value from Excel
                Input_DiscQty = TestInfo_sheet.cell(row=row_value,column=11).value        # Get Input DiscQty value from Excel
                Quantity = TestInfo_sheet.cell(row=row_value ,column=12).value            #Get Quantity from from Excel
                Price = TestInfo_sheet.cell(row=row_value ,column=13).value               #Get Price value from Excel
                TriggerPrice = TestInfo_sheet.cell(row=row_value, column=14).value        # Get Trigger Price value from Excel
                DiscQty = TestInfo_sheet.cell(row=row_value, column=15).value            # Get Disc Qty value from Excel
                OpenQty = TestInfo_sheet.cell(row=row_value ,column=16).value            #Get Open Qty value from Excel
                TradedQty = TestInfo_sheet.cell(row=row_value ,column=17).value          #Get Traded Qty value from Excel
                AvgPrice = TestInfo_sheet.cell(row=row_value ,column=18).value           #Get Average Price value from Excel
                Product = TestInfo_sheet.cell(row=row_value ,column=19).value            #Get Product from Excel
                Account_ID = TestInfo_sheet.cell(row=row_value ,column=20).value         #Get Account_Id from Excel
                TypeOfOrder = TestInfo_sheet.cell(row=row_value, column=21).value         # Get Type Of Order from Excel
                ID =OP.PlaceLMTOrder(Result_path,TC_NO,row_value,Test_Scenario,Test_Case,Trans_Type,Exchange,Order_Type,Order_Duration,Trading_Symbol,Input_Quantity,Input_Price,Input_TriggerPrice,Input_DiscQty,Quantity,Price,TriggerPrice,DiscQty,Product,Account_ID,Remark,OpenQty,TradedQty,AvgPrice,TypeOfOrder)     # send Order values to place / Modify /Cancel Order
                Remark = str(ID)
                row_value = row_value + 1                                                 # Increment Index value of Row
                column_value = column_value + 1                                           # Increment Index value of column
                #print(str(TC_NO))


    def SLLMT_Order_TestCases(self,Result_path):
        self.Focus_On_Exe()                                               # Focus On Logged in Exe
        self.Invoke_MarketWatch()
        self.Focus_On_MW()                                                     #Keep Focus in Market Watch
        TestInfo_sheet = Get_File_Path("TestData.xlsx", 2)  # Method to Active the TestData.xlsx sheet
        row_value = 2                                                      # Since Testcase starts from 2nd row in NorenTraderShortCutTestcases.xlsx
        column_value = 1                                                       #Colum start with 1
        TC_NO = 0                                                             #Initialize 0 as TestCase Number
        SUB_TC_NO = 0
        Remark = 0
        for row in range(row_value, 108):
            for column in range(column_value, 108):
                Test_Scenario = TestInfo_sheet.cell(row=row_value, column=1).value  # Get TestScenario from Excel
                if Test_Scenario == None :                                           #Test Scenario Column is None ,increment sub TestCase Number and  do not increment TestCase Number
                    print("DONT")
                else:
                    TC_NO = TC_NO+1                                                   #TestScenario present in column then Increment TestCase Number and initialize
                    #SUB_TC_NO = SUB_TC_NO+1
                    print("üpdated Numbr: "+str(TC_NO))

                Test_Case = TestInfo_sheet.cell(row=row_value, column=2).value         # Get Test Case from Excel
                print("TestCaseNumber_" + str(TC_NO) + ": " + "TestScenario is: "+str(Test_Scenario)+"TestCase: "+str(Test_Case))
                Trans_Type = TestInfo_sheet.cell(row=row_value ,column=3).value          #Get TransType from Excel
                Exchange = TestInfo_sheet.cell(row=row_value ,column=4).value            #Get Exchange from Excel
                Input_OrderType = TestInfo_sheet.cell(row=row_value ,column=5).value          #Get Order Type from Excel
                Order_Duration = TestInfo_sheet.cell(row=row_value ,column=6).value      #Get Order Duration from Excel
                Trading_Symbol = TestInfo_sheet.cell(row=row_value ,column=7).value      #Get Trading Symbol from Excel
                Input_Quantity = TestInfo_sheet.cell(row=row_value, column=8).value      # Get Input Quantity from from Excel
                Input_Price = TestInfo_sheet.cell(row=row_value, column=9).value         # Get Input Price value from Excel
                Input_TriggerPrice = TestInfo_sheet.cell(row=row_value ,column=10).value #Get Input Trigger Price value from Excel
                Input_DiscQty = TestInfo_sheet.cell(row=row_value ,column=11).value      # Get Input DiscQty value from Excel
                Quantity = TestInfo_sheet.cell(row=row_value ,column=12).value           #Get Quantity from from Excel
                Price = TestInfo_sheet.cell(row=row_value ,column=13).value              #Get Price value from Excel
                TriggerPrice = TestInfo_sheet.cell(row=row_value ,column=14).value       #Get Trigger Price value from Excel
                DiscQty = TestInfo_sheet.cell(row=row_value, column=15).value            # Get Disc Qty value from Excel
                OpenQty = TestInfo_sheet.cell(row=row_value ,column=16).value            #Get Open Qty value from Excel
                TradedQty = TestInfo_sheet.cell(row=row_value ,column=17).value          #Get Traded Qty value from Excel
                AvgPrice = TestInfo_sheet.cell(row=row_value ,column=18).value           #Get Average Price value from Excel
                OrderType = TestInfo_sheet.cell(row=row_value ,column=19).value           #Get OrderType value from Excel
                Product = TestInfo_sheet.cell(row=row_value ,column=20).value            #Get Product from Excel
                Account_ID = TestInfo_sheet.cell(row=row_value ,column=21).value         #Get Account_Id from Excel
                TypeOfOrder = TestInfo_sheet.cell(row=row_value ,column=22).value         #Get Type Of from Excel
                ID =OP.PlaceSLOrder(Result_path,TC_NO,row_value,Test_Scenario,Test_Case,Trans_Type,Exchange,Input_OrderType,Order_Duration,Trading_Symbol,Input_Quantity,Input_Price,Input_TriggerPrice,Input_DiscQty,OrderType,Quantity,Price,TriggerPrice,DiscQty,Product,Account_ID,Remark,OpenQty,TradedQty,AvgPrice,TypeOfOrder)     # send Order values to place / Modify /Cancel Order
                Remark = str(ID)
                row_value = row_value + 1                                                 # Increment Index value of Row
                column_value = column_value + 1                                           # Increment Index value of column
                #print(str(TC_NO))


    def MKT_Order_TestCases(self,Result_path):
        self.Focus_On_Exe()  # Focus On Logged in Exe
        self.Invoke_MarketWatch()
        self.Focus_On_MW()  # Keep Focus in Market Watch
        TestInfo_sheet = Get_File_Path("TestData.xlsx", 3)  # Method to Active the TestData.xlsx sheet
        row_value = 44                         # Since Testcase starts from 2nd row in NorenTraderShortCutTestcases.xlsx
        column_value = 1                      # Colum start with 1
        TC_NO = 0                             # Initialize 0 as TestCase Number
        SUB_TC_NO = 0
        Remark = 0
        for row in range(row_value, 150):
            for column in range(column_value, 150):
                Test_Scenario = TestInfo_sheet.cell(row=row_value, column=1).value  # Get TestScenario from Excel
                if Test_Scenario == None:  # Test Scenario Column is None ,increment sub TestCase Number and  do not increment TestCase Number
                    print("DONT")
                else:
                    TC_NO = TC_NO + 1  # TestScenario present in column then Increment TestCase Number and initialize
                    # SUB_TC_NO = SUB_TC_NO+1
                    print("üpdated Numbr: " + str(TC_NO))

                Test_Case = TestInfo_sheet.cell(row=row_value, column=2).value  # Get Test Case from Excel
                print("TestCaseNumber_" + str(TC_NO) + ": " + "TestScenario is: " + str(Test_Scenario) + "TestCase: " + str(Test_Case))
                Trans_Type = TestInfo_sheet.cell(row=row_value, column=3).value  # Get TransType from Excel
                Exchange = TestInfo_sheet.cell(row=row_value, column=4).value  # Get Exchange from Excel
                Input_OrderType = TestInfo_sheet.cell(row=row_value, column=5).value  # Get Order Type from Excel
                Order_Duration = TestInfo_sheet.cell(row=row_value, column=6).value  # Get Order Duration from Excel
                Trading_Symbol = TestInfo_sheet.cell(row=row_value, column=7).value  # Get Trading Symbol from Excel
                Input_Quantity = TestInfo_sheet.cell(row=row_value,column=8).value  # Get Input Quantity from from Excel
                Input_Price = TestInfo_sheet.cell(row=row_value, column=9).value  # Get Input Price value from Excel
                Input_TriggerPrice = TestInfo_sheet.cell(row=row_value,column=10).value  # Get Input Trigger Price value from Excel
                Input_DiscQty = TestInfo_sheet.cell(row=row_value,column=11).value  # Get Input DiscQty value from Excel
                Quantity = TestInfo_sheet.cell(row=row_value, column=12).value  # Get Quantity from from Excel
                Price = TestInfo_sheet.cell(row=row_value, column=13).value  # Get Price value from Excel
                TriggerPrice = TestInfo_sheet.cell(row=row_value, column=14).value  # Get Trigger Price value from Excel
                DiscQty = TestInfo_sheet.cell(row=row_value, column=15).value  # Get DiscQty value from Excel
                OpenQty = TestInfo_sheet.cell(row=row_value, column=16).value  # Get Open Qty value from Excel
                TradedQty = TestInfo_sheet.cell(row=row_value, column=17).value  # Get Traded Qty value from Excel
                AvgPrice = TestInfo_sheet.cell(row=row_value, column=18).value  # Get Average Price value from Excel
                OrderType = TestInfo_sheet.cell(row=row_value, column=19).value  # Get Average Price value from Excel
                Product = TestInfo_sheet.cell(row=row_value, column=20).value  # Get Product from Excel
                Account_ID = TestInfo_sheet.cell(row=row_value, column=21).value  # Get Account_Id from Excel
                TypeOfOrder = TestInfo_sheet.cell(row=row_value, column=22).value
                ID = OP.PlaceMKTOrder(Result_path,TC_NO,row_value,Test_Scenario,Test_Case,Trans_Type,Exchange,Input_OrderType,Order_Duration,Trading_Symbol,Input_Quantity,Input_Price,Input_TriggerPrice,Input_DiscQty,Quantity,Price,TriggerPrice,DiscQty,Product,Account_ID,Remark,OpenQty,TradedQty,AvgPrice,OrderType,TypeOfOrder)  # send Order values to place / Modify /Cancel Order
                Remark = str(ID)
                row_value = row_value + 1  # Increment Index value of Row
                column_value = column_value + 1  # Increment Index value of column
                # print(str(TC_NO))

    def SLMKT_Order_TestCases(self,Result_path):
        self.Focus_On_Exe()                                                  # Focus On Logged in Exe
        self.Invoke_MarketWatch()
        self.Focus_On_MW()                                                    #Keep Focus in Market Watch
        TestInfo_sheet = Get_File_Path("TestData.xlsx", 4)  # Method to Active the TestData.xlsx sheet
        row_value = 2                                                   # Since Testcase starts from 2nd row in NorenTraderShortCutTestcases.xlsx
        column_value = 1                                                       #Colum start with 1
        TC_NO = 0                                                              #Initialize 0 as TestCase Number
        SUB_TC_NO = 0
        Remark = 0
        for row in range(row_value, 105):
            for column in range(column_value, 105):
                Test_Scenario = TestInfo_sheet.cell(row=row_value, column=1).value  # Get TestScenario from Excel
                if Test_Scenario == None :                                           #Test Scenario Column is None ,increment sub TestCase Number and  do not increment TestCase Number
                    print("DONT")
                else:
                    TC_NO = TC_NO+1                                                   #TestScenario present in column then Increment TestCase Number and initialize
                    #SUB_TC_NO = SUB_TC_NO+1
                    print("üpdated Numbr: "+str(TC_NO))

                Test_Case = TestInfo_sheet.cell(row=row_value, column=2).value         # Get Test Case from Excel
                print("TestCaseNumber_" + str(TC_NO) + ": " + "TestScenario is: "+str(Test_Scenario)+"TestCase: "+str(Test_Case))
                Trans_Type = TestInfo_sheet.cell(row=row_value ,column=3).value          #Get TransType from Excel
                Exchange = TestInfo_sheet.cell(row=row_value ,column=4).value            #Get Exchange from Excel
                Input_OrderType = TestInfo_sheet.cell(row=row_value ,column=5).value          #Get Order Type from Excel
                Order_Duration = TestInfo_sheet.cell(row=row_value ,column=6).value      #Get Order Duration from Excel
                Trading_Symbol = TestInfo_sheet.cell(row=row_value ,column=7).value      #Get Trading Symbol from Excel
                Input_Quantity = TestInfo_sheet.cell(row=row_value, column=8).value      # Get Input Quantity from from Excel
                Input_Price = TestInfo_sheet.cell(row=row_value, column=9).value         # Get Input Price value from Excel
                Input_TriggerPrice = TestInfo_sheet.cell(row=row_value ,column=10).value #Get Input Trigger Price value from Excel
                Input_DiscQty = TestInfo_sheet.cell(row=row_value ,column=11).value      # Get Input DiscQty value from Excel
                Quantity = TestInfo_sheet.cell(row=row_value ,column=12).value           #Get Quantity from from Excel
                Price = TestInfo_sheet.cell(row=row_value ,column=13).value              #Get Price value from Excel
                TriggerPrice = TestInfo_sheet.cell(row=row_value ,column=14).value       #Get Trigger Price value from Excel
                DiscQty = TestInfo_sheet.cell(row=row_value, column=15).value            # Get Disc Qty value from Excel
                OpenQty = TestInfo_sheet.cell(row=row_value ,column=16).value            #Get Open Qty value from Excel
                TradedQty = TestInfo_sheet.cell(row=row_value ,column=17).value          #Get Traded Qty value from Excel
                AvgPrice = TestInfo_sheet.cell(row=row_value ,column=18).value           #Get Average Price value from Excel
                OrderType = TestInfo_sheet.cell(row=row_value ,column=19).value           #Get OrderType value from Excel
                Product = TestInfo_sheet.cell(row=row_value ,column=20).value            #Get Product from Excel
                Account_ID = TestInfo_sheet.cell(row=row_value ,column=21).value         #Get Account_Id from Excel
                TypeOfOrder = TestInfo_sheet.cell(row=row_value ,column=22).value         #Get Type Of from Excel
                ID =OP.PlaceSLMKTOrder(Result_path,TC_NO,row_value,Test_Scenario,Test_Case,Trans_Type,Exchange,Input_OrderType,Order_Duration,Trading_Symbol,Input_Quantity,Input_Price,Input_TriggerPrice,Input_DiscQty,OrderType,Quantity,Price,TriggerPrice,DiscQty,Product,Account_ID,Remark,OpenQty,TradedQty,AvgPrice,TypeOfOrder)     # send Order values to place / Modify /Cancel Order
                Remark = str(ID)
                row_value = row_value + 1                                                 # Increment Index value of Row
                column_value = column_value + 1                                           # Increment Index value of column
                #print(str(TC_NO))