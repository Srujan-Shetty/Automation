import time
from pywinauto.keyboard import send_keys
from pywinauto import keyboard
from UI_Core_Function_Prerequsite import General_Prerequisite
from UI_Core_Function_HelperFunction import Get_File_Path ,Select_From_DropDown
from UI_Core_Function_TestData import LoginWindow_TestData
from pywinauto import Application
from UI_Core_Function_WriteLog import Log
PR = General_Prerequisite()
LoginTestData = LoginWindow_TestData()
Write_Log = Log()
class PriceCursorAction():
    def TC_EXECUTION(self,Test_Case,Window_Name,Result_path):
        print("Test_Case:" + str(Test_Case))
        print("Window_Name:" + str(Window_Name))
        #print("Inv_Usng:" + str(Inv_Usng))
        print("Result_path:" + str(Result_path))
        self.path = LoginTestData.Get_LoginExe_path()  # EXE Path
        self.app = Application(backend='uia').connect(path=self.path)
        self.App_MainWnd = self.app.window(title="Noren Trader", framework_id="WPF")
        self.App_MainWnd.set_focus()
        #PR.Focus_On_Exe()
        #a = PR.App_MainWnd
        match Test_Case:
            case "NORMAL_ORDER_ENTRY_INVOKE_CASES":
                if str(Window_Name) == "Buy Order Entry":
                    Window = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    keyboard.send_keys('{F1}')
                    self.Execute_TestCase_Normal_Order_Entry_Invoke_Cases(Result_path,Window_Name, Test_Case,Window)
                else:
                    Window = self.App_MainWnd.child_window(title="Sell Order Entry", auto_id="OrderEntrywin",class_name="Window")
                    keyboard.send_keys('{F2}')
                    self.Execute_TestCase_Normal_Order_Entry_Invoke_Cases(Result_path, Window_Name, Test_Case, Window)

            case "NORMAL_ORDER_ENTRY_CLOSE_CASES":
                if str(Window_Name) == "Buy Order Entry":
                    Window = self.App_MainWnd.child_window(title="Buy Order Entry", auto_id="OrderEntrywin",class_name="Window")
                    keyboard.send_keys('{F1}')
                    self.Execute_TestCase_delete_and_enter(Result_path,Window_Name, Test_Case, Window)
                else:
                    Window = self.App_MainWnd.child_window(title="Sell Order Entry", auto_id="OrderEntrywin",class_name="Window")
                    keyboard.send_keys('{F2}')
                    self.Execute_TestCase_delete_and_enter(Result_path,Window_Name, Test_Case, Window)

            case "NORMAL_ORDER_ENTRY_MINIMIZE_CASES":
                if str(Window_Name) == "Buy Order Entry":
                    Window = self.App_MainWnd.child_window(title="Buy Order Entry", auto_id="OrderEntrywin",class_name="Window")
                    keyboard.send_keys('{F1}')
                    self.Execute_TestCase_backspace_and_enter(Result_path,Window_Name, Test_Case, Window)
                else:
                    Window = self.App_MainWnd.child_window(title="Sell Order Entry", auto_id="OrderEntrywin",class_name="Window")
                    keyboard.send_keys('{F2}')
                    self.Execute_TestCase_backspace_and_enter(Result_path, Window_Name, Test_Case, Window)
            case "NORMAL_ORDER_ENTRY_MAXIMIZE_CASES":
                if str(Window_Name) == "Buy Order Entry":
                    Window = self.App_MainWnd.child_window(title="Buy Order Entry", auto_id="OrderEntrywin",class_name="Window")
                    keyboard.send_keys('{F1}')
                    self.Execute_TestCase_Price_Increment_NSE(Result_path,Window_Name, Test_Case,Window)
                    self.Execute_TestCase_Price_Increment_CDS(Result_path,Window_Name, Test_Case, Window)
                else:
                    Window = self.App_MainWnd.child_window(title="Sell Order Entry", auto_id="OrderEntrywin",class_name="Window")
                    keyboard.send_keys('{F2}')
                    self.Execute_TestCase_Price_Increment_NSE(Result_path,Window_Name, Test_Case, Window)
                    self.Execute_TestCase_Price_Increment_CDS(Result_path,Window_Name, Test_Case, Window)
            case "NORMAL_ORDER_ENTRY_TAB_FUNCTIONALITY":
                if str(Window_Name) == "Buy Order Entry":
                    Window = self.App_MainWnd.child_window(title="Buy Order Entry", auto_id="OrderEntrywin",class_name="Window")
                    keyboard.send_keys('{F1}')
                    self.Execute_TestCase_Price_Decrement_NSE(Result_path, Window_Name, Test_Case, Window)
                    self.Execute_TestCase_Price_Decrement_CDS(Result_path,Window_Name, Test_Case, Window)
                else:
                    Window = self.App_MainWnd.child_window(title="Sell Order Entry", auto_id="OrderEntrywin",class_name="Window")
                    keyboard.send_keys('{F2}')
                    self.Execute_TestCase_Price_Decrement_NSE(Result_path,Window_Name, Test_Case, Window)
                    self.Execute_TestCase_Price_Decrement_CDS(Result_path, Window_Name, Test_Case, Window)
            case "NORMAL_ORDER-ENTRY_REVERSE_TAB_FUNCTIONALITY":
                if str(Window_Name) == "Buy Order Entry":
                    Window = self.App_MainWnd.child_window(title="Buy Order Entry", auto_id="OrderEntrywin",class_name="Window")
                    keyboard.send_keys('{F1}')
                    self.Execute_TestCase_Right_arrow_move(Result_path,Window_Name, Test_Case, Window)
                else:
                    Window = self.App_MainWnd.child_window(title="Sell Order Entry", auto_id="OrderEntrywin",class_name="Window")
                    keyboard.send_keys('{F2}')
                    self.Execute_TestCase_Right_arrow_move(Result_path,Window_Name, Test_Case, Window)



    def Execute_TestCase_Normal_Order_Entry_Invoke_Cases(self,Result_path,Window_Name, Test_Case,Window):
        TestInfo_sheet = Get_File_Path("TESTCASE.xlsx", 1)
        row_value = 2
        column_value = 1
        TC_No = 0
        while row_value <= 91:  # Loop until row_value reaches 72
            # Focus Price field of Window
            print("wnd is"+str(Window_Name))
            # Select_PriceTextBox = Window.child_window(auto_id="PriceTxt", class_name="UTF_EditBox")
            # Enter_Price = Select_PriceTextBox.child_window(auto_id="PART_Editor", class_name="TextBox").set_focus()

            #Get TC_No ,DgtsToselect,Input_Price,Position,Expected_Price from file "TESTCASE.xlsx"
            TC_NO = TestInfo_sheet.cell(row=row_value, column=1).value
            DgtsToselect = TestInfo_sheet.cell(row=row_value, column=2).value
            Input_Price = TestInfo_sheet.cell(row=row_value, column=3).value
            Position = TestInfo_sheet.cell(row=row_value, column=4).value
            Expected_Price = TestInfo_sheet.cell(row=row_value, column=5).value
            print("TC_NO:"+str(TC_No))
            print("DgtsToselect:" + str(DgtsToselect))
            print("Input_Price:" + str(Input_Price))
            print("Position:" + str(Position))
            print("Expected_Price:" + str(Expected_Price))
            print("Row value:"+str(row_value))

            match DgtsToselect:
                case "ENTER_PRICE":
                    EnteredPrice  = str(Expected_Price)
                    Enter_Price.set_edit_text(str(Expected_Price))
                    Actual_Price = Enter_Price.get_value()
                    print("Actual_Price:", Enter_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.general_log_select_and_edit(Result_path,EnteredPrice,Test_Case,row_value,Window_Name,TC_NO,DgtsToselect,Result,Input_Price,Position,Expected_Price,Actual_Price)
                case "selecting 1 digit":
                    if str(Position) == "1":
                        time.sleep(10)
                        Enter_Price.type_keys('{HOME}+{RIGHT}', set_foreground=True)
                        time.sleep(10)
                        # Type 'Input_Price' to replace selected 'Position'
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2":
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)  # replace with 2
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5":
                        # Step 1: Move cursor before the 5th character (HOME + 4x RIGHT)
                        # Step 2: Select 1 character (the 5th character)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}', set_foreground=True)
                        # Step 3: Type replacement value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "6":
                        #Move to 6th position: HOME + 5x RIGHT
                        #Select 6th digit: SHIFT + RIGHT
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}', set_foreground=True)
                        # Replace the selected digit with Input_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "7":
                        #Move to 7th position: HOME + 6x RIGHT
                        #Select 7th digit: SHIFT + RIGHT
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}', set_foreground=True)
                        # Replace the selected digit with Input_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "8":
                        #Move to 8th position: HOME + 7x RIGHT
                        #Select 8th digit: SHIFT + RIGHT
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}', set_foreground=True)
                        # Replace the selected digit with Input_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "9":
                        #Move to 9th position: HOME + 8x RIGHT
                        #Select 9th digit: SHIFT + RIGHT
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}', set_foreground=True)
                        # Replace the selected digit with Input_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "10":
                        # Move to 10th position: HOME + 9x RIGHT
                        # Select 10th digit: SHIFT + RIGHT
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}',set_foreground=True)
                        # Replace the selected digit with Input_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)

                    elif str(Position) == "5,6":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}{RIGHT}',set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                        #Result = self.Check_Result(Expected_Price, Actual_Price)
                        #Write_Log.Write_SELECT_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect, Expected_Price, Actual_Price, Result,row_value)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_SELECT_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect, Position,Expected_Price, Actual_Price, Result, row_value)
                case "selecting 2 digits":
                    if str(Position) == "1,2":
                        # Step 1: Move cursor to the beginning and select 2 characters
                        Enter_Price.type_keys('{HOME}+{RIGHT}+{RIGHT}', set_foreground=True)
                        # Step 2: Type the new value (e.g., replace with "42")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2,3":
                        # Step 1: Move to before position 2
                        # Step 2: Select 2 characters: positions 2 and 3
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        # Step 3: Type the new value (e.g., "42")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3,4":
                        # Step 1: Move to just before position 3 (HOME + 2x RIGHT)
                        # Step 2: Hold Shift and press RIGHT twice to select positions 3 and 4
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        # Step 3: Replace selected text with your input value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4,5":
                        # Step 1: Move to position 4 (HOME + 3x RIGHT)
                        # Step 2: Shift + RIGHT x2 to select 2 digits (position 4 and 5)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        # Step 3: Replace with your input (e.g., "99")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5,6":
                        # Step 1: Move cursor before position 5 (HOME + 4x RIGHT)
                        # Step 2: Shift + RIGHT x2 to select characters 5 and 6
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        # Step 3: Replace selected characters with the new value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "6,7":
                        # Step 1: Move to the 6th character (HOME + 5x RIGHT)
                        # Step 2: Shift + RIGHT twice to select positions 6 and 7
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        # Step 3: Replace the selected characters with your value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "8,9":
                        # Step 1: Move to the 8th character (HOME + 5x RIGHT)
                        # Step 2: Shift + RIGHT twice to select positions 8 and 9
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        # Step 3: Replace the selected characters with your value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "9,10":
                        # Step 1: Move to the 9th character (HOME + 8x RIGHT)
                        # Step 2: Shift + RIGHT twice to select positions 9 and 10
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        # Step 3: Replace the selected characters with your value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_SELECT_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position, Expected_Price, Actual_Price, Result, row_value)
                case "selecting 3 digits":
                    if str(Position) == "1,2,3":
                        # Step 1: Move to start of the text
                        # Step 2: Select 3 characters (positions 1, 2, and 3)
                        Enter_Price.type_keys('{HOME}+{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        # Step 3: Replace selected text with new value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2,3,4":
                        # Step 1: Move cursor just before position 2 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 2, 3, and 4)
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3,4,5":
                        # Step 1: Move cursor just before position 3 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 3, 4, and 5)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4,5,6":
                        # Step 1: Move cursor just before position 4 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 4, 5, and 6)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5,6,7":
                        # Step 1: Move cursor just before position 5 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 5, 6, and 7)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "7,8":
                        # Step 1: Move cursor just before position 6 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 6, 7, and 8)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_SELECT_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position, Expected_Price, Actual_Price, Result, row_value)
                case "selecting 4 digits":
                    if str(Position) == "1,2,3,4":
                        #Enter_Price.type_keys('{HOME}', set_foreground=True)  # Move to beginning
                        Enter_Price.type_keys('{HOME}+{RIGHT 4}', set_foreground=True)  # Select 4 digits
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)  # Type new value
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2,3,4,5":
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3,4,5,6":
                        #Enter_Price.type_keys('{HOME}{RIGHT 3}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4,5,6,7":
                        #Enter_Price.type_keys('{HOME}{RIGHT 4}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5,6,7,8":
                        #Enter_Price.type_keys('{HOME}{RIGHT 5}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "7,8,9,10":
                        #Enter_Price.type_keys('{HOME}{RIGHT 7}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_SELECT_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position, Expected_Price, Actual_Price, Result, row_value)
                case "selecting 5 digits":
                    if str(Position) == "1,2,3,4,5":
                        Enter_Price.type_keys('{HOME}+{RIGHT 5}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2,3,4,5,6":
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT 5}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3,4,5,6,7":
                        # Move to position 3 (2x RIGHT), then select next 5 digits (3–7)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT 5}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4,5,6,7,8":
                        # Move to position 4 (3x RIGHT), then select 5 digits
                        #Enter_Price.type_keys('{HOME}{RIGHT 3}+{RIGHT 5}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT 5}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_SELECT_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position, Expected_Price, Actual_Price, Result, row_value)
                case "selecting 6 digits":
                    if str(Position) == "1,2,3,4,5,6":
                        Enter_Price.type_keys('{HOME}+{RIGHT 6}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_SELECT_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect, Position,Expected_Price, Actual_Price, Result, row_value)


                case "edit 2 decimal places":
                    if str(Position) == "6":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "6,7":
                        # Move to the 6th character and select 2 characters (6th and 7th)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        # Replace with the desired value (e.g., "35")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "7,8":
                        # Move to the 7th character and select 2 characters (7th and 8th)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        # Replace with the desired value (e.g., "35")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "8,9":
                        # Move to the 7th character and select 2 characters (7th and 8th)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        # Replace with the desired value (e.g., "35")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "9,10":
                        # Move to the 7th character and select 2 characters (7th and 8th)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        # Replace with the desired value (e.g., "35")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_SELECT_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position,Expected_Price, Actual_Price, Result, row_value)

            row_value += 1
    def Execute_TestCase_delete_and_enter(self,Result_path,Window_Name, Test_Case,Window):
        TestInfo_sheet = Get_File_Path("TESTCASE.xlsx", 1)
        row_value = 2
        column_value = 1
        TC_No = 0
        while row_value <= 122:  # Loop until row_value reaches 72
            # Focus Price field of Window
            print("wnd is"+str(Window_Name))
            Select_PriceTextBox = Window.child_window(auto_id="PriceTxt", class_name="UTF_EditBox")
            Enter_Price = Select_PriceTextBox.child_window(auto_id="PART_Editor", class_name="TextBox").set_focus()

            #Get TC_No ,DgtsToselect,Input_Price,Position,Expected_Price from file "TESTCASE.xlsx"
            TC_NO = TestInfo_sheet.cell(row=row_value, column=1).value
            DgtsToselect = TestInfo_sheet.cell(row=row_value, column=2).value
            Input_Price = TestInfo_sheet.cell(row=row_value, column=3).value
            Position = TestInfo_sheet.cell(row=row_value, column=4).value
            Expected_Price = TestInfo_sheet.cell(row=row_value, column=5).value
            print("TC_NO:"+str(TC_No))
            print("DgtsToselect:" + str(DgtsToselect))
            print("Input_Price:" + str(Input_Price))
            print("Position:" + str(Position))
            print("Expected_Price:" + str(Expected_Price))


            match DgtsToselect:
                case "ENTER_PRICE":
                    EnteredPrice  = str(Expected_Price)
                    Enter_Price.set_edit_text(str(Expected_Price))
                    Actual_Price = Enter_Price.get_value()
                    print("Actual_Price:", Enter_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.general_log_delete_and_edit(Result_path,EnteredPrice,Test_Case,row_value,Window_Name,TC_NO,DgtsToselect,Result,Input_Price,Position,Expected_Price,Actual_Price)
                case "selecting 1 digit":
                    if str(Position) == "1":
                        Enter_Price.type_keys('{HOME}', set_foreground=True)           # Step 1: Move to the 1st position
                        Enter_Price.type_keys('{DEL}', set_foreground=True)            # Step 2: Delete the character at 1st position
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)   # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                         # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2":
                        Enter_Price.type_keys('{HOME}{RIGHT}', set_foreground=True)    # Step 1: Move to the 2nd position (HOME + 1x RIGHT)
                        Enter_Price.type_keys('{DEL}',set_foreground=True)             # Step 2: Delete the character at 2nd position
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)   # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                         # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}',set_foreground=True)           # Step 1: Move to the 3rd position (HOME + 2x RIGHT)
                        Enter_Price.type_keys('{DEL}',set_foreground=True)                          # Step 2: Delete the character at 3rd position
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)                # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                      # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 4th position (HOME + 3x RIGHT)
                        Enter_Price.type_keys('{DEL}',set_foreground=True)                        # Step 2: Delete the character at 4th position
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)              # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                    # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 5th position (HOME + 4x RIGHT)
                        Enter_Price.type_keys('{DEL}',set_foreground=True)                        # Step 2: Delete the character at 5th position
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)              # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                    # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "6":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 6th position (HOME + 5x RIGHT)
                        Enter_Price.type_keys('{DEL}',set_foreground=True)                                      # Step 2: Delete the character at 6th position
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)                            # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                                  # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "7":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 7th position (HOME + 6x RIGHT)
                        Enter_Price.type_keys('{DEL}',set_foreground=True)                                             # Step 2: Delete the character at 7th position
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)                                   # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                                         # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "8":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 8th position (HOME + 7x RIGHT)
                        Enter_Price.type_keys('{DEL}',set_foreground=True)                                                    # Step 2: Delete the character at 8th position
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)                                          # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()  # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "9":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 9th position (HOME + 8x RIGHT)
                        Enter_Price.type_keys('{DEL}',set_foreground=True)                                                           # Step 2: Delete the character at 9th position
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)                                                 # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                                                       # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "10":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 10th position (HOME + 9x RIGHT)
                        Enter_Price.type_keys('{DEL}',set_foreground=True)                                                                  # Step 2: Delete the character at 10th position
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)                                                        # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                                                              # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)

                    elif str(Position) == "5,6":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}{RIGHT}',set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)

                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DELETE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect, Position,Expected_Price, Actual_Price, Result, row_value)
                case "selecting 2 digits":
                    if str(Position) == "1,2":
                        # Step 1: Move cursor to the beginning and select 2 characters
                        Enter_Price.type_keys('{HOME}+{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 2: Type the new value (e.g., replace with "42")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2,3":
                        # Step 1: Move to before position 2
                        # Step 2: Select 2 characters: positions 2 and 3
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 3: Type the new value (e.g., "42")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3,4":
                        # Step 1: Move to just before position 3 (HOME + 2x RIGHT)
                        # Step 2: Hold Shift and press RIGHT twice to select positions 3 and 4
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 3: Replace selected text with your input value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4,5":
                        # Step 1: Move to position 4 (HOME + 3x RIGHT)
                        # Step 2: Shift + RIGHT x2 to select 2 digits (position 4 and 5)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        # Step 3: Replace with your input (e.g., "99")
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5,6":
                        # Step 1: Move cursor before position 5 (HOME + 4x RIGHT)
                        # Step 2: Shift + RIGHT x2 to select characters 5 and 6
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 3: Replace selected characters with the new value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "6,7":
                        # Step 1: Move to the 6th character (HOME + 5x RIGHT)
                        # Step 2: Shift + RIGHT twice to select positions 6 and 7
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 3: Replace the selected characters with your value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "8,9":
                        # Step 1: Move to the 8th character (HOME + 5x RIGHT)
                        # Step 2: Shift + RIGHT twice to select positions 8 and 9
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 3: Replace the selected characters with your value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "9,10":
                        # Step 1: Move to the 9th character (HOME + 8x RIGHT)
                        # Step 2: Shift + RIGHT twice to select positions 9 and 10
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 3: Replace the selected characters with your value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DELETE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position, Expected_Price, Actual_Price, Result, row_value)
                case "selecting 3 digits":
                    if str(Position) == "1,2,3":
                        # Step 1: Move to start of the text
                        # Step 2: Select 3 characters (positions 1, 2, and 3)
                        Enter_Price.type_keys('{HOME}+{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 3: Replace selected text with new value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2,3,4":
                        # Step 1: Move cursor just before position 2 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 2, 3, and 4)
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3,4,5":
                        # Step 1: Move cursor just before position 3 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 3, 4, and 5)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4,5,6":
                        # Step 1: Move cursor just before position 4 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 4, 5, and 6)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5,6,7":
                        # Step 1: Move cursor just before position 5 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 5, 6, and 7)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "7,8":
                        # Step 1: Move cursor just before position 6 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 6, 7, and 8)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DELETE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position, Expected_Price, Actual_Price, Result, row_value)
                case "selecting 4 digits":
                    if str(Position) == "1,2,3,4":
                        #Enter_Price.type_keys('{HOME}', set_foreground=True)  # Move to beginning
                        Enter_Price.type_keys('{HOME}+{RIGHT 4}', set_foreground=True)  # Select 4 digits
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)  # Type new value
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2,3,4,5":
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3,4,5,6":
                        #Enter_Price.type_keys('{HOME}{RIGHT 3}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4,5,6,7":
                        #Enter_Price.type_keys('{HOME}{RIGHT 4}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5,6,7,8":
                        #Enter_Price.type_keys('{HOME}{RIGHT 5}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "7,8,9,10":
                        #Enter_Price.type_keys('{HOME}{RIGHT 7}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DELETE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position, Expected_Price, Actual_Price, Result, row_value)
                case "selecting 5 digits":
                    if str(Position) == "1,2,3,4,5":
                        Enter_Price.type_keys('{HOME}+{RIGHT 5}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2,3,4,5,6":
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT 5}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3,4,5,6,7":
                        # Move to position 3 (2x RIGHT), then select next 5 digits (3–7)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT 5}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4,5,6,7,8":
                        # Move to position 4 (3x RIGHT), then select 5 digits
                        #Enter_Price.type_keys('{HOME}{RIGHT 3}+{RIGHT 5}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT 5}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DELETE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position, Expected_Price, Actual_Price, Result, row_value)
                case "selecting 6 digits":
                    if str(Position) == "1,2,3,4,5,6":
                        Enter_Price.type_keys('{HOME}+{RIGHT 6}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DELETE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect, Position,Expected_Price, Actual_Price, Result, row_value)


                case "edit 2 decimal places":
                    if str(Position) == "6":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}', set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "6,7":
                        # Move to the 6th character and select 2 characters (6th and 7th)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Replace with the desired value (e.g., "35")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "7,8":
                        # Move to the 7th character and select 2 characters (7th and 8th)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Replace with the desired value (e.g., "35")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "8,9":
                        # Move to the 7th character and select 2 characters (7th and 8th)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Replace with the desired value (e.g., "35")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "9,10":
                        # Move to the 7th character and select 2 characters (7th and 8th)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        Enter_Price.type_keys('{DEL}', set_foreground=True)
                        # Replace with the desired value (e.g., "35")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DELETE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position,Expected_Price, Actual_Price, Result, row_value)
            row_value += 1
    def Execute_TestCase_backspace_and_enter(self, Result_path,Window_Name, Test_Case, Window):
        TestInfo_sheet = Get_File_Path("TESTCASE.xlsx", 1)
        row_value = 2
        column_value = 1
        TC_No = 0
        while row_value <= 122:  # Loop until row_value reaches 72
            # Focus Price field of Window
            print("wnd is" + str(Window_Name))
            Select_PriceTextBox = Window.child_window(auto_id="PriceTxt", class_name="UTF_EditBox")
            Enter_Price = Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()

            # Get TC_No ,DgtsToselect,Input_Price,Position,Expected_Price from file "TESTCASE.xlsx"
            TC_NO = TestInfo_sheet.cell(row=row_value, column=1).value
            DgtsToselect = TestInfo_sheet.cell(row=row_value, column=2).value
            Input_Price = TestInfo_sheet.cell(row=row_value, column=3).value
            Position = TestInfo_sheet.cell(row=row_value, column=4).value
            Expected_Price = TestInfo_sheet.cell(row=row_value, column=5).value
            print("TC_NO:" + str(TC_No))
            print("DgtsToselect:" + str(DgtsToselect))
            print("Input_Price:" + str(Input_Price))
            print("Position:" + str(Position))
            print("Expected_Price:" + str(Expected_Price))

            match DgtsToselect:
                case "ENTER_PRICE":
                    EnteredPrice  = str(Expected_Price)
                    Enter_Price.set_edit_text(str(Expected_Price))
                    Actual_Price = Enter_Price.get_value()
                    print("Actual_Price:", Enter_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.general_log_backspace_and_edit(Result_path,EnteredPrice,Test_Case,row_value,Window_Name,TC_NO,DgtsToselect,Result,Input_Price,Position,Expected_Price,Actual_Price)
                case "selecting 1 digit":
                    if str(Position) == "1":
                        Enter_Price.type_keys('{HOME}{RIGHT}', set_foreground=True)         # Step 1: Move to the 1st position
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)    # Step 2: Backspace to delete the 1st character
                        time.sleep(5)
                        Enter_Price.type_keys(str(Input_Price),set_foreground=True)  # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                       # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 2nd position (HOME + 1x RIGHT)
                        #Enter_Price.type_keys('{BACKSPACE}',set_foreground=True)  # Step 2: Backspace to delete the 2bd character
                        send_keys('{BACKSPACE}')
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)  # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()  # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)

                    elif str(Position) == "3":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 3rd position (HOME + 2x RIGHT)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)          # Step 2: Backspace to delete the 3st character
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        Enter_Price.type_keys(str(Input_Price),set_foreground=True)        # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                             # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)     # Step 1: Move to the 4th position (HOME + 3x RIGHT)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)                    # Step 2: Backspace to delete the 4th character
                        send_keys('{BACKSPACE}')
                        Enter_Price.type_keys(str(Input_Price),set_foreground=True)                  # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                       # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5":
                        Enter_Price.type_keys('{HOME}{RIGHT 5}',set_foreground=True)  # Step 1: Move to the 5th position (HOME + 4x RIGHT)
                        time.sleep(5)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)                        # Step 2: Backspace to delete the 5th character
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        Enter_Price.type_keys(str(Input_Price),set_foreground=True)                      # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                           # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "6":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 6th position (HOME + 5x RIGHT)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)                               # Step 2: Backspace to delete the 6th character
                        send_keys('{BACKSPACE}')
                        Enter_Price.type_keys(str(Input_Price),set_foreground=True)                             # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                                  # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "7":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 7th position (HOME + 6x RIGHT)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)                                     # Step 2: Backspace to delete the 7th character
                        send_keys('{BACKSPACE}')
                        Enter_Price.type_keys(str(Input_Price),set_foreground=True)                                   # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                                         # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "8":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 8th position (HOME + 7x RIGHT)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)                                             # Step 2: Backspace to delete the 8th character
                        send_keys('{BACKSPACE}')
                        Enter_Price.type_keys(str(Input_Price),set_foreground=True)                                           # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                                                # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "9":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 9th position (HOME + 8x RIGHT)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)                                                    # Step 2: Backspace to delete the 9th character
                        send_keys('{BACKSPACE}')
                        Enter_Price.type_keys(str(Input_Price),set_foreground=True)                                                  # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                                                       # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "10":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}',set_foreground=True)  # Step 1: Move to the 10th position (HOME + 9x RIGHT)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)                                                           # Step 2: Backspace to delete the 10th character
                        send_keys('{BACKSPACE}')
                        Enter_Price.type_keys(str(Input_Price),set_foreground=True)                                                         # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                                                                              # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)

                    elif str(Position) == "5,6":
                        #Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}{RIGHT}',set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT 4}+{RIGHT}', set_foreground=True)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)

                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_BACKSPACE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position,Expected_Price, Actual_Price, Result, row_value)
                case "selecting 2 digits":
                    if str(Position) == "1,2":
                        # Step 1: Move cursor to the beginning and select 2 characters
                        Enter_Price.type_keys('{HOME}+{RIGHT}+{RIGHT}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 2: Type the new value (e.g., replace with "42")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2,3":
                        # Step 1: Move to before position 2
                        # Step 2: Select 2 characters: positions 2 and 3
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 3: Type the new value (e.g., "42")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3,4":
                        # Step 1: Move to just before position 3 (HOME + 2x RIGHT)
                        # Step 2: Hold Shift and press RIGHT twice to select positions 3 and 4
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 3: Replace selected text with your input value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4,5":
                        # Step 1: Move to position 4 (HOME + 3x RIGHT)
                        # Step 2: Shift + RIGHT x2 to select 2 digits (position 4 and 5)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        # Step 3: Replace with your input (e.g., "99")
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5,6":
                        # Step 1: Move cursor before position 5 (HOME + 4x RIGHT)
                        # Step 2: Shift + RIGHT x2 to select characters 5 and 6
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 3: Replace selected characters with the new value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "6,7":
                        # Step 1: Move to the 6th character (HOME + 5x RIGHT)
                        # Step 2: Shift + RIGHT twice to select positions 6 and 7
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 3: Replace the selected characters with your value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "8,9":
                        # Step 1: Move to the 8th character (HOME + 5x RIGHT)
                        # Step 2: Shift + RIGHT twice to select positions 8 and 9
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 3: Replace the selected characters with your value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "9,10":
                        # Step 1: Move to the 9th character (HOME + 8x RIGHT)
                        # Step 2: Shift + RIGHT twice to select positions 9 and 10
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 3: Replace the selected characters with your value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_BACKSPACE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position,Expected_Price, Actual_Price, Result, row_value)
                case "selecting 3 digits":
                    if str(Position) == "1,2,3":
                        # Step 1: Move to start of the text
                        # Step 2: Select 3 characters (positions 1, 2, and 3)
                        Enter_Price.type_keys('{HOME}+{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 3: Replace selected text with new value
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2,3,4":
                        # Step 1: Move cursor just before position 2 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 2, 3, and 4)
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3,4,5":
                        # Step 1: Move cursor just before position 3 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 3, 4, and 5)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4,5,6":
                        # Step 1: Move cursor just before position 4 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 4, 5, and 6)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5,6,7":
                        # Step 1: Move cursor just before position 5 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 5, 6, and 7)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "7,8":
                        # Step 1: Move cursor just before position 6 (HOME + 1x RIGHT)
                        # Step 2: Select 3 characters (positions 6, 7, and 8)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Step 3: Replace with Edit_Price
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_BACKSPACE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position,Expected_Price, Actual_Price, Result, row_value)
                case "selecting 4 digits":
                    if str(Position) == "1,2,3,4":
                        # Enter_Price.type_keys('{HOME}', set_foreground=True)  # Move to beginning
                        Enter_Price.type_keys('{HOME}+{RIGHT 4}', set_foreground=True)  # Select 4 digits
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)  # Type new value
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2,3,4,5":
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3,4,5,6":
                        # Enter_Price.type_keys('{HOME}{RIGHT 3}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4,5,6,7":
                        # Enter_Price.type_keys('{HOME}{RIGHT 4}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT 4}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5,6,7,8":
                        # Enter_Price.type_keys('{HOME}{RIGHT 5}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT 4}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "7,8,9,10":
                        # Enter_Price.type_keys('{HOME}{RIGHT 7}+{RIGHT 4}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT 4}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_BACKSPACE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position,Expected_Price, Actual_Price, Result, row_value)
                case "selecting 5 digits":
                    if str(Position) == "1,2,3,4,5":
                        Enter_Price.type_keys('{HOME}+{RIGHT 5}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "2,3,4,5,6":
                        Enter_Price.type_keys('{HOME}{RIGHT}+{RIGHT 5}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "3,4,5,6,7":
                        # Move to position 3 (2x RIGHT), then select next 5 digits (3–7)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}+{RIGHT 5}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4,5,6,7,8":
                        # Move to position 4 (3x RIGHT), then select 5 digits
                        # Enter_Price.type_keys('{HOME}{RIGHT 3}+{RIGHT 5}', set_foreground=True)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}+{RIGHT 5}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_BACKSPACE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position,Expected_Price, Actual_Price, Result, row_value)
                case "selecting 6 digits":
                    if str(Position) == "1,2,3,4,5,6":
                        Enter_Price.type_keys('{HOME}+{RIGHT 6}', set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_BACKSPACE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position,Expected_Price, Actual_Price, Result, row_value)

                case "edit 2 decimal places":
                    if str(Position) == "6":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "6,7":
                        # Move to the 6th character and select 2 characters (6th and 7th)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Replace with the desired value (e.g., "35")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "7,8":
                        # Move to the 7th character and select 2 characters (7th and 8th)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Replace with the desired value (e.g., "35")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "8,9":
                        # Move to the 7th character and select 2 characters (7th and 8th)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Replace with the desired value (e.g., "35")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "9,10":
                        # Move to the 7th character and select 2 characters (7th and 8th)
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}+{RIGHT}+{RIGHT}',set_foreground=True)
                        time.sleep(0.2)
                        send_keys('{BACKSPACE}')
                        time.sleep(0.2)
                        #Enter_Price.type_keys('{BACKSPACE}', set_foreground=True)
                        # Replace with the desired value (e.g., "35")
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)
                        Actual_Price = Enter_Price.get_value()
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_BACKSPACE_AND_EDIT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect,Position,Expected_Price, Actual_Price, Result, row_value)
            row_value += 1
    def Execute_TestCase_Price_Increment_NSE(self, Result_path, Window_Name, Test_Case, Window):
        # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
        self.select_Exch = Window.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
        Select_From_DropDown(self.select_Exch, "NSE")

        # keyboard.send_keys('{TAB}')
        # keyboard.send_keys('{TAB}')

        # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
        self.Custom = Window.child_window(auto_id="ScripSearchBox", class_name="ScripSearchBox")
        self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
        self.Input_TradingSymbol.set_edit_text(" ")
        self.Input_TradingSymbol.set_edit_text("ACC-EQ")
        self.Input_TradingSymbol.set_focus()
        keyboard.send_keys('{DOWN}')
        keyboard.send_keys('{ENTER}')
        # keyboard.send_keys('{TAB}')

        #----------------------------------Get TestExecution data from TESTCASE file-------------------------------------------------#
        TestInfo_sheet = Get_File_Path("TESTCASE.xlsx", 1)
        row_value = 123
        TC_No = 0
        while row_value <= 133:  # Loop until row_value reaches 72
            # Focus Price field of Window
            print("wnd is" + str(Window_Name))
            Select_PriceTextBox = Window.child_window(auto_id="PriceTxt", class_name="UTF_EditBox")
            Enter_Price = Select_PriceTextBox.child_window(auto_id="PART_Editor", class_name="TextBox").set_focus()

            # Get TC_No ,Expected_Price from file "TESTCASE.xlsx"
            TC_NO = TestInfo_sheet.cell(row=row_value, column=1).value
            DgtsToselect = TestInfo_sheet.cell(row=row_value, column=2).value
            Expected_Price = TestInfo_sheet.cell(row=row_value, column=5).value
            print("TC_NO:" + str(TC_No))
            print("DgtsToselect:" + str(DgtsToselect))
            print("Expected_Price:" + str(Expected_Price))

            match DgtsToselect:
                case "ENTER_PRICE":
                    EnteredPrice  = str(Expected_Price)
                    Enter_Price.set_edit_text(str(Expected_Price))
                    Actual_Price = Enter_Price.get_value()
                    print("Actual_Price:", Enter_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.general_log_Increment(Result_path, EnteredPrice, Test_Case, row_value,Window_Name,"NSE",TC_NO, DgtsToselect, Result,Expected_Price, Actual_Price)

                case "Increment Once":
                    Enter_Price.set_focus()
                    keyboard.send_keys("{UP}")
                    Actual_Price = Enter_Price.get_value()              # Step 4: Fetch the actual price
                    print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_UP_ARROW_INCREMENT_CASES(Result_path, TC_NO, Window_Name, Test_Case, "NSE",row_value, DgtsToselect, Result,Expected_Price, Actual_Price)

                case "Increment 10 times":
                    Enter_Price.set_focus()
                    for i in range(10):
                        keyboard.send_keys("{UP}")
                        time.sleep(0.2)                                # small delay between key presses
                    Actual_Price = Enter_Price.get_value()             # Step 4: Fetch the actual price
                    print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_UP_ARROW_INCREMENT_CASES(Result_path, TC_NO, Window_Name, Test_Case, "NSE",row_value, DgtsToselect, Result,Expected_Price, Actual_Price)
                case "Increment 26 times":
                    Enter_Price.set_focus()
                    for i in range(26):
                        keyboard.send_keys("{UP}")
                        time.sleep(0.2)  # small delay between key presses
                    Actual_Price = Enter_Price.get_value()  # Step 4: Fetch the actual price
                    print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_UP_ARROW_INCREMENT_CASES(Result_path, TC_NO, Window_Name, Test_Case, "NSE",row_value, DgtsToselect, Result,Expected_Price, Actual_Price)
            row_value += 1
    def Execute_TestCase_Price_Increment_CDS(self, Result_path,Window_Name, Test_Case, Window):
        # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
        self.select_Exch = Window.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
        Select_From_DropDown(self.select_Exch, "CDS")
        keyboard.send_keys('{TAB}')
        keyboard.send_keys('{TAB}')

        # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
        self.Custom = Window.child_window(auto_id="ScripSearchBox", class_name="ScripSearchBox")
        self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
        self.Input_TradingSymbol.set_edit_text(" ")
        # self.Input_TradingSymbol.set_edit_text("USDINR")
        self.Input_TradingSymbol.set_focus()
        keyboard.send_keys('{DOWN}')
        keyboard.send_keys('{ENTER}')
        keyboard.send_keys('{TAB}')
        #----------------------------------Get TestExecution data from TESTCASE file-------------------------------------------------#
        TestInfo_sheet = Get_File_Path("TESTCASE.xlsx", 1)
        row_value = 123
        TC_No = 0
        while row_value <= 133:  # Loop until row_value reaches 72
            # Focus Price field of Window
            print("wnd is" + str(Window_Name))
            Select_PriceTextBox = Window.child_window(auto_id="PriceTxt", class_name="UTF_EditBox")
            Enter_Price = Select_PriceTextBox.child_window(auto_id="PART_Editor", class_name="TextBox").set_focus()

            # Get TC_No ,Expected_Price from file "TESTCASE.xlsx"
            TC_NO = TestInfo_sheet.cell(row=row_value, column=1).value
            DgtsToselect = TestInfo_sheet.cell(row=row_value, column=2).value
            Expected_Price = TestInfo_sheet.cell(row=row_value, column=6).value
            print("TC_NO:" + str(TC_No))
            print("DgtsToselect:" + str(DgtsToselect))
            print("Expected_Price:" + str(Expected_Price))

            match DgtsToselect:
                case "ENTER_PRICE":
                    EnteredPrice  = str(Expected_Price)
                    Enter_Price.set_edit_text(str(Expected_Price))
                    Actual_Price = Enter_Price.get_value()
                    print("Actual_Price:", Enter_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.general_log_Increment(Result_path, EnteredPrice, Test_Case, row_value,Window_Name,"CDS",TC_NO, DgtsToselect, Result,Expected_Price, Actual_Price)
                case "Increment Once":
                    Enter_Price.set_focus()
                    keyboard.send_keys("{UP}")
                    Actual_Price = Enter_Price.get_value()              # Step 4: Fetch the actual price
                    print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_UP_ARROW_INCREMENT_CASES(Result_path, TC_NO, Window_Name, Test_Case, "CDS",row_value, DgtsToselect, Result,Expected_Price, Actual_Price)
                case "Increment 10 times":
                    Enter_Price.set_focus()
                    for i in range(10):
                        keyboard.send_keys("{UP}")
                        time.sleep(0.2)                                # small delay between key presses
                    Actual_Price = Enter_Price.get_value()             # Step 4: Fetch the actual price
                    print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_UP_ARROW_INCREMENT_CASES(Result_path, TC_NO, Window_Name, Test_Case, "CDS",row_value, DgtsToselect, Result,Expected_Price, Actual_Price)
                case "Increment 26 times":
                    Enter_Price.set_focus()
                    for i in range(26):
                        keyboard.send_keys("{UP}")
                        time.sleep(0.2)  # small delay between key presses
                    Actual_Price = Enter_Price.get_value()  # Step 4: Fetch the actual price
                    print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_UP_ARROW_INCREMENT_CASES(Result_path, TC_NO, Window_Name, Test_Case, "CDS",row_value, DgtsToselect, Result,Expected_Price, Actual_Price)
            row_value += 1
    def Execute_TestCase_Price_Decrement_NSE(self, Result_path, Window_Name, Test_Case, Window):
        # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
        self.select_Exch = Window.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
        Select_From_DropDown(self.select_Exch, "NSE")
        time.sleep(5)
        keyboard.send_keys('{TAB}')
        keyboard.send_keys('{TAB}')
        time.sleep(5)
        # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
        self.Custom = Window.child_window(auto_id="ScripSearchBox", class_name="ScripSearchBox")
        self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
        self.Input_TradingSymbol.set_edit_text(" ")
        self.Input_TradingSymbol.set_edit_text("BIOCON-EQ")
        self.Input_TradingSymbol.set_focus()
        keyboard.send_keys('{DOWN}')
        keyboard.send_keys('{ENTER}')
        keyboard.send_keys('{TAB}')
        #----------------------------------Get TestExecution data from TESTCASE file-------------------------------------------------#
        TestInfo_sheet = Get_File_Path("TESTCASE.xlsx", 1)
        row_value = 134
        TC_No = 0
        while row_value <= 146:  # Loop until row_value reaches 72
            # Focus Price field of Window
            print("wnd is" + str(Window_Name))
            Select_PriceTextBox = Window.child_window(auto_id="PriceTxt", class_name="UTF_EditBox")
            Enter_Price = Select_PriceTextBox.child_window(auto_id="PART_Editor", class_name="TextBox").set_focus()

            # Get TC_No ,Expected_Price from file "TESTCASE.xlsx"
            TC_NO = TestInfo_sheet.cell(row=row_value, column=1).value
            DgtsToselect = TestInfo_sheet.cell(row=row_value, column=2).value
            Expected_Price = TestInfo_sheet.cell(row=row_value, column=5).value
            print("TC_NO:" + str(TC_No))
            print("DgtsToselect:" + str(DgtsToselect))
            print("Expected_Price:" + str(Expected_Price))

            match DgtsToselect:
                case "ENTER_PRICE":
                    EnteredPrice  = str(Expected_Price)
                    Enter_Price.set_edit_text(str(Expected_Price))
                    Actual_Price = Enter_Price.get_value()
                    print("Actual_Price:", Enter_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.general_log_Decrement(Result_path, EnteredPrice, Test_Case, row_value,Window_Name,"NSE",TC_NO, DgtsToselect, Result,Expected_Price, Actual_Price)
                case "Decrement Once":
                    Enter_Price.set_focus()
                    keyboard.send_keys("{DOWN}")
                    Actual_Price = Enter_Price.get_value()              # Step 4: Fetch the actual price
                    print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DOWN_ARROW_DECREMENT_CASES(Result_path, TC_NO, Window_Name, Test_Case, "NSE",row_value, DgtsToselect, Result,Expected_Price, Actual_Price)
                case "Decrement 10 times":
                    Enter_Price.set_focus()
                    for i in range(10):
                        keyboard.send_keys("{DOWN}")
                        time.sleep(0.2)                                # small delay between key presses
                    Actual_Price = Enter_Price.get_value()             # Step 4: Fetch the actual price
                    print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DOWN_ARROW_DECREMENT_CASES(Result_path, TC_NO, Window_Name, Test_Case, "NSE",row_value, DgtsToselect, Result,Expected_Price, Actual_Price)
                case "Decrement 26 times":
                    Enter_Price.set_focus()
                    for i in range(26):
                        keyboard.send_keys("{DOWN}")
                        time.sleep(0.2)  # small delay between key presses
                    Actual_Price = Enter_Price.get_value()  # Step 4: Fetch the actual price
                    print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DOWN_ARROW_DECREMENT_CASES(Result_path, TC_NO, Window_Name, Test_Case, "NSE",row_value, DgtsToselect, Result,Expected_Price, Actual_Price)
            row_value += 1
    def Execute_TestCase_Price_Decrement_CDS(self, Result_path,Window_Name, Test_Case, Window):
        # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
        self.select_Exch = Window.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
        Select_From_DropDown(self.select_Exch, "CDS")
        keyboard.send_keys('{TAB}')
        keyboard.send_keys('{TAB}')
        # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
        self.Custom = Window.child_window(auto_id="ScripSearchBox", class_name="ScripSearchBox")
        self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
        self.Input_TradingSymbol.set_edit_text(" ")
        self.Input_TradingSymbol.set_edit_text("EURUSD")
        self.Input_TradingSymbol.set_focus()
        keyboard.send_keys('{DOWN}')
        keyboard.send_keys('{ENTER}')
        keyboard.send_keys('{TAB}')
        #----------------------------------Get TestExecution data from TESTCASE file-------------------------------------------------#
        TestInfo_sheet = Get_File_Path("TESTCASE.xlsx", 1)
        row_value = 134
        TC_No = 0
        while row_value <= 146:  # Loop until row_value reaches 72
            # Focus Price field of Window
            print("wnd is" + str(Window_Name))
            Select_PriceTextBox = Window.child_window(auto_id="PriceTxt", class_name="UTF_EditBox")
            Enter_Price = Select_PriceTextBox.child_window(auto_id="PART_Editor", class_name="TextBox").set_focus()

            # Get TC_No ,Expected_Price from file "TESTCASE.xlsx"
            TC_NO = TestInfo_sheet.cell(row=row_value, column=1).value
            DgtsToselect = TestInfo_sheet.cell(row=row_value, column=2).value
            Expected_Price = TestInfo_sheet.cell(row=row_value, column=6).value
            print("TC_NO:" + str(TC_No))
            print("DgtsToselect:" + str(DgtsToselect))
            print("Expected_Price:" + str(Expected_Price))

            match DgtsToselect:
                case "ENTER_PRICE":
                    EnteredPrice  = str(Expected_Price)
                    Enter_Price.set_edit_text(str(Expected_Price))
                    Actual_Price = Enter_Price.get_value()
                    print("Actual_Price:", Enter_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.general_log_Decrement(Result_path, EnteredPrice, Test_Case, row_value,Window_Name,"CDS",TC_NO, DgtsToselect, Result,Expected_Price, Actual_Price)
                case "Decrement Once":
                    Enter_Price.set_focus()
                    keyboard.send_keys("{DOWN}")
                    Actual_Price = Enter_Price.get_value()              # Step 4: Fetch the actual price
                    print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DOWN_ARROW_DECREMENT_CASES(Result_path, TC_NO, Window_Name, Test_Case, "CDS",row_value, DgtsToselect, Result,Expected_Price, Actual_Price)
                case "Decrement 10 times":
                    Enter_Price.set_focus()
                    for i in range(10):
                        keyboard.send_keys("{DOWN}")
                        time.sleep(0.2)                                # small delay between key presses
                    Actual_Price = Enter_Price.get_value()             # Step 4: Fetch the actual price
                    print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DOWN_ARROW_DECREMENT_CASES(Result_path, TC_NO, Window_Name, Test_Case, "CDS",row_value, DgtsToselect, Result,Expected_Price, Actual_Price)
                case "Decrement 26 times":
                    Enter_Price.set_focus()
                    for i in range(26):
                        keyboard.send_keys("{DOWN}")
                        time.sleep(0.2)  # small delay between key presses
                    Actual_Price = Enter_Price.get_value()  # Step 4: Fetch the actual price
                    print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_DOWN_ARROW_DECREMENT_CASES(Result_path, TC_NO, Window_Name, Test_Case, "CDS",row_value, DgtsToselect, Result,Expected_Price, Actual_Price)
            row_value += 1

    def Execute_TestCase_Right_arrow_move(self,Result_path,Window_Name, Test_Case, Window):
        TestInfo_sheet = Get_File_Path("TESTCASE.xlsx", 1)
        row_value = 147
        column_value = 1
        TC_No = 0
        while row_value <= 158:  # Loop until row_value reaches 72
            # Focus Price field of Window
            print("wnd is" + str(Window_Name))
            Select_PriceTextBox = Window.child_window(auto_id="PriceTxt", class_name="UTF_EditBox")
            Enter_Price = Select_PriceTextBox.child_window(auto_id="PART_Editor", class_name="TextBox").set_focus()

            # Get TC_No ,DgtsToselect,Input_Price,Position,Expected_Price from file "TESTCASE.xlsx"
            TC_NO = TestInfo_sheet.cell(row=row_value, column=1).value
            DgtsToselect = TestInfo_sheet.cell(row=row_value, column=2).value
            Input_Price = TestInfo_sheet.cell(row=row_value, column=3).value
            Position = TestInfo_sheet.cell(row=row_value, column=4).value
            Expected_Price = TestInfo_sheet.cell(row=row_value, column=5).value
            print("TC_NO:" + str(TC_No))
            print("DgtsToselect:" + str(DgtsToselect))
            print("Input_Price:" + str(Input_Price))
            print("Position:" + str(Position))
            print("Expected_Price:" + str(Expected_Price))

            match DgtsToselect:
                case "ENTER_PRICE":
                    EnteredPrice  = str(Expected_Price)
                    Enter_Price.set_edit_text(str(Expected_Price))
                    Actual_Price = Enter_Price.get_value()
                    print("Actual_Price:", Enter_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.general_log_move_rightarrow_and_Insert(Result_path, EnteredPrice, Test_Case, row_value,Window_Name,TC_NO,DgtsToselect,Result,Input_Price,Position,Expected_Price,Actual_Price)
                case "Insert 1 digit":
                    if str(Position) == "2":
                        keyboard.send_keys('{RIGHT 1}')
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)  # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()  # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}')
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)  # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()  # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "9":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}{RIGHT}')
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)  # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()  # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_RIGHTARROW_AND_INSERT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect, Position, Expected_Price, Actual_Price, Result,row_value)
                case "Insert 2 digit":
                    if str(Position) == "3":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}')
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)  # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                        # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "4":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}')
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)  # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                        # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    elif str(Position) == "5":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}{RIGHT}{RIGHT}')
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)  # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                        # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_RIGHTARROW_AND_INSERT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price,DgtsToselect, Position, Expected_Price, Actual_Price,Result, row_value)
                case "Insert 3 digit":
                    if str(Position) == "3":
                        Enter_Price.type_keys('{HOME}{RIGHT}{RIGHT}')
                        Enter_Price.type_keys(str(Input_Price), set_foreground=True)  # Step 3: Enter the new value
                        Actual_Price = Enter_Price.get_value()                         # Step 4: Fetch the actual price
                        print("Actual_Price:", Actual_Price)
                    Result = self.Check_Result(Expected_Price, Actual_Price)
                    Write_Log.Write_RIGHTARROW_AND_INSERT_CASES(Result_path, TC_NO, Window_Name, Test_Case, Input_Price, DgtsToselect, Position, Expected_Price, Actual_Price,Result, row_value)

            row_value += 1
    def Check_Result(self,Expected_Price,Actual_Price):
        if str(Expected_Price) == str(Actual_Price):
            Result = "PASS"
        else :
            Result = "FAIL"
        return str(Result)























