from UI_Core_Function_HelperFunction import Get_File_Path
class LoginWindow_TestData:
    def Get_LoginExe_path(self):
        TestInfo_sheet = Get_File_Path("TESTDATA.xlsx" ,0 )     #Method to Activate TestCaseList.xlsx sheet
        Exe_Path = TestInfo_sheet.cell(row=3, column=2).value
        return Exe_Path
    def Get_Login_UserName(self):
        TestInfo_sheet = Get_File_Path("TESTDATA.xlsx", 0)      #Method to Activate TestCaseList.xlsx sheet
        USERNAME = TestInfo_sheet.cell(row=4, column=2).value
        return USERNAME
    def Get_Login_Password(self):
        TestInfo_sheet = Get_File_Path("TESTDATA.xlsx", 0)       #Method to Activate TestCaseList.xlsx sheet
        PASSWORD = TestInfo_sheet.cell(row=5, column=2).value
        return PASSWORD