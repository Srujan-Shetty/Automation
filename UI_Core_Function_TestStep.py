from UI_Core_Function_HelperFunction import Get_File_Path
from UI_Core_Function_Action import PriceCursorAction

AC = PriceCursorAction()
def GeneralTestStep(Result_path):
    TestInfo_sheet = Get_File_Path("TESTCASE.xlsx", 0)
    row_value = 2
    #column_value = 1
    TC_NO = 0
    while row_value <= 13:  # Loop until row_value reaches 72
        TC_NO = TestInfo_sheet.cell(row=row_value, column = 1).value
        Test_Case = TestInfo_sheet.cell(row=row_value, column=2).value
        Window_Name = TestInfo_sheet.cell(row=row_value, column=3).value
        print("Execute Case for "+ Window_Name)
        #Inv_Usng = TestInfo_sheet.cell(row=row_value, column=4).value
        AC.TC_EXECUTION(Test_Case,Window_Name,Result_path)
        row_value+= 1
















