from UI_Core_Function_TestCase import Testcase
from UI_Core_Function_WriteLog import Log

TC = Testcase()                                                 #Create Object TC for Class TestCase
LG = Log()                                                      #Create Object LG for Class Log



Result_Path =LG.CreateResultFolder("UI_CORE_FUNCTION")               #Function To Create Result Folder which returns Result Path
LG.CreateExcelFile(Result_Path)                                  #Function to create TestResult.xlsx Log in the Result path
#----PREREQUISITE-----#
TC.LoginExe()                                                    #Function to Login Exe
#----TESTCASES-----#
TC.UI_CORE_FUNCTION_TESTCASES(Result_Path)     #Call a function to Execute TestCases




