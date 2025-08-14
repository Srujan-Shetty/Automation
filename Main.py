from WriteLog  import Log                                      #Import Class Log from WriteLog File (Class to print ResultLog)
from UI_OrderFlow_TestCase import TestCase                     #Import Class TestCase from UI_OrderFlow_TestCase File ( Class which has List Of Testcase )


TC = TestCase()                                                 #Create Object TC for Class TestCase
LG = Log()                                                      #Create Object LG for Class Log
Result_Path =LG.CreateResultFolder("GENERAL_UI_TESTING_NSE")    #Function To Create Result Folder which returns Result Path
LG.CreateLogTextFile(Result_Path)                               #Function To create Result Log in the Result path
LG.CreateExcelFile(Result_Path)                                 #Function to create TestResult.xlsx Log in the Result path
#----PREREQUISITE-----#
# TC.PuttyPublish()
TC.LoginExe()                                                   #Function to Login Exe
#----TESTCASES-----#
TC.ORDER_FLOW_TESTCASES(Result_Path)                            #Call a function to Execute TestCases








