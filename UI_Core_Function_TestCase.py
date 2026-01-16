from UI_Core_Function_Prerequsite import General_Prerequisite
from UI_Core_Function_TestStep import GeneralTestStep
Prereqisite = General_Prerequisite()
class Testcase():
    def LoginExe(self):                                        #Function contains List of Test Steps to Login EXE
        Prereqisite.Invoke_Exe()                               #Call Function to Invoke .exe
        Prereqisite.Login_Exe()                                #Call Function to Login EXE
        Prereqisite.Login_Exch_Download()                      #Call Function to Download Exchange
        Prereqisite.Focus_On_Exe()                             #Call Function to Focus on Logged in exe
        Prereqisite.Invoke_MarketWatch()                       #call Function to Invoke Market watch Grp and select scrip


    def UI_CORE_FUNCTION_TESTCASES(self,Result_path):
        GeneralTestStep(Result_path)