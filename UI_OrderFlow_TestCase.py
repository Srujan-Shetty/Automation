from UI_OrderFlow_TestStep import General_Prerequisite ,GeneralTestStep        #Import Class General_Prerequisite,GeneralTestStep from File UI_OrderFlow_TestStep

Prereqisite = General_Prerequisite()
TestStep = GeneralTestStep()
class TestCase():                                              #Class contains TestCases
    #----------------------------------------Publish Backend values in Putty----------------------------------------------#
    def PuttyPublish(self):
        Prereqisite.Invoke_PuttyExe()

    #----------------------------------------Login EXE--------------------------------------------------------------------#
    def LoginExe(self):                                        #Function contains List of Test Steps to Login EXE
        Prereqisite.Invoke_Exe()                               #Call Function to Invoke .exe
        Prereqisite.Login_Exe()                                #Call Function to Login EXE
        Prereqisite.Login_Exch_Download()                      #Call Function to Download Exchange
        Prereqisite.Focus_On_Exe()                             #call Function to Focus on Logged in exe

    # ----------------------------------------Add Scrips to Market Watch--------------------------------------------------#
    def AddScripsToMW(self):                                   #Function contains List of Test Step to Add Scrip to Market Watch
        Prereqisite.Focus_On_Exe()                             #Call Function to Focus on Logged in exe
        Prereqisite.Invoke_MarketWatch()                       #Call Function to Invoke Market Watch

    # ----------------------------------------Complete Order Flow TestCase -----------------------------------------------#
    def ORDER_FLOW_TESTCASES(self,Result_path):                  #Function to execute all Testcase
        # TestStep.LMT_Order_TestCases(Result_path)                #Call Function to execute TestCases for LMT Order
        TestStep.SLLMT_Order_TestCases(Result_path)              #Call Function to execute TestCases for SL-LMT Order
        # TestStep.MKT_Order_TestCases(Result_path)                # Call Function to execute TestCases for MKT Order
        # TestStep.SLMKT_Order_TestCases(Result_path)              # Call Function to execute TestCases for MKT Order
