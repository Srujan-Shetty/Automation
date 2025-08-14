from pywinauto import Application
from pywinauto import keyboard
import time
from UI_OrderFlow_TestData import LoginWindow_TestData
import pyperclip
import math
from WriteLog import Log
from UI_Order_Flow_helper import Select_From_DropDown ,Get_File_Path

#comment
LoginTestData = LoginWindow_TestData()
LG = Log()
class OrderPlacements():
    def PlaceLMTOrder(self,Result_path,TC_NO,row_value,Test_Scenario,Test_Case,Trans_Type,Exchange,Order_Type,Order_Duration,Trading_Symbol,Input_Quantity,Input_Price,Input_TriggerPrice,Input_DiscQty,Quantity,Price,TriggerPrice,DiscQty,Product,Account_ID,Remark,OpenQty,TradedQty,AvgPrice,TypeOfOrder):
        match Test_Case:
            case "Place":
                    if Trans_Type == "BUY":
                        self.Focus_On_Exe()
                        keyboard.send_keys('{F1}')
                        self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                        #-----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                        self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo", control_type="ComboBox").wrapper_object()
                        Select_From_DropDown(self.select_Exch , str(Exchange))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType", class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.select_OrderType, str(Order_Type))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                        self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur", class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.select_Duration, str(Order_Duration))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                        self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                        self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch",class_name="TextEdit")
                        self.Input_TradingSymbol.set_edit_text(" ")
                        self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                        self.Input_TradingSymbol.set_focus()
                        keyboard.send_keys('{DOWN}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                        self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt" ,class_name="UTF_EditBox")
                        self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Remark = time.time()
                        self.Enter_Remark.set_edit_text(self.Remark)
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                        self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Enter_Qty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                        self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt", class_name="UTF_EditBox")
                        self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Enter_Price.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                        self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                        self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                        Select_From_DropDown(self.Select_Product, str(Product))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                        self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                        self.AccountSearch.click_input()
                        self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.AccountSearch.click_input()
                        self.Account_EditBox.set_edit_text('')
                        self.Account_EditBox.set_edit_text(str(Account_ID))
                        self.Account_EditBox.exists(timeout=10)
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                        self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                        keyboard.send_keys('{ESC}')
                        # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                        self.Get_Place_OrdDetails(Result_path,str(self.Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO,row_value,Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)                        #Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                    else:
                        self.Focus_On_Exe()
                        keyboard.send_keys('{F2}')
                        keyboard.send_keys('{F2}')
                        self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                        # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                        self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo", control_type="ComboBox").wrapper_object()
                        Select_From_DropDown(self.select_Exch, str(Exchange))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType", class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.select_OrderType, str(Order_Type))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                        self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.select_Duration, str(Order_Duration))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                        self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                        self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                        self.Input_TradingSymbol.set_edit_text(" ")
                        self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                        self.Input_TradingSymbol.set_focus()
                        keyboard.send_keys('{DOWN}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                        self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                        self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Remark = time.time()
                        self.Enter_Remark.set_edit_text(self.Remark)
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                        self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Enter_Qty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                        self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Enter_Price.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                        self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                        self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd", class_name="ComboBox")
                        Select_From_DropDown(self.Select_Product, str(Product))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                        self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                        self.AccountSearch.click_input()
                        self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.AccountSearch.click_input()
                        self.Account_EditBox.set_edit_text('')
                        self.Account_EditBox.set_edit_text(str(Account_ID))
                        self.Account_EditBox.exists(timeout=10)
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Click On Sell Button--------------------------------------------------#
                        self.Click_BuyButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                        keyboard.send_keys('{ESC}')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Place_OrdDetails(Result_path,str(self.Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)                            #Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                    return str(self.Remark)
            case "Modify":
                keyboard.send_keys('{F3}')                                              #Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook" ,auto_id="OrderBook" ,class_name="OrderBook")
                self.OrderBook.set_focus()
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                self.FilterRow_dataGrid = self.PendingOrderGrp.child_window(auto_id="uiOrderBookOpen")
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                # self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
                self.PendingOrderGrp.type_keys(str(Remark))

                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":                                                #If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')                                       #Select Order
                    keyboard.send_keys('+{F2}')                                        #Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").set_focus()
                    keyboard.send_keys('{TAB}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    print("OrderType"+str(Order_Type))
                    self.Get_Modify_OrdDetails(Result_path,str(Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)                    #Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                else:

                    if Trans_Type == "SELL":                                                  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')                                          # Select Order
                        keyboard.send_keys('+{F2}')                                           # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin", class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").set_focus()
                        keyboard.send_keys('{TAB}')
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt", class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Modify_OrdDetails(Result_path,str(self.Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)
            case "Cancel":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook", class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.Click_Remark_Header = self.Header.child_window(title="OrdRemarks", framework_id="WPF")
                self.FilterRow = self.PendingOrderGrp.child_window(title="AutoFilterRow", auto_id="AutoFilterRow",class_name="AutoFilterRowControl")
                # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
                self.Filter_Remark = self.FilterRow.child_window(auto_id="OrdRemarks",class_name="FilterCellContentPresenter")
                self.Filter_Remark.click_input()
                self.Filter_Remark.type_keys(str(Remark))
                keyboard.send_keys('{DOWN}')
                keyboard.send_keys('{DELETE}')
                keyboard.send_keys('{ENTER}')
                keyboard.send_keys('^F')
                # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                self.Get_Cancel_OrdDetails(Result_path,str(self.Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)

            case "PlaceRedReject":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch,str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Order_Type))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product,str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_PlaceReject_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty,OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario, Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Order_Type))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_SellButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_PlaceReject_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty,OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario, Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(self.Remark)

            case "ModifyRedReject":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.PendingOrderGrp.type_keys(str(Remark))
                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").set_focus()
                    keyboard.send_keys('{TAB}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    time.sleep(1)
                    #keyboard.send_keys('{TAB}')
                    #time.sleep(1)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_Modify_OrdDetails(Result_path,str(Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").set_focus()
                        keyboard.send_keys('{TAB}')
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        #keyboard.send_keys('{TAB}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Modify_OrdDetails(Result_path,str(Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(Remark)
            case "ModifyToCompleteTrade":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.PendingOrderGrp.type_keys(str(Remark))
                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").set_focus()
                    keyboard.send_keys('{TAB}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    time.sleep(1)
                    # keyboard.send_keys('{TAB}')
                    # time.sleep(1)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty,OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").set_focus()
                        keyboard.send_keys('{TAB}')
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        # keyboard.send_keys('{TAB}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty,OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(Remark)
            case "ModifyToPartialTrade":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.PendingOrderGrp.type_keys(str(Remark))
                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").set_focus()
                    keyboard.send_keys('{TAB}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    time.sleep(1)
                    # keyboard.send_keys('{TAB}')
                    # time.sleep(1)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_PartialTrade_OrdDetails(Result_path,str(self.Remark), Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").set_focus()
                        keyboard.send_keys('{TAB}')
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        # keyboard.send_keys('{TAB}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_PartialTrade_OrdDetails(Result_path,str(self.Remark), Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)
            case "PartialTrade":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Order_Type))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_PartialTrade_OrdDetails(Result_path,str(self.Remark), Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Order_Type))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_PartialTrade_OrdDetails(Result_path,str(self.Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(self.Remark)
            case "CompleteTrade":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Order_Type))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty,OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Order_Type))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty,OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,Order_Type,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(self.Remark)

            case _:
                return "NA"

    def PlaceSLOrder(self,Result_path,TC_NO,row_value,Test_Scenario,Test_Case,Trans_Type,Exchange,Input_OrderType,Order_Duration,Trading_Symbol,Input_Quantity,Input_Price,Input_TriggerPrice,Input_DiscQty,OrderType,Quantity,Price,TriggerPrice,DiscQty,Product,Account_ID,Remark,OpenQty,TradedQty,AvgPrice,TypeOfOrder):
        match Test_Case:
            case "Place":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    #print("OrderType: "+ str(OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter TRIGGER PRICE in  Order Entry------------------------------------------------------#
                    self.Select_TriggerPriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_TriggerPriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_Place_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty, OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter TRIGGER PRICE in  Order Entry------------------------------------------------------#
                    self.Select_TriggerPriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_TriggerPriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_Place_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty ,OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(self.Remark)
            case "Modify":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                self.FilterRow_dataGrid = self.PendingOrderGrp.child_window(auto_id="uiOrderBookOpen")
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                # self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
                self.PendingOrderGrp.type_keys(str(Remark))

                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter TriggerPrice in  Modify Order Entry---------------------------------------------------#
                    self.Select_TriggerPrice = self.Modify_Buy_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_Modify_OrdDetails(Result_path, str(Remark), Quantity, Price, TriggerPrice,DiscQty,OpenQty, TradedQty, AvgPrice,TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                else:

                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_TriggerPrice = self.Modify_Sell_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Modify_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty, OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)
            case "Cancel":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.Click_Remark_Header = self.Header.child_window(title="OrdRemarks", framework_id="WPF")
                self.FilterRow = self.PendingOrderGrp.child_window(title="AutoFilterRow", auto_id="AutoFilterRow",class_name="AutoFilterRowControl")
                # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
                self.Filter_Remark = self.FilterRow.child_window(auto_id="OrdRemarks",class_name="FilterCellContentPresenter")
                self.Filter_Remark.click_input()
                self.PendingOrderGrp.type_keys(str(Remark))
                keyboard.send_keys('{DOWN}')
                keyboard.send_keys('{DELETE}')
                keyboard.send_keys('{ENTER}')
                keyboard.send_keys('^F')
                # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                self.Get_Cancel_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice,TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
            case "PlaceRedReject":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_TriggerPrice = self.Normal_Buy_Order_Entry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    self.Enter_TriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_TriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_PlaceReject_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice, TC_NO, row_value,Test_Scenario, Trans_Type, Input_OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_TriggerPrice = self.Modify_Sell_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_SellButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_PlaceReject_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice, TC_NO, row_value,Test_Scenario, Trans_Type, Input_OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(self.Remark)
            case "ModifyRedReject":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.PendingOrderGrp.type_keys(str(Remark))
                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter TriggerPrice in  Modify Order Entry---------------------------------------------------#
                    self.Select_TriggerPrice = self.Modify_Buy_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_Modify_OrdDetails(Result_path, str(Remark), Quantity, Price,TriggerPrice,DiscQty,OpenQty, TradedQty, AvgPrice,TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                else:
                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_TriggerPrice = self.Modify_Sell_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Modify_OrdDetails(Result_path, str(Remark), Quantity, Price,TriggerPrice ,DiscQty,OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(Remark)
            case "TriggerOrderModify":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                self.FilterRow_dataGrid = self.PendingOrderGrp.child_window(auto_id="uiOrderBookOpen")
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                # self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
                self.PendingOrderGrp.type_keys(str(Remark))

                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter TriggerPrice in  Modify Order Entry---------------------------------------------------#
                    #self.Select_TriggerPrice = self.Modify_Buy_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    #self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                    #self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_Modify_OrdDetails(Result_path, str(Remark), Quantity, Price, TriggerPrice, DiscQty,OpenQty, TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type, OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                else:

                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                        #self.Select_TriggerPrice = self.Modify_Sell_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                        #self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                        #self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                        #keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Modify_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice, TC_NO, row_value,Test_Scenario, Trans_Type, OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)
            case "TriggerOrderModifyRedReject":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.PendingOrderGrp.type_keys(str(Remark))
                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    # -----------------------------------------Enter TriggerPrice in  Modify Order Entry---------------------------------------------------#
                    #self.Select_TriggerPrice = self.Modify_Buy_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    #self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                    #self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_Modify_OrdDetails(Result_path, str(Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario, Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                        #self.Select_TriggerPrice = self.Modify_Sell_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                        #self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                        #self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Modify_OrdDetails(Result_path, str(Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario, Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)
            case "TriggerPrderModifyToCompleteTrade":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                self.FilterRow_dataGrid = self.PendingOrderGrp.child_window(auto_id="uiOrderBookOpen")
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                # self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
                self.PendingOrderGrp.type_keys(str(Remark))

                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter TriggerPrice in  Modify Order Entry---------------------------------------------------#
                    # self.Select_TriggerPrice = self.Modify_Buy_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    # self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                    # self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    # keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty, OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                else:
                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                        # self.Select_TriggerPrice = self.Modify_Sell_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                        # self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                        # self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                        # keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty, OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)
            case "PartialTrade":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_PartialTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price,DiscQty, OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')

                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_PartialTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price, DiscQty,OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(self.Remark)
            case "CompleteTrade":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price,DiscQty, OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_SellButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price,DiscQty, OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(self.Remark)
            case "Create_LTP":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')


                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')

                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_SellButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                return str(self.Remark)
            case _:
                return "NA"
    def PlaceMKTOrder(self,Result_path,TC_NO,row_value,Test_Scenario,Test_Case,Trans_Type,Exchange,Input_OrderType,Order_Duration,Trading_Symbol,Input_Quantity,Input_Price,Input_TriggerPrice,Input_DiscQty,Quantity,Price,TriggerPrice,DiscQty,Product,Account_ID,Remark,OpenQty,TradedQty,AvgPrice,OrderType,TypeOfOrder):
        match Test_Case:
            case "Place":
                    print(str(Product))
                    if Trans_Type == "BUY":
                        self.Focus_On_Exe()
                        keyboard.send_keys('{F1}')
                        self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                        #-----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                        self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo", control_type="ComboBox").wrapper_object()
                        Select_From_DropDown(self.select_Exch, str(Exchange))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType", class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                        self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur", class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.select_Duration, str(Order_Duration))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                        self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                        self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch",class_name="TextEdit")
                        self.Input_TradingSymbol.set_edit_text(" ")
                        self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                        self.Input_TradingSymbol.set_focus()
                        keyboard.send_keys('{DOWN}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                        self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt" ,class_name="UTF_EditBox")
                        self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Remark = time.time()
                        self.Enter_Remark.set_edit_text(self.Remark)
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                        self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Enter_Qty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                        #self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt", class_name="UTF_EditBox")
                        #self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        #self.Enter_Price.set_edit_text(str(Input_Price))
                        #keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                        self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                        self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                        Select_From_DropDown(self.Select_Product, str(Product))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                        self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                        self.AccountSearch.click_input()
                        self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.AccountSearch.click_input()
                        self.Account_EditBox.set_edit_text('')
                        self.Account_EditBox.set_edit_text(str(Account_ID))
                        self.Account_EditBox.exists(timeout=10)
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                        self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                        keyboard.send_keys('{ESC}')
                        # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                        self.Get_Place_OrdDetails(Result_path,str(self.Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO,row_value,Test_Scenario,Trans_Type,OrderType,TypeOfOrder)                        #Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                    else:
                        self.Focus_On_Exe()
                        keyboard.send_keys('{F2}')
                        self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                        # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                        self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo", control_type="ComboBox").wrapper_object()
                        Select_From_DropDown(self.select_Exch, str(Exchange))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType", class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                        self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.select_Duration, str(Order_Duration))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                        self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                        self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                        self.Input_TradingSymbol.set_edit_text(" ")
                        self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                        self.Input_TradingSymbol.set_focus()
                        keyboard.send_keys('{DOWN}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                        self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                        self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Remark = time.time()
                        self.Enter_Remark.set_edit_text(self.Remark)
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                        self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Enter_Qty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                        #self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        #self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        #self.Enter_Price.set_edit_text(str(Input_Price))
                        #keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                        self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                        self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                        self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd", class_name="ComboBox")
                        Select_From_DropDown(self.Select_Product, str(Product))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                        self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                        self.AccountSearch.click_input()
                        self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.AccountSearch.click_input()
                        self.Account_EditBox.set_edit_text('')
                        self.Account_EditBox.set_edit_text(str(Account_ID))
                        self.Account_EditBox.exists(timeout=10)
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Click On Sell Button--------------------------------------------------#
                        self.Click_BuyButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                        keyboard.send_keys('{ESC}')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Place_OrdDetails(Result_path,str(self.Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,OrderType,TypeOfOrder)                            #Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                    return str(self.Remark)
            case "Modify":
                keyboard.send_keys('{F3}')                                              #Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook" ,auto_id="OrderBook" ,class_name="OrderBook")
                self.OrderBook.set_focus()
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                self.FilterRow_dataGrid = self.PendingOrderGrp.child_window(auto_id="uiOrderBookOpen")
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                # self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
                self.PendingOrderGrp.type_keys(str(Remark))

                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":                                                #If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')                                       #Select Order
                    keyboard.send_keys('+{F2}')                                        #Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    #self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                    #self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{ENTER}')
                    #keyboard.send_keys('^F')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    #print("OrderType"+str(Order_Type))
                    self.Get_Modify_OrdDetails(Result_path,str(Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,OrderType,TypeOfOrder)                    #Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                else:

                    if Trans_Type == "SELL":                                                  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')                                          # Select Order
                        keyboard.send_keys('+{F2}')                                           # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin", class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt", class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        #keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        #self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        #self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        #self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        #keyboard.send_keys('{ENTER}')
                        #keyboard.send_keys('^F')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Modify_OrdDetails(Result_path,str(self.Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)
            case "Cancel":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook", class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.Click_Remark_Header = self.Header.child_window(title="OrdRemarks", framework_id="WPF")
                self.FilterRow = self.PendingOrderGrp.child_window(title="AutoFilterRow", auto_id="AutoFilterRow",class_name="AutoFilterRowControl")
                # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
                self.Filter_Remark = self.FilterRow.child_window(auto_id="OrdRemarks",class_name="FilterCellContentPresenter")
                self.Filter_Remark.click_input()
                self.Filter_Remark.type_keys(str(Remark))
                keyboard.send_keys('{DOWN}')
                keyboard.send_keys('{DELETE}')
                keyboard.send_keys('{ENTER}')
                keyboard.send_keys('^F')
                # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                self.Get_Cancel_OrdDetails(Result_path,str(self.Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
            case "PlaceRedReject":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    #self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    #self.Enter_Price.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_PlaceReject_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice, TC_NO, row_value,Test_Scenario, Trans_Type, Input_OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    #self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    #self.Enter_Price.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_SellButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_PlaceReject_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice, TC_NO, row_value,Test_Scenario, Trans_Type, Input_OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(self.Remark)

            case "ModifyRedReject":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.PendingOrderGrp.type_keys(str(Remark))
                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    #self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    #self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{ENTER}')
                    #time.sleep(1)
                    #keyboard.send_keys('{TAB}')
                    #time.sleep(1)
                    #keyboard.send_keys('{ENTER}')
                    #keyboard.send_keys('^F')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_Modify_OrdDetails(Result_path,str(Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        #keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        #self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        #self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        #self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        #keyboard.send_keys('{ENTER}')
                        #keyboard.send_keys('{TAB}')
                        #keyboard.send_keys('{ENTER}')
                        #keyboard.send_keys('^F')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Modify_OrdDetails(Result_path,str(Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(Remark)
            case "ModifyToPartialTrade":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.PendingOrderGrp.type_keys(str(Remark))
                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    #self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    #self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    time.sleep(1)
                    # keyboard.send_keys('{TAB}')
                    # time.sleep(1)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_PartialTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice, TC_NO, row_value,Test_Scenario, Trans_Type, OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        #self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        #self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        #self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        #keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        # keyboard.send_keys('{TAB}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_PartialTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice, TC_NO, row_value,Test_Scenario, Trans_Type, OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)
            case "ModifyToCompleteTrade":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.PendingOrderGrp.type_keys(str(Remark))
                print(str(Remark))
                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    #self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    #self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    time.sleep(1)
                    # keyboard.send_keys('{TAB}')
                    # time.sleep(1)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice, TC_NO, row_value,Test_Scenario, Trans_Type, OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        #self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        #self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        #self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        #keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        # keyboard.send_keys('{TAB}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice, TC_NO, row_value,Test_Scenario, Trans_Type, OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

            case "PartialTrade":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    #self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    #self.Enter_Price.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_PartialTrade_OrdDetails(Result_path,str(self.Remark), Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    #self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    #self.Enter_Price.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_PartialTrade_OrdDetails(Result_path,str(self.Remark),Quantity,Price,TriggerPrice,DiscQty,OpenQty,TradedQty,AvgPrice,TC_NO, row_value,Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(self.Remark)
            case "CompleteTrade":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    #self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    #self.Enter_Price.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty,OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    #self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    #self.Enter_Price.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty,OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(self.Remark)

            case "Create_LTP":
                Remark_1 = str(Remark)
                print("Updated Remark: "+str(Remark_1))
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    #self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    #self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    #self.Remark = time.time()
                    #self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')

                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    #self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    #self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    #self.Remark = time.time()
                    #self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_SellButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')

                return str(Remark_1)




            case _:
                return "NA"
    def PlaceSLMKTOrder(self,Result_path,TC_NO,row_value,Test_Scenario,Test_Case,Trans_Type,Exchange,Input_OrderType,Order_Duration,Trading_Symbol,Input_Quantity,Input_Price,Input_TriggerPrice,Input_DiscQty,OrderType,Quantity,Price,TriggerPrice,DiscQty,Product,Account_ID,Remark,OpenQty,TradedQty,AvgPrice,TypeOfOrder):
        match Test_Case:
            case "Place":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    #self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    #self.Enter_Price.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter TRIGGER PRICE in  Order Entry------------------------------------------------------#
                    self.Select_TriggerPriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_TriggerPriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    print(str(Product))
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_Place_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty, OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    #self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    #self.Enter_Price.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter TRIGGER PRICE in  Order Entry------------------------------------------------------#
                    self.Select_TriggerPriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_TriggerPriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_Place_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty ,OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(self.Remark)
            case "Modify":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                self.FilterRow_dataGrid = self.PendingOrderGrp.child_window(auto_id="uiOrderBookOpen")
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                # self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
                self.PendingOrderGrp.type_keys(str(Remark))

                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    #self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    #self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter TriggerPrice in  Modify Order Entry---------------------------------------------------#
                    self.Select_TriggerPrice = self.Modify_Buy_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_Modify_OrdDetails(Result_path, str(Remark), Quantity, Price, TriggerPrice,DiscQty,OpenQty, TradedQty, AvgPrice,TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                else:

                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        #self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        #self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        #self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        #keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_TriggerPrice = self.Modify_Sell_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Modify_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty, OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)
            case "Cancel":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.Click_Remark_Header = self.Header.child_window(title="OrdRemarks", framework_id="WPF")
                self.FilterRow = self.PendingOrderGrp.child_window(title="AutoFilterRow", auto_id="AutoFilterRow",class_name="AutoFilterRowControl")
                # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
                self.Filter_Remark = self.FilterRow.child_window(auto_id="OrdRemarks",class_name="FilterCellContentPresenter")
                self.Filter_Remark.click_input()
                self.Filter_Remark.type_keys(str(Remark))
                keyboard.send_keys('{DOWN}')
                keyboard.send_keys('{DELETE}')
                keyboard.send_keys('{ENTER}')
                keyboard.send_keys('^F')
                # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                self.Get_Cancel_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice,TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
            case "PlaceRedReject":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    #self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    #self.Enter_Price.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_TriggerPrice = self.Normal_Buy_Order_Entry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    self.Enter_TriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_TriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_PlaceReject_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice, TC_NO, row_value,Test_Scenario, Trans_Type, Input_OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    #self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    #self.Enter_Price.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_TriggerPrice = self.Normal_Sell_Order_Entry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    self.Enter_TriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_TriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_SellButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_PlaceReject_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice, TC_NO, row_value,Test_Scenario, Trans_Type, Input_OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(self.Remark)
            case "ModifyRedReject":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.PendingOrderGrp.type_keys(str(Remark))
                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    #self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    #self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    #self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter TriggerPrice in  Modify Order Entry---------------------------------------------------#
                    self.Select_TriggerPrice = self.Modify_Buy_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_Modify_OrdDetails(Result_path, str(Remark), Quantity, Price,TriggerPrice,DiscQty,OpenQty, TradedQty, AvgPrice,TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                else:
                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        #self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        #self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        #self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        #keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_TriggerPrice = self.Modify_Sell_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Modify_OrdDetails(Result_path, str(Remark), Quantity, Price,TriggerPrice ,DiscQty,OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                return str(Remark)
            case "TriggerOrderModify":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                self.FilterRow_dataGrid = self.PendingOrderGrp.child_window(auto_id="uiOrderBookOpen")
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                # self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
                self.PendingOrderGrp.type_keys(str(Remark))

                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    # Quantity = TestInfo_sheet.cell(row=row_value, column=8).value      #Get Modify Qty value
                    # Price = TestInfo_sheet.cell(row=row_value, column=9).value         #Get Modify price Value
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter TriggerPrice in  Modify Order Entry---------------------------------------------------#
                    #self.Select_TriggerPrice = self.Modify_Buy_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    #self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                    #self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    #keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_Modify_OrdDetails(Result_path, str(Remark), Quantity, Price, TriggerPrice, DiscQty,OpenQty, TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type, OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                else:

                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                        #self.Select_TriggerPrice = self.Modify_Sell_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                        #self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                        #self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                        #keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Modify_OrdDetails(Result_path, str(self.Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty, TradedQty, AvgPrice, TC_NO, row_value,Test_Scenario, Trans_Type, OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)
            case "TriggerOrderModifyRedReject":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.PendingOrderGrp.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                self.PendingOrderGrp.type_keys(str(Remark))
                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    # -----------------------------------------Enter TriggerPrice in  Modify Order Entry---------------------------------------------------#
                    #self.Select_TriggerPrice = self.Modify_Buy_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    #self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                    #self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_Modify_OrdDetails(Result_path, str(Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario, Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                        #self.Select_TriggerPrice = self.Modify_Sell_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                        #self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                        #self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_Modify_OrdDetails(Result_path, str(Remark), Quantity, Price, TriggerPrice,DiscQty, OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario, Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)
            case "TriggerPrderModifyToCompleteTrade":
                keyboard.send_keys('{F3}')  # Invoke Order Book
                self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook",class_name="OrderBook")
                self.OrderBook.set_focus()
                self.Clear_filter = self.OrderBook.child_window(title="Clear Filters",class_name="Button").click_input()
                self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
                self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders",class_name="OpenOrdBook")
                self.PendingOrderGrp.click_input()
                self.FilterRow_dataGrid = self.PendingOrderGrp.child_window(auto_id="uiOrderBookOpen")
                # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
                self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
                # self.Header.click_input()
                # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
                keyboard.send_keys('^F')
                # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
                self.PendingOrderGrp.type_keys(str(Remark))

                # -----------------------------------------------------Select Order and invoke Modify Order Entry -------------------------------------#
                if Trans_Type == "BUY":  # If Trans Type is BUY ,Invoke Buy Modify Order Entry
                    keyboard.send_keys('{DOWN}')  # Select Order
                    keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                    self.Modify_Buy_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                    TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                    self.Modify_Buy_OrderEntry.set_focus()
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.Modify_OrderType = self.Modify_Buy_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_Qty = self.Modify_Buy_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                    self.Select_Price = self.Modify_Buy_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter TriggerPrice in  Modify Order Entry---------------------------------------------------#
                    # self.Select_TriggerPrice = self.Modify_Buy_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                    # self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                    # self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                    # keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                    self.Select_DiscQty = self.Modify_Buy_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                    self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('^F')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty, OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                else:
                    if Trans_Type == "SELL":  # If Trans Type is SELL ,Invoke SELL Modify Order Entry
                        keyboard.send_keys('{DOWN}')  # Select Order
                        keyboard.send_keys('+{F2}')  # Invoke Modify Order Entry
                        self.Modify_Sell_OrderEntry = self.App_MainWnd.child_window(auto_id="OrderEntrywin",class_name="Window")
                        TestInfo_sheet = Get_File_Path("TestData.xlsx", 1)  # Method to Active the TestData.xlsx sheet
                        self.Modify_Sell_OrderEntry.set_focus()
                        # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                        self.Modify_OrderType = self.Modify_Sell_OrderEntry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                        Select_From_DropDown(self.Modify_OrderType, str(Input_OrderType))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_Qty = self.Modify_Sell_OrderEntry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyQty = self.Select_Qty.child_window(auto_id="PART_Editor", class_name="TextBox")
                        self.Enter_ModifyQty.set_edit_text(str(Input_Quantity))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Price in  Modify Order Entry---------------------------------------------------#
                        self.Select_Price = self.Modify_Sell_OrderEntry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyPrice = self.Select_Price.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyPrice.set_edit_text(str(Input_Price))
                        keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Trigger Price in  Modify Order Entry---------------------------------------------------#
                        # self.Select_TriggerPrice = self.Modify_Sell_OrderEntry.child_window(auto_id="TriggerTxt",class_name="UTF_EditBox")
                        # self.Enter_ModifyTriggerPrice = self.Select_TriggerPrice.child_window(auto_id="PART_Editor",class_name="TextBox")
                        # self.Enter_ModifyTriggerPrice.set_edit_text(str(Input_TriggerPrice))
                        # keyboard.send_keys('{TAB}')
                        # -----------------------------------------Enter Disc Qty in  Modify Order Entry---------------------------------------------------#
                        self.Select_DiscQty = self.Modify_Sell_OrderEntry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                        self.Enter_ModifyDiscQty = self.Select_DiscQty.child_window(auto_id="PART_Editor",class_name="TextBox")
                        self.Enter_ModifyDiscQty.set_edit_text(str(Input_DiscQty))
                        keyboard.send_keys('{ENTER}')
                        keyboard.send_keys('^F')
                        # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                        self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price,TriggerPrice,DiscQty, OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(Remark)
            case "PartialTrade":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_PartialTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price,DiscQty, OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_PartialTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price, DiscQty,OpenQty, TradedQty,AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(self.Remark)
            case "CompleteTrade":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Number ------------------------------------------------------------------#
                    self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price,DiscQty, OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file

                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_SellButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                    # -----------------------------------------Get Order Details ------------------------------------------------------------------#
                    self.Get_CompleteTrade_OrdDetails(Result_path, str(self.Remark), Quantity, Price,DiscQty, OpenQty,TradedQty, AvgPrice, TC_NO, row_value, Test_Scenario,Trans_Type,OrderType,TypeOfOrder)  # Function to invoke Putty and Get ORD-UPD from Saf.SAF.cjd file
                return str(self.Remark)
            case "Create_LTP":
                if Trans_Type == "BUY":
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F1}')
                    self.Normal_Buy_Order_Entry = self.App_MainWnd.child_window(title="Buy Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Buy_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Buy_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Buy_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Buy_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Buy_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Buy_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Buy_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Buy Button----------------------------------------------------------------#
                    self.Click_BuyButton = self.Normal_Buy_Order_Entry.child_window(title="Buy",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')


                else:
                    self.Focus_On_Exe()
                    keyboard.send_keys('{F2}')
                    self.Normal_Sell_Order_Entry = self.App_MainWnd.child_window(title="Sell Order Entry",auto_id="OrderEntrywin",class_name="Window")
                    # -----------------------------------------Select EXCAHANGE from Order Entry-------------------------------------------------#
                    self.select_Exch = self.Normal_Sell_Order_Entry.child_window(auto_id="ExchangeCombo",control_type="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Exch, str(Exchange))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER TYPE from Order Entry------------------------------------------------#
                    self.select_OrderType = self.Normal_Sell_Order_Entry.child_window(auto_id="ComboOrdType",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_OrderType, str(Input_OrderType))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select ORDER DURATION from Order Entry---------------------------------------------#
                    self.select_Duration = self.Normal_Sell_Order_Entry.child_window(auto_id="comboOrdDur",class_name="ComboBox").wrapper_object()
                    Select_From_DropDown(self.select_Duration, str(Order_Duration))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select TRADING SYMBOL from Order Entry---------------------------------------------#
                    self.Custom = self.Normal_Sell_Order_Entry.child_window(auto_id="ScripSearchBox",class_name="ScripSearchBox")
                    self.Input_TradingSymbol = self.Custom.child_window(auto_id="editSearch", class_name="TextEdit")
                    self.Input_TradingSymbol.set_edit_text(" ")
                    self.Input_TradingSymbol.set_edit_text(str(Trading_Symbol))
                    self.Input_TradingSymbol.set_focus()
                    keyboard.send_keys('{DOWN}')
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')

                    # -----------------------------------------Enter Remark in  Order Entry------------------------------------------------------#
                    self.Remark_TextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="RemarksTxt",class_name="UTF_EditBox")
                    self.Enter_Remark = self.Remark_TextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Remark = time.time()
                    self.Enter_Remark.set_edit_text(self.Remark)
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter Qty in  Order Entry--------------------------------------------------------#
                    self.Select_QtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="QtyTxt",class_name="UTF_EditBox")
                    self.Enter_Qty = self.Select_QtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Qty.set_edit_text(str(Input_Quantity))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter PRICE in  Order Entry------------------------------------------------------#
                    self.Select_PriceTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="PriceTxt",class_name="UTF_EditBox")
                    self.Enter_Price = self.Select_PriceTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_Price.set_edit_text(str(Input_Price))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Enter DiscQty in  Order Entry------------------------------------------------------#
                    self.Select_DiscQtyTextBox = self.Normal_Sell_Order_Entry.child_window(auto_id="DiscQtyTxt",class_name="UTF_EditBox")
                    self.Enter_DiscQty = self.Select_DiscQtyTextBox.child_window(auto_id="PART_Editor",class_name="TextBox").set_focus()
                    self.Enter_DiscQty.set_edit_text(str(Input_DiscQty))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select PRODUCT from  Order Entry--------------------------------------------------#
                    self.Select_Product = self.Normal_Sell_Order_Entry.child_window(auto_id="comboProd",class_name="ComboBox")
                    Select_From_DropDown(self.Select_Product, str(Product))
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Select Account ID from  Order Entry--------------------------------------------------#
                    self.AccountSearch = self.Normal_Sell_Order_Entry.child_window(auto_id="ClientSearchBox",class_name="ClientSearchBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox = self.AccountSearch.child_window(auto_id="PART_Editor", class_name="TextBox")
                    self.AccountSearch.click_input()
                    self.Account_EditBox.set_edit_text('')
                    self.Account_EditBox.set_edit_text(str(Account_ID))
                    self.Account_EditBox.exists(timeout=10)
                    keyboard.send_keys('{ENTER}')
                    keyboard.send_keys('{TAB}')
                    # -----------------------------------------Click On Sell Button--------------------------------------------------#
                    self.Click_SellButton = self.Normal_Sell_Order_Entry.child_window(title="Sell",auto_id="OrderPlaceBtn",class_name="Button").click_input()
                    keyboard.send_keys('{ESC}')
                return str(self.Remark)
            case _:
                return "NA"







    def Focus_On_Exe(self):
        self.path = LoginTestData.Get_LoginExe_path()  # EXE Path
        self.app = Application(backend='uia').connect(path=self.path)
        self.App_MainWnd = self.app.window(title="Noren Trader", framework_id="WPF")
        self.App_MainWnd.set_focus()
    def Get_Place_OrdDetails(self,Result_path,Remark,Expected_Quantity,Expected_Price,Expected_TriggerPrice,Expected_DiscQty,Expected_OpenQty,Expected_TradeQty,Expected_AvgPrice,TC_NO,row_value,Test_Scenario,Expected_TransType,Expected_OrdType,TypeOfOrder):
        keyboard.send_keys('{F3}')  # Invoke Order Book
        keyboard.send_keys('{F3}')  # Invoke Order Book
        self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook", class_name="OrderBook")
        if self.OrderBook.exists(timeout=100):
            print("OB is open ")
        else:
            print("OB does not exist")
        self.OrderBook.set_focus()
        # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
        self.Clear_filter = self.OrderBook.child_window(title="Clear Filters", class_name="Button").click_input()
        self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
        self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders", class_name="OpenOrdBook")
        self.PendingOrderGrp.click_input()
        self.FilterRow_dataGrid = self.PendingOrderGrp.child_window(auto_id="uiOrderBookOpen")
        # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
        self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
        # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
        keyboard.send_keys('^F')
        # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
        self.PendingOrderGrp.type_keys(str(Remark))
        # ------------------------------------------------------Get Order Row From OrderBook ------------------------------------------------#
        keyboard.send_keys('{DOWN}')                                     #To select Order Row
        keyboard.send_keys('^C')                                         #To Copy Order Row
        keyboard.send_keys('^C')
        text = pyperclip.paste()                                         #Paste Row values in Clipboard
        array_from_copied_content = text.split()                         #Split Row and store in the Array
        print("Array from copied content:", array_from_copied_content)
        Length = len(array_from_copied_content) /2                       #Devide Array by 2 to get total Legth of Order Row
        Order_Detail_Array_Length = math.trunc(Length)                   #Get Truncated Value
        # ------------------------------------------------------Get Qty Value in OrderDetails ------------------------------------------------#
        try :
            Qty_index_in_Array = array_from_copied_content.index("Qty")  # Search string "Qty" in Array
            Qty_value_index_in_Array = Order_Detail_Array_Length + Qty_index_in_Array  # Get Qty value's index number
            Qty_i = round(Qty_value_index_in_Array, 0)  # Truncade Decimal value
            Actual_Qty = array_from_copied_content[Qty_i]  # Get Qty Value using Index
            print("Expected Qty is:" + str(Expected_Quantity))
            print("Actual Qty is :" + str(Actual_Qty))
        except ValueError:
            print("Order Not present in the Order Book --> Order may be rejected ")
        # ------------------------------------------------------Get Price Value in OrderDetails ------------------------------------------------#
        Price_index_in_Array = array_from_copied_content.index("Price")  # Search string "Price" in Array
        Price_value_index_in_Array = Order_Detail_Array_Length + Price_index_in_Array  # Get Price value's index number
        Price_i = round(Price_value_index_in_Array, 0)     # Truncade Decimal value
        Actual_Price = array_from_copied_content[Price_i]  # Get Price Value using Index
        print("Expected Price is :" + str(Expected_Price))
        print("Actual Price is :" + str(Actual_Price))
        # ------------------------------------------------------Get TriggerPrice Value in OrderDetails ------------------------------------------------#
        TriggerPrice_index_in_Array = array_from_copied_content.index("TriggerPrice")  # Search string "TriggerPrice" in Array
        TriggerPrice_value_index_in_Array = Order_Detail_Array_Length + TriggerPrice_index_in_Array  # Get TriggerPrice value's index number
        TriggerPrice_i = round(TriggerPrice_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TriggerPrice = array_from_copied_content[TriggerPrice_i]  # Get Price Value using Index
        print("Expected TriggerPrice is :" + str(Expected_TriggerPrice))
        print("Actual TriggerPrice is :" + str(Actual_TriggerPrice))
        # ------------------------------------------------------Get DiscQty Value in OrderDetails ------------------------------------------------#
        DiscQty_index_in_Array = array_from_copied_content.index("DiscQty")  # Search string "DiscQty" in Array
        DiscQty_value_index_in_Array = Order_Detail_Array_Length + DiscQty_index_in_Array  # Get DiscQty value's index number
        DiscQty_i = round(DiscQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_DiscQty = array_from_copied_content[DiscQty_i]  # Get DiscQty Value using Index
        print("Expected DiscQty is :" + str(Expected_DiscQty))
        print("Actual DiscQty is :" + str(Actual_DiscQty))
        # ------------------------------------------------------Get OpenQty Value in OrderDetails ------------------------------------------------#
        OpenQty_index_in_Array = array_from_copied_content.index("OpenQty")  # Search string "OpenQty" in Array
        OpenQty_value_index_in_Array = Order_Detail_Array_Length + OpenQty_index_in_Array  # Get OpenQty value's index number
        OpenQty_i = round(OpenQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_OpenQty = array_from_copied_content[OpenQty_i]  # Get OpenQty Value using Index
        print("Expected OpenQty is :" + str(Expected_OpenQty))
        print("Actual OpenQty is :" + str(Actual_OpenQty))
        # ------------------------------------------------------Get TradeQty Value in OrderDetails ------------------------------------------------#
        TradeQty_index_in_Array = array_from_copied_content.index("TradeQty")  # Search string "OpenQty" in Array
        TradeQty_value_index_in_Array = Order_Detail_Array_Length + TradeQty_index_in_Array  # Get TradeQty value's index number
        TradeQty_i = round(TradeQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TradeQty = array_from_copied_content[TradeQty_i]  # Get TradeQty Value using Index
        print("Expected TradeQty is :" + str(Expected_TradeQty))
        print("Actual TradeQty is :" + str(Actual_TradeQty))
        # ------------------------------------------------------Get AvgPrice Value in OrderDetails ------------------------------------------------#
        AvgPrice_index_in_Array = array_from_copied_content.index("AvgPrice")  # Search string "OpenQty" in Array
        AvgPrice_value_index_in_Array = Order_Detail_Array_Length + AvgPrice_index_in_Array  # Get AvgPrice value's index number
        AvgPrice_i = round(AvgPrice_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_AvgPrice = array_from_copied_content[AvgPrice_i]  # Get AvgPrice Value using Index
        print("Expected AvgPrice is :" + str(Expected_AvgPrice))
        print("Actual AvgPrice is :" + str(Actual_AvgPrice))
        # ------------------------------------------------------Get Transaction Type Value in OrderDetails ------------------------------------------------#
        TransType_index_in_Array = array_from_copied_content.index("BuySell")  # Search string "BuySell" in Array
        TransType_value_index_in_Array = Order_Detail_Array_Length + TransType_index_in_Array  # Get TransType value's index number
        TransType_i = round(TransType_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TransType = array_from_copied_content[TransType_i]  # Get AvgPrice Value using Index
        print("Expected TransType is :" + str(Expected_TransType))
        print("Actual TransType is :" + str(Actual_TransType))
        # ------------------------------------------------------Get Order Type Value in OrderDetails ------------------------------------------------#
        OrderType_index_in_Array = array_from_copied_content.index("OrdType")  # Search string "OrdType" in Array
        OrdType_value_index_in_Array = Order_Detail_Array_Length + OrderType_index_in_Array  # Get OrdType value's index number
        OrdType_i = round(OrdType_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_OrdType = array_from_copied_content[OrdType_i]  # Get OrderType Value using Index
        print("Expected OrderType is :" + str(Expected_OrdType))
        print("Actual OrderType is :" + str(Actual_OrdType))



        # ------------------------------------------------------Compare OrderDetails ------------------------------------------------#
        if str(Expected_Quantity) == str(Actual_Qty) and str(Expected_Price) == str(Actual_Price) and str(Expected_OpenQty) == str(Actual_OpenQty) and str(Expected_TradeQty) == str(Actual_TradeQty) and str(Expected_AvgPrice) == str(Actual_AvgPrice) and str(Expected_TransType) == str(Actual_TransType) and str(Expected_TriggerPrice) == str(Actual_TriggerPrice) and str(Expected_DiscQty) == str(Actual_DiscQty) and str(Expected_OrdType) == str(Actual_OrdType):
            Test_Status = "PASS"
            print("TC is Pass")
            print(str(Expected_OrdType) + str(Actual_OrdType))
            if str(TypeOfOrder) == "LMT":
                LG.Write_log_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty)
                #LG.Write_Excel_ActualResult_LMTOrder(Result_path,TC_NO, row_value, Test_Scenario, Test_Status)
            if str(TypeOfOrder) == "SL-LMT":
                LG.Write_log_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
                #LG.Write_Excel_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Test_Status)
            if str(TypeOfOrder) == "MKT":
                LG.Write_log_ActualResult_MKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "SL-MKT":
                LG.Write_log_ActualResult_SLMKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)

        else:
            Test_Status = "FAIL"
            print("TC is Fail")
            if str(TypeOfOrder) == "LMT":
                LG.Write_log_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty)
                #LG.Write_Excel_ActualResult_LMTOrder(Result_path,TC_NO, row_value, Test_Scenario, Test_Status)
            if str(TypeOfOrder) == "SL-LMT":
                LG.Write_log_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
                #LG.Write_Excel_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Test_Status)
            if str(TypeOfOrder) == "MKT":
                LG.Write_log_ActualResult_MKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice,Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "SL-MKT":
                LG.Write_log_ActualResult_SLMKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)

        # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
        keyboard.send_keys('^F')

    def Get_Modify_OrdDetails(self,Result_path ,Remark,Expected_Quantity,Expected_Price,Expected_TriggerPrice,Expected_DiscQty,Expected_OpenQty,Expected_TradeQty,Expected_AvgPrice,TC_NO, row_value,Test_Scenario,Expected_TransType,Expected_OrdType,TypeOfOrder):
        time.sleep(1)
        keyboard.send_keys('^C')                                    # To Copy Order Row
        keyboard.send_keys('^C')
        text = pyperclip.paste()                                    # Paste Row values in Clipboard
        array_from_copied_content = text.split()                    # Split Row and store in the Array
        print("Array from copied content:", array_from_copied_content)
        # print("Order Details"+str(len(array_from_copied_content)))
        # print("Qty: "+str(array_from_copied_content.index("Qty")))
        Length = len(array_from_copied_content) / 2                 # Devide Array by 2 to get total Legth of Order Row
        Order_Detail_Array_Length = math.trunc(Length)              # Get Truncated Value
        # ------------------------------------------------------Get Qty Value in OrderDetails ------------------------------------------------#
        Qty_index_in_Array = array_from_copied_content.index("Qty")  # Search string "Qty" in Array
        # print("Qty_Index "+ str(Qty_index_in_Array))
        Qty_value_index_in_Array = Order_Detail_Array_Length + Qty_index_in_Array  # Get Qty value's index number
        Qty_i = round(Qty_value_index_in_Array, 0)                   # Truncade Decimal value
        # print("Qty_value: " + str(Qty_i))
        Actual_Qty = array_from_copied_content[Qty_i]                # Get Qty Value using Index
        print("Expected Qty is :" + str(Expected_Quantity))
        print("Actual Qty is :" + str(Actual_Qty))
        # ------------------------------------------------------Get Price Value in OrderDetails ------------------------------------------------#
        Price_index_in_Array = array_from_copied_content.index("Price")  # Search string "Price" in Array
        # print("Price_Index "+ str(Price_index_in_Array))
        Price_value_index_in_Array = Order_Detail_Array_Length + Price_index_in_Array  # Get Price value's index number
        Price_i = round(Price_value_index_in_Array, 0)                   # Truncade Decimal value
        # print("Price_value: " + str(Price_i))
        Actual_Price = array_from_copied_content[Price_i]                # Get Price Value using Index
        print("Expected Price is :" + str(Expected_Price))
        print("Actual Price is :" + str(Actual_Price))
        # ------------------------------------------------------Get TriggerPrice Value in OrderDetails ------------------------------------------------#
        TriggerPrice_index_in_Array = array_from_copied_content.index("TriggerPrice")  # Search string "TriggerPrice" in Array
        TriggerPrice_value_index_in_Array = Order_Detail_Array_Length + TriggerPrice_index_in_Array  # Get TriggerPrice value's index number
        TriggerPrice_i = round(TriggerPrice_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TriggerPrice = array_from_copied_content[TriggerPrice_i]  # Get Price Value using Index
        print("Expected TriggerPrice is :" + str(Expected_TriggerPrice))
        print("Actual TriggerPrice is :" + str(Actual_TriggerPrice))
        # ------------------------------------------------------Get DiscQty Value in OrderDetails ------------------------------------------------#
        DiscQty_index_in_Array = array_from_copied_content.index("DiscQty")  # Search string "DiscQty" in Array
        DiscQty_value_index_in_Array = Order_Detail_Array_Length + DiscQty_index_in_Array  # Get DiscQty value's index number
        DiscQty_i = round(DiscQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_DiscQty = array_from_copied_content[DiscQty_i]  # Get DiscQty Value using Index
        print("Expected DiscQty is :" + str(Expected_DiscQty))
        print("Actual DiscQty is :" + str(Actual_DiscQty))
        # ------------------------------------------------------Get OpenQty Value in OrderDetails ------------------------------------------------#
        OpenQty_index_in_Array = array_from_copied_content.index("OpenQty")  # Search string "OpenQty" in Array
        OpenQty_value_index_in_Array = Order_Detail_Array_Length + OpenQty_index_in_Array  # Get OpenQty value's index number
        OpenQty_i = round(OpenQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_OpenQty = array_from_copied_content[OpenQty_i]  # Get OpenQty Value using Index
        print("Expected OpenQty is :" + str(Expected_OpenQty))
        print("Actual OpenQty is :" + str(Actual_OpenQty))
        # ------------------------------------------------------Get TradeQty Value in OrderDetails ------------------------------------------------#
        TradeQty_index_in_Array = array_from_copied_content.index("TradeQty")  # Search string "OpenQty" in Array
        TradeQty_value_index_in_Array = Order_Detail_Array_Length + TradeQty_index_in_Array  # Get TradeQty value's index number
        TradeQty_i = round(TradeQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TradeQty = array_from_copied_content[TradeQty_i]  # Get TradeQty Value using Index
        print("Expected TradeQty is :" + str(Expected_TradeQty))
        print("Actual TradeQty is :" + str(Actual_TradeQty))
        # ------------------------------------------------------Get AvgPrice Value in OrderDetails ------------------------------------------------#
        AvgPrice_index_in_Array = array_from_copied_content.index("AvgPrice")  # Search string "OpenQty" in Array
        AvgPrice_value_index_in_Array = Order_Detail_Array_Length + AvgPrice_index_in_Array  # Get AvgPrice value's index number
        AvgPrice_i = round(AvgPrice_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_AvgPrice = array_from_copied_content[AvgPrice_i]  # Get AvgPrice Value using Index
        print("Expected AvgPrice is :" + str(Expected_AvgPrice))
        print("Actual AvgPrice is :" + str(Actual_AvgPrice))
        # ------------------------------------------------------Get Transaction Type Value in OrderDetails ------------------------------------------------#
        TransType_index_in_Array = array_from_copied_content.index("BuySell")  # Search string "BuySell" in Array
        TransType_value_index_in_Array = Order_Detail_Array_Length + TransType_index_in_Array  # Get TransType value's index number
        TransType_i = round(TransType_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TransType = array_from_copied_content[TransType_i]  # Get AvgPrice Value using Index
        print("Expected TransType is :" + str(Expected_TransType))
        print("Actual TransType is :" + str(Actual_TransType))
        # ------------------------------------------------------Get Order Type Value in OrderDetails ------------------------------------------------#
        OrderType_index_in_Array = array_from_copied_content.index("OrdType")  # Search string "OrdType" in Array
        OrdType_value_index_in_Array = Order_Detail_Array_Length + OrderType_index_in_Array  # Get OrdType value's index number
        OrdType_i = round(OrdType_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_OrdType = array_from_copied_content[OrdType_i]  # Get OrderType Value using Index
        print("Expected OrderType is :" + str(Expected_OrdType))
        print("Actual OrderType is :" + str(Actual_OrdType))
        # ------------------------------------------------------Compare OrderDetails ------------------------------------------------#
        if str(Expected_Quantity) == str(Actual_Qty) and str(Expected_Price) == str(Actual_Price)and str(Expected_OpenQty) == str(Actual_OpenQty) and str(Expected_TradeQty) == str(Actual_TradeQty) and str(Expected_AvgPrice) == str(Actual_AvgPrice) and str(Expected_TransType) == str(Actual_TransType) and str(Expected_TriggerPrice) == str(Actual_TriggerPrice) and str(Expected_DiscQty) == str(Actual_DiscQty) and str(Expected_OrdType) == str(Actual_OrdType):
            Test_Status = "PASS"
            #print("Ordertype"+str(OrdType))
            print("TC is Pass")
            if str(TypeOfOrder) == "LMT":
                LG.Write_log_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price,Actual_OpenQty,Expected_OpenQty,Actual_TradeQty,Expected_TradeQty,Actual_AvgPrice,Expected_AvgPrice,Test_Status,Expected_TransType,Actual_TransType,Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty)
                #LG.Write_Excel_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Test_Status)
            if str(TypeOfOrder) == "SL-LMT":
                LG.Write_log_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
                #LG.Write_Excel_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Test_Status)
            if str(TypeOfOrder) == "MKT":
                LG.Write_log_ActualResult_MKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "SL-MKT":
                LG.Write_log_ActualResult_SLMKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)

        else:
            Test_Status = "FAIL"
            print("TC is Fail")
            if str(TypeOfOrder) == "LMT":
                LG.Write_log_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price,Actual_OpenQty,Expected_OpenQty,Actual_TradeQty,Expected_TradeQty,Actual_AvgPrice,Expected_AvgPrice,Test_Status,Expected_TransType,Actual_TransType,Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty)
                #LG.Write_Excel_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Test_Status)
            if str(TypeOfOrder) == "SL-LMT":
                LG.Write_log_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
                #LG.Write_Excel_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Test_Status)
            if str(TypeOfOrder) == "MKT":
                LG.Write_log_ActualResult_MKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "SL-MKT":
                LG.Write_log_ActualResult_SLMKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)


    def Get_Cancel_OrdDetails(self,Result_path,Remark,Expected_Quantity,Expected_Price,Expected_TriggerPrice,Expected_DiscQty,Expected_OpenQty,Expected_TradeQty,Expected_AvgPrice,TC_NO, row_value,Test_Scenario,Expected_TransType,Expected_OrdType,TypeOfOrder):
        keyboard.send_keys('{F3}')  # Invoke Order Book
        self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook", class_name="OrderBook")
        self.OrderBook.set_focus()
        # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
        self.Clear_filter = self.OrderBook.child_window(title="Clear Filters", class_name="Button").click_input()
        self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
        self.ClosedOrderGrp = self.OrderBookLayout.child_window(title="Closed Orders", class_name="CompOrdBook")
        self.ClosedOrderGrp.click_input()
        self.FilterRow_dataGrid = self.ClosedOrderGrp.child_window(auto_id="uiOrderBookExe")
        # -----------------------------------------------Select Closed Order Tab column Header-----------------------------------------------#
        self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
        #self.Header.click_input()
        # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
        keyboard.send_keys('^F')
        # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
        self.ClosedOrderGrp.type_keys(str(Remark))
        keyboard.send_keys('{DOWN}')
        keyboard.send_keys('^C')  # To Copy Order Row
        keyboard.send_keys('^C')
        text = pyperclip.paste()  # Paste Row values in Clipboard
        array_from_copied_content = text.split()  # Split Row and store in the Array
        print("Array from copied content:", array_from_copied_content)
        # print("Order Details"+str(len(array_from_copied_content)))
        # print("Qty: "+str(array_from_copied_content.index("Qty")))
        Length = len(array_from_copied_content) / 2  # Devide Array by 2 to get total Legth of Order Row
        Order_Detail_Array_Length = math.trunc(Length)  # Get Truncated Value
        # ------------------------------------------------------Get Qty Value in OrderDetails ------------------------------------------------#
        Qty_index_in_Array = array_from_copied_content.index("Qty")  # Search string "Qty" in Array
        # print("Qty_Index "+ str(Qty_index_in_Array))
        Qty_value_index_in_Array = Order_Detail_Array_Length + Qty_index_in_Array  # Get Qty value's index number
        Qty_i = round(Qty_value_index_in_Array, 0)  # Truncade Decimal value
        # print("Qty_value: " + str(Qty_i))
        Actual_Qty = array_from_copied_content[Qty_i]  # Get Qty Value using Index
        print("Expected Qty is :" + str(Expected_Quantity))
        print("Actual Qty is :" + str(Actual_Qty))
        # ------------------------------------------------------Get Price Value in OrderDetails ------------------------------------------------#
        Price_index_in_Array = array_from_copied_content.index("Price")  # Search string "Price" in Array
        # print("Price_Index "+ str(Price_index_in_Array))
        Price_value_index_in_Array = Order_Detail_Array_Length + Price_index_in_Array  # Get Price value's index number
        Price_i = round(Price_value_index_in_Array, 0)  # Truncade Decimal value
        # print("Price_value: " + str(Price_i))
        Actual_Price = array_from_copied_content[Price_i]  # Get Price Value using Index
        print("Expected Price is :" + str(Expected_Price))
        print("Actual Price is :" + str(Actual_Price))
        # ------------------------------------------------------Get TriggerPrice Value in OrderDetails ------------------------------------------------#
        TriggerPrice_index_in_Array = array_from_copied_content.index("TriggerPrice")  # Search string "TriggerPrice" in Array
        TriggerPrice_value_index_in_Array = Order_Detail_Array_Length + TriggerPrice_index_in_Array  # Get TriggerPrice value's index number
        TriggerPrice_i = round(TriggerPrice_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TriggerPrice = array_from_copied_content[TriggerPrice_i]  # Get Price Value using Index
        print("Expected TriggerPrice is :" + str(Expected_TriggerPrice))
        print("Actual TriggerPrice is :" + str(Actual_TriggerPrice))
        # ------------------------------------------------------Get DiscQty Value in OrderDetails ------------------------------------------------#
        DiscQty_index_in_Array = array_from_copied_content.index("DiscQty")  # Search string "DiscQty" in Array
        DiscQty_value_index_in_Array = Order_Detail_Array_Length + DiscQty_index_in_Array  # Get DiscQty value's index number
        DiscQty_i = round(DiscQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_DiscQty = array_from_copied_content[DiscQty_i]  # Get DiscQty Value using Index
        print("Expected DiscQty is :" + str(Expected_DiscQty))
        print("Actual DiscQty is :" + str(Actual_DiscQty))
        # ------------------------------------------------------Get OpenQty Value in OrderDetails ------------------------------------------------#
        OpenQty_index_in_Array = array_from_copied_content.index("OpenQty")  # Search string "OpenQty" in Array
        OpenQty_value_index_in_Array = Order_Detail_Array_Length + OpenQty_index_in_Array  # Get OpenQty value's index number
        OpenQty_i = round(OpenQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_OpenQty = array_from_copied_content[OpenQty_i]  # Get OpenQty Value using Index
        print("Expected OpenQty is :" + str(Expected_OpenQty))
        print("Actual OpenQty is :" + str(Actual_OpenQty))
        # ------------------------------------------------------Get TradeQty Value in OrderDetails ------------------------------------------------#
        TradeQty_index_in_Array = array_from_copied_content.index("TradeQty")  # Search string "OpenQty" in Array
        TradeQty_value_index_in_Array = Order_Detail_Array_Length + TradeQty_index_in_Array  # Get TradeQty value's index number
        TradeQty_i = round(TradeQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TradeQty = array_from_copied_content[TradeQty_i]  # Get TradeQty Value using Index
        print("Expected TradeQty is :" + str(Expected_TradeQty))
        print("Actual TradeQty is :" + str(Actual_TradeQty))
        # ------------------------------------------------------Get AvgPrice Value in OrderDetails ------------------------------------------------#
        AvgPrice_index_in_Array = array_from_copied_content.index("AvgPrice")  # Search string "OpenQty" in Array
        AvgPrice_value_index_in_Array = Order_Detail_Array_Length + AvgPrice_index_in_Array  # Get AvgPrice value's index number
        AvgPrice_i = round(AvgPrice_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_AvgPrice = array_from_copied_content[AvgPrice_i]  # Get AvgPrice Value using Index
        print("Expected AvgPrice is :" + str(Expected_AvgPrice))
        print("Actual AvgPrice is :" + str(Actual_AvgPrice))
        # ------------------------------------------------------Get Transaction Type Value in OrderDetails ------------------------------------------------#
        TransType_index_in_Array = array_from_copied_content.index("BuySell")  # Search string "BuySell" in Array
        TransType_value_index_in_Array = Order_Detail_Array_Length + TransType_index_in_Array  # Get TransType value's index number
        TransType_i = round(TransType_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TransType = array_from_copied_content[TransType_i]  # Get AvgPrice Value using Index
        print("Expected TransType is :" + str(Expected_TransType))
        print("Actual TransType is :" + str(Actual_TransType))
        # ------------------------------------------------------Get Order Type Value in OrderDetails ------------------------------------------------#
        OrderType_index_in_Array = array_from_copied_content.index("OrdType")  # Search string "OrdType" in Array
        OrdType_value_index_in_Array = Order_Detail_Array_Length + OrderType_index_in_Array  # Get OrdType value's index number
        OrdType_i = round(OrdType_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_OrdType = array_from_copied_content[OrdType_i]  # Get OrderType Value using Index
        print("Expected OrderType is :" + str(Expected_OrdType))
        print("Actual OrderType is :" + str(Actual_OrdType))
        # ------------------------------------------------------Compare OrderDetails ------------------------------------------------#
        if str(Expected_Quantity) == str(Actual_Qty) and str(Expected_Price) == str(Actual_Price) and str(Expected_OpenQty) == str(Actual_OpenQty) and str(Expected_TradeQty) == str(Actual_TradeQty) and str(Expected_AvgPrice) == str(Actual_AvgPrice) and str(Expected_TransType) == str(Actual_TransType) and str(Expected_TriggerPrice) == str(Actual_TriggerPrice) and str(Expected_DiscQty) == str(Actual_DiscQty) and str(Expected_OrdType) == str(Actual_OrdType):
            Test_Status = "PASS"
            print("TC is Pass")
            if str(TypeOfOrder) == "LMT":
                LG.Write_log_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty)
                #LG.Write_Excel_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Test_Status)
            if str(TypeOfOrder) == "SL-LMT":
                LG.Write_log_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status,Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
                #LG.Write_Excel_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Test_Status)
            if str(TypeOfOrder) == "MKT":
                LG.Write_log_ActualResult_MKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "SL-MKT":
                LG.Write_log_ActualResult_SLMKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)

        else:
            Test_Status = "FAIL"
            print("TC is Fail")
            if str(TypeOfOrder) == "LMT":
                LG.Write_log_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price,Actual_OpenQty,Expected_OpenQty,Actual_TradeQty,Expected_TradeQty,Actual_AvgPrice,Expected_AvgPrice,Test_Status,Expected_TransType,Actual_TransType,Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty)
                #LG.Write_Excel_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Test_Status)
            if str(TypeOfOrder) == "SL-LMT":
                LG.Write_log_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
                #LG.Write_Excel_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Test_Status)
            if str(TypeOfOrder) == "MKT":
                LG.Write_log_ActualResult_MKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "SL-MKT":
                LG.Write_log_ActualResult_SLMKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)

        keyboard.send_keys('^F')
    def Get_PlaceReject_OrdDetails(self,Result_path,Remark,Expected_Quantity,Expected_Price,Expected_TriggerPrice,Expected_DiscQty,Expected_OpenQty,Expected_TradeQty,Expected_AvgPrice,TC_NO, row_value,Test_Scenario,Expected_TransType,Expected_OrdType,TypeOfOrder):
        keyboard.send_keys('{F3}')  # Invoke Order Book
        self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook", class_name="OrderBook")
        self.OrderBook.set_focus()
        # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
        self.Clear_filter = self.OrderBook.child_window(title="Clear Filters", class_name="Button").click_input()
        self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
        self.ClosedOrderGrp = self.OrderBookLayout.child_window(title="Closed Orders", class_name="CompOrdBook")
        self.ClosedOrderGrp.click_input()
        self.FilterRow_dataGrid = self.ClosedOrderGrp.child_window(auto_id="uiOrderBookExe")
        # -----------------------------------------------Select Closed Order Tab column Header-----------------------------------------------#
        self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
        #self.Header.click_input()
        # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
        keyboard.send_keys('^F')
        # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
        self.ClosedOrderGrp.type_keys(str(Remark))
        keyboard.send_keys('{DOWN}')
        keyboard.send_keys('^C')  # To Copy Order Row
        keyboard.send_keys('^C')
        text = pyperclip.paste()  # Paste Row values in Clipboard
        array_from_copied_content = text.split()  # Split Row and store in the Array
        print("Array from copied content:", array_from_copied_content)
        # print("Order Details"+str(len(array_from_copied_content)))
        # print("Qty: "+str(array_from_copied_content.index("Qty")))
        Length = len(array_from_copied_content) / 2  # Devide Array by 2 to get total Legth of Order Row
        Order_Detail_Array_Length = math.trunc(Length)  # Get Truncated Value
        # ------------------------------------------------------Get Qty Value in OrderDetails ------------------------------------------------#
        Qty_index_in_Array = array_from_copied_content.index("Qty")  # Search string "Qty" in Array
        # print("Qty_Index "+ str(Qty_index_in_Array))
        Qty_value_index_in_Array = Order_Detail_Array_Length + Qty_index_in_Array  # Get Qty value's index number
        Qty_i = round(Qty_value_index_in_Array, 0)  # Truncade Decimal value
        # print("Qty_value: " + str(Qty_i))
        Actual_Qty = array_from_copied_content[Qty_i]  # Get Qty Value using Index
        print("Expected Qty is :" + str(Expected_Quantity))
        print("Actual Qty is :" + str(Actual_Qty))
        # ------------------------------------------------------Get Price Value in OrderDetails ------------------------------------------------#
        Price_index_in_Array = array_from_copied_content.index("Price")  # Search string "Price" in Array
        # print("Price_Index "+ str(Price_index_in_Array))
        Price_value_index_in_Array = Order_Detail_Array_Length + Price_index_in_Array  # Get Price value's index number
        Price_i = round(Price_value_index_in_Array, 0)  # Truncade Decimal value
        # print("Price_value: " + str(Price_i))
        Actual_Price = array_from_copied_content[Price_i]  # Get Price Value using Index
        print("Expected Price is :" + str(Expected_Price))
        print("Actual Price is :" + str(Actual_Price))
        # ------------------------------------------------------Get TriggerPrice Value in OrderDetails ------------------------------------------------#
        TriggerPrice_index_in_Array = array_from_copied_content.index("TriggerPrice")  # Search string "TriggerPrice" in Array
        TriggerPrice_value_index_in_Array = Order_Detail_Array_Length + TriggerPrice_index_in_Array  # Get TriggerPrice value's index number
        TriggerPrice_i = round(TriggerPrice_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TriggerPrice = array_from_copied_content[TriggerPrice_i]  # Get Price Value using Index
        print("Expected TriggerPrice is :" + str(Expected_TriggerPrice))
        print("Actual TriggerPrice is :" + str(Actual_TriggerPrice))
        # ------------------------------------------------------Get DiscQty Value in OrderDetails ------------------------------------------------#
        DiscQty_index_in_Array = array_from_copied_content.index("DiscQty")  # Search string "DiscQty" in Array
        DiscQty_value_index_in_Array = Order_Detail_Array_Length + DiscQty_index_in_Array  # Get DiscQty value's index number
        DiscQty_i = round(DiscQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_DiscQty = array_from_copied_content[DiscQty_i]  # Get DiscQty Value using Index
        print("Expected DiscQty is :" + str(Expected_DiscQty))
        print("Actual DiscQty is :" + str(Actual_DiscQty))
        # ------------------------------------------------------Get OpenQty Value in OrderDetails ------------------------------------------------#
        OpenQty_index_in_Array = array_from_copied_content.index("OpenQty")  # Search string "OpenQty" in Array
        OpenQty_value_index_in_Array = Order_Detail_Array_Length + OpenQty_index_in_Array  # Get OpenQty value's index number
        OpenQty_i = round(OpenQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_OpenQty = array_from_copied_content[OpenQty_i]  # Get OpenQty Value using Index
        print("Expected OpenQty is :" + str(Expected_OpenQty))
        print("Actual OpenQty is :" + str(Actual_OpenQty))
        # ------------------------------------------------------Get TradeQty Value in OrderDetails ------------------------------------------------#
        TradeQty_index_in_Array = array_from_copied_content.index("TradeQty")  # Search string "OpenQty" in Array
        TradeQty_value_index_in_Array = Order_Detail_Array_Length + TradeQty_index_in_Array  # Get TradeQty value's index number
        TradeQty_i = round(TradeQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TradeQty = array_from_copied_content[TradeQty_i]  # Get TradeQty Value using Index
        print("Expected TradeQty is :" + str(Expected_TradeQty))
        print("Actual TradeQty is :" + str(Actual_TradeQty))
        # ------------------------------------------------------Get AvgPrice Value in OrderDetails ------------------------------------------------#
        AvgPrice_index_in_Array = array_from_copied_content.index("AvgPrice")  # Search string "OpenQty" in Array
        AvgPrice_value_index_in_Array = Order_Detail_Array_Length + AvgPrice_index_in_Array  # Get AvgPrice value's index number
        AvgPrice_i = round(AvgPrice_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_AvgPrice = array_from_copied_content[AvgPrice_i]  # Get AvgPrice Value using Index
        print("Expected AvgPrice is :" + str(Expected_AvgPrice))
        print("Actual AvgPrice is :" + str(Actual_AvgPrice))
        # ------------------------------------------------------Get Transaction Type Value in OrderDetails ------------------------------------------------#
        TransType_index_in_Array = array_from_copied_content.index("BuySell")  # Search string "BuySell" in Array
        TransType_value_index_in_Array = Order_Detail_Array_Length + TransType_index_in_Array  # Get TransType value's index number
        TransType_i = round(TransType_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TransType = array_from_copied_content[TransType_i]  # Get AvgPrice Value using Index
        print("Expected TransType is :" + str(Expected_TransType))
        print("Actual TransType is :" + str(Actual_TransType))
        # ------------------------------------------------------Get Order Type Value in OrderDetails ------------------------------------------------#
        OrderType_index_in_Array = array_from_copied_content.index("OrdType")  # Search string "OrdType" in Array
        OrdType_value_index_in_Array = Order_Detail_Array_Length + OrderType_index_in_Array  # Get OrdType value's index number
        OrdType_i = round(OrdType_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_OrdType = array_from_copied_content[OrdType_i]  # Get OrderType Value using Index
        print("Expected OrderType is :" + str(Expected_OrdType))
        print("Actual OrderType is :" + str(Actual_OrdType))
        # ------------------------------------------------------Compare OrderDetails ------------------------------------------------#
        if str(Expected_Quantity) == str(Actual_Qty) and str(Expected_Price) == str(Actual_Price) and str(Expected_OpenQty) == str(Actual_OpenQty) and str(Expected_TradeQty) == str(Actual_TradeQty) and str(Expected_AvgPrice) == str(Actual_AvgPrice) and str(Expected_TransType) == str(Actual_TransType) and str(Expected_TriggerPrice) == str(Actual_TriggerPrice) and str(Expected_DiscQty) == str(Actual_DiscQty) and str(Expected_OrdType) == str(Actual_OrdType):
            Test_Status = "PASS"
            print("TC is Pass")
            if str(TypeOfOrder) == "LMT":
                LG.Write_log_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty)
            if str(TypeOfOrder) == "SL-LMT":
                LG.Write_log_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status,Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "MKT":
                LG.Write_log_ActualResult_MKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "SL-MKT":
                LG.Write_log_ActualResult_SLMKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)

        else:
            Test_Status = "FAIL"
            print("TC is Fail")
            if str(TypeOfOrder) == "LMT":
                LG.Write_log_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price,Actual_OpenQty,Expected_OpenQty,Actual_TradeQty,Expected_TradeQty,Actual_AvgPrice,Expected_AvgPrice,Test_Status,Expected_TransType,Actual_TransType,Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty)
            if str(TypeOfOrder) == "SL-LMT":
                LG.Write_log_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "MKT":
                LG.Write_log_ActualResult_MKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradeQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "SL-MKT":
                LG.Write_log_ActualResult_SLMKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradeQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)

        keyboard.send_keys('^F')
    def Get_PartialTrade_OrdDetails(self,Result_path,Remark,Expected_Quantity,Expected_Price,Expected_TriggerPrice,Expected_DiscQty,Expected_OpenQty,Expected_TradedQty,Expected_AvgPrice,TC_NO, row_value,Test_Scenario,Expected_TransType,Expected_OrdType,TypeOfOrder):
        keyboard.send_keys('{F3}')  # Invoke Order Book
        keyboard.send_keys('{F3}')  # Invoke Order Book
        self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook", class_name="OrderBook")
        self.OrderBook.set_focus()
        # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
        self.Clear_filter = self.OrderBook.child_window(title="Clear Filters", class_name="Button").click_input()
        self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
        self.PendingOrderGrp = self.OrderBookLayout.child_window(title="Pending Orders", class_name="OpenOrdBook")
        self.PendingOrderGrp.click_input()
        self.FilterRow_dataGrid = self.PendingOrderGrp.child_window(auto_id="uiOrderBookOpen")
        # -----------------------------------------------Select Pending Order Tab column Header-----------------------------------------------#
        self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel",class_name="ItemsControlBase")
        # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
        keyboard.send_keys('^F')
        # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
        self.PendingOrderGrp.type_keys(str(Remark))
        # ------------------------------------------------------Get Order Row From OrderBook ------------------------------------------------#
        keyboard.send_keys('{DOWN}')  # To select Order Row
        keyboard.send_keys('^C')  # To Copy Order Row
        keyboard.send_keys('^C')
        text = pyperclip.paste()  # Paste Row values in Clipboard
        array_from_copied_content = text.split()  # Split Row and store in the Array
        print("Array from copied content:", array_from_copied_content)
        Length = len(array_from_copied_content) / 2  # Devide Array by 2 to get total Legth of Order Row
        Order_Detail_Array_Length = math.trunc(Length)  # Get Truncated Value
        # ------------------------------------------------------Get Qty Value in OrderDetails ------------------------------------------------#
        Qty_index_in_Array = array_from_copied_content.index("Qty")  # Search string "Qty" in Array
        Qty_value_index_in_Array = Order_Detail_Array_Length + Qty_index_in_Array  # Get Qty value's index number
        Qty_i = round(Qty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_Qty = array_from_copied_content[Qty_i]  # Get Qty Value using Index
        print("Expected Qty is:" + str(Expected_Quantity))
        print("Actual Qty is :" + str(Actual_Qty))
        # ------------------------------------------------------Get Price Value in OrderDetails ------------------------------------------------#
        Price_index_in_Array = array_from_copied_content.index("Price")  # Search string "Price" in Array
        Price_value_index_in_Array = Order_Detail_Array_Length + Price_index_in_Array  # Get Price value's index number
        Price_i = round(Price_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_Price = array_from_copied_content[Price_i]  # Get Price Value using Index
        print("Expected Price is :" + str(Expected_Price))
        print("Actual Price is :" + str(Actual_Price))
        # ------------------------------------------------------Get TriggerPrice Value in OrderDetails ------------------------------------------------#
        TriggerPrice_index_in_Array = array_from_copied_content.index("TriggerPrice")  # Search string "TriggerPrice" in Array
        TriggerPrice_value_index_in_Array = Order_Detail_Array_Length + TriggerPrice_index_in_Array  # Get TriggerPrice value's index number
        TriggerPrice_i = round(TriggerPrice_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TriggerPrice = array_from_copied_content[TriggerPrice_i]  # Get Price Value using Index
        print("Expected TriggerPrice is :" + str(Expected_TriggerPrice))
        print("Actual TriggerPrice is :" + str(Actual_TriggerPrice))
        # ------------------------------------------------------Get DiscQty Value in OrderDetails ------------------------------------------------#
        DiscQty_index_in_Array = array_from_copied_content.index("DiscQty")  # Search string "DiscQty" in Array
        DiscQty_value_index_in_Array = Order_Detail_Array_Length + DiscQty_index_in_Array  # Get DiscQty value's index number
        DiscQty_i = round(DiscQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_DiscQty = array_from_copied_content[DiscQty_i]  # Get DiscQty Value using Index
        print("Expected DiscQty is :" + str(Expected_DiscQty))
        print("Actual DiscQty is :" + str(Actual_DiscQty))
        # ------------------------------------------------------Get OpenQty Value in OrderDetails ------------------------------------------------#
        OpenQty_index_in_Array = array_from_copied_content.index("OpenQty")  # Search string "OpenQty" in Array
        OpenQty_value_index_in_Array = Order_Detail_Array_Length + OpenQty_index_in_Array  # Get OpenQty value's index number
        OpenQty_i = round(OpenQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_OpenQty = array_from_copied_content[OpenQty_i]  # Get OpenQty Value using Index
        print("Expected OpenQty is :" + str(Expected_OpenQty))
        print("Actual OpenQty is :" + str(Actual_OpenQty))
        # ------------------------------------------------------Get Traded Qty Value in OrderDetails ------------------------------------------------#
        TradedQty_index_in_Array = array_from_copied_content.index("TradeQty")  # Search string "TradeQty" in Array
        TradedQty_value_index_in_Array = Order_Detail_Array_Length + TradedQty_index_in_Array  # Get TradeQty value's index number
        TradeQty_i = round(TradedQty_value_index_in_Array, 0)                                  # Truncade Decimal value
        Actual_TradeQty = array_from_copied_content[TradeQty_i]                                # Get TradeQty Value using Index
        print("Expected TradeQty is :" + str(Expected_TradedQty))
        print("Actual TradeQty is :" + str(Actual_TradeQty))
        # ------------------------------------------------------Get AvgPrice Value in OrderDetails ------------------------------------------------#
        AvgPrice_index_in_Array = array_from_copied_content.index("AvgPrice")  # Search string "AvgPrice" in Array
        AvgPrice_value_index_in_Array = Order_Detail_Array_Length + AvgPrice_index_in_Array  # Get AvgPrice value's index number
        AvgPrice_i = round(AvgPrice_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_AvgPrice = array_from_copied_content[AvgPrice_i]  # Get AvgPrice Value using Index
        print("Expected AvgPrice is :" + str(Expected_AvgPrice))
        print("Actual AvgPrice is :" + str(Actual_AvgPrice))
        # ------------------------------------------------------Get Transaction Type Value in OrderDetails ------------------------------------------------#
        TransType_index_in_Array = array_from_copied_content.index("BuySell")  # Search string "BuySell" in Array
        TransType_value_index_in_Array = Order_Detail_Array_Length + TransType_index_in_Array  # Get TransType value's index number
        TransType_i = round(TransType_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TransType = array_from_copied_content[TransType_i]  # Get AvgPrice Value using Index
        print("Expected TransType is :" + str(Expected_TransType))
        print("Actual TransType is :" + str(Actual_TransType))
        # ------------------------------------------------------Get Order Type Value in OrderDetails ------------------------------------------------#
        OrderType_index_in_Array = array_from_copied_content.index("OrdType")  # Search string "OrdType" in Array
        OrdType_value_index_in_Array = Order_Detail_Array_Length + OrderType_index_in_Array  # Get OrdType value's index number
        OrdType_i = round(OrdType_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_OrdType = array_from_copied_content[OrdType_i]  # Get OrderType Value using Index
        print("Expected OrderType is :" + str(Expected_OrdType))
        print("Actual OrderType is :" + str(Actual_OrdType))
        # ------------------------------------------------------Compare OrderDetails ------------------------------------------------#
        if str(Expected_Quantity) == str(Actual_Qty) and str(Expected_Price) == str(Actual_Price) and str(Expected_OpenQty) == str(Actual_OpenQty) and str(Expected_TradedQty) == str(Actual_TradeQty) and str(Expected_AvgPrice) == str(Actual_AvgPrice) and str(Expected_TransType) == str(Actual_TransType) and str(Expected_TriggerPrice) == str(Actual_TriggerPrice) and str(Expected_DiscQty) == str(Actual_DiscQty) and str(Expected_OrdType) == str(Actual_OrdType):
            Test_Status = "PASS"
            print("TC is Pass")
            if str(TypeOfOrder) == "LMT":
                LG.Write_log_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradedQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty)
            if str(TypeOfOrder) == "SL-LMT":
                LG.Write_log_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradedQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status,Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "MKT":
                LG.Write_log_ActualResult_MKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradedQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "SL-MKT":
                LG.Write_log_ActualResult_SLMKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradedQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)

        else:
            Test_Status = "FAIL"
            print("TC is Fail")
            if str(TypeOfOrder) == "LMT":
                LG.Write_log_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradedQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty)
            if str(TypeOfOrder) == "SL-LMT":
                LG.Write_log_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradedQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status,Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "MKT":
                LG.Write_log_ActualResult_MKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradedQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "SL-MKT":
                LG.Write_log_ActualResult_SLMKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradedQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)

        # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
        keyboard.send_keys('^F')

    def Get_CompleteTrade_OrdDetails(self,Result_path,Remark,Expected_Quantity,Expected_Price,Expected_TriggerPrice,Expected_DiscQty,Expected_OpenQty,Expected_TradedQty,Expected_AvgPrice,TC_NO, row_value,Test_Scenario,Expected_TransType,Expected_OrdType,TypeOfOrder):
        keyboard.send_keys('{F3}')  # Invoke Order Book
        keyboard.send_keys('{F3}')  # Invoke Order Book
        self.OrderBook = self.App_MainWnd.child_window(title="OrderBook", auto_id="OrderBook", class_name="OrderBook")
        self.OrderBook.set_focus()
        # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
        self.Clear_filter = self.OrderBook.child_window(title="Clear Filters", class_name="Button").click_input()
        self.OrderBookLayout = self.App_MainWnd.child_window(title="OBLayoutManager", auto_id="OBLayoutManager",class_name="UFT_DockLayoutManager")
        self.ClosedOrderGrp = self.OrderBookLayout.child_window(title="Closed Orders", class_name="CompOrdBook")
        self.ClosedOrderGrp.click_input()
        self.FilterRow_dataGrid = self.ClosedOrderGrp.child_window(auto_id="uiOrderBookExe")
        # -----------------------------------------------Select Closed Order Tab column Header-----------------------------------------------#
        self.Header = self.FilterRow_dataGrid.child_window(title="HeaderPanel", auto_id="HeaderPanel", class_name="ItemsControlBase")
        # ------------------------------------------------------Invoke Filter Option ---------------------------------------------------------#
        keyboard.send_keys('^F')
        # ------------------------------------------------------Input Remark in Remark Filter ------------------------------------------------#
        self.ClosedOrderGrp.type_keys(str(Remark))
        keyboard.send_keys('{DOWN}')
        keyboard.send_keys('^C')  # To Copy Order Row
        keyboard.send_keys('^C')
        text = pyperclip.paste()  # Paste Row values in Clipboard
        array_from_copied_content = text.split()  # Split Row and store in the Array
        print("Array from copied content:", array_from_copied_content)
        # print("Order Details"+str(len(array_from_copied_content)))
        # print("Qty: "+str(array_from_copied_content.index("Qty")))
        Length = len(array_from_copied_content) / 2  # Devide Array by 2 to get total Legth of Order Row
        Order_Detail_Array_Length = math.trunc(Length)  # Get Truncated Value
        # ------------------------------------------------------Get Qty Value in OrderDetails ------------------------------------------------#
        Qty_index_in_Array = array_from_copied_content.index("Qty")  # Search string "Qty" in Array
        Qty_value_index_in_Array = Order_Detail_Array_Length + Qty_index_in_Array  # Get Qty value's index number
        Qty_i = round(Qty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_Qty = array_from_copied_content[Qty_i]  # Get Qty Value using Index
        print("Expected Qty is:" + str(Expected_Quantity))
        print("Actual Qty is :" + str(Actual_Qty))
        # ------------------------------------------------------Get Price Value in OrderDetails ------------------------------------------------#
        Price_index_in_Array = array_from_copied_content.index("Price")  # Search string "Price" in Array
        Price_value_index_in_Array = Order_Detail_Array_Length + Price_index_in_Array  # Get Price value's index number
        Price_i = round(Price_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_Price = array_from_copied_content[Price_i]  # Get Price Value using Index
        print("Expected Price is :" + str(Expected_Price))
        print("Actual Price is :" + str(Actual_Price))
        # ------------------------------------------------------Get TriggerPrice Value in OrderDetails ------------------------------------------------#
        TriggerPrice_index_in_Array = array_from_copied_content.index("TriggerPrice")  # Search string "TriggerPrice" in Array
        TriggerPrice_value_index_in_Array = Order_Detail_Array_Length + TriggerPrice_index_in_Array  # Get TriggerPrice value's index number
        TriggerPrice_i = round(TriggerPrice_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TriggerPrice = array_from_copied_content[TriggerPrice_i]  # Get Price Value using Index
        print("Expected TriggerPrice is :" + str(Expected_TriggerPrice))
        print("Actual TriggerPrice is :" + str(Actual_TriggerPrice))
        # ------------------------------------------------------Get DiscQty Value in OrderDetails ------------------------------------------------#
        DiscQty_index_in_Array = array_from_copied_content.index("DiscQty")  # Search string "DiscQty" in Array
        DiscQty_value_index_in_Array = Order_Detail_Array_Length + DiscQty_index_in_Array  # Get DiscQty value's index number
        DiscQty_i = round(DiscQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_DiscQty = array_from_copied_content[DiscQty_i]  # Get DiscQty Value using Index
        print("Expected DiscQty is :" + str(Expected_DiscQty))
        print("Actual DiscQty is :" + str(Actual_DiscQty))
        # ------------------------------------------------------Get OpenQty Value in OrderDetails ------------------------------------------------#
        OpenQty_index_in_Array = array_from_copied_content.index("OpenQty")  # Search string "OpenQty" in Array
        OpenQty_value_index_in_Array = Order_Detail_Array_Length + OpenQty_index_in_Array  # Get OpenQty value's index number
        OpenQty_i = round(OpenQty_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_OpenQty = array_from_copied_content[OpenQty_i]  # Get OpenQty Value using Index
        print("Expected OpenQty is :" + str(Expected_OpenQty))
        print("Actual OpenQty is :" + str(Actual_OpenQty))
        # ------------------------------------------------------Get Traded Qty Value in OrderDetails ------------------------------------------------#
        TradedQty_index_in_Array = array_from_copied_content.index("TradeQty")  # Search string "TradeQty" in Array
        TradedQty_value_index_in_Array = Order_Detail_Array_Length + TradedQty_index_in_Array  # Get TradeQty value's index number
        TradeQty_i = round(TradedQty_value_index_in_Array, 0)                                  # Truncade Decimal value
        Actual_TradeQty = array_from_copied_content[TradeQty_i]                                # Get TradeQty Value using Index
        print("Expected TradeQty is :" + str(Expected_TradedQty))
        print("Actual TradeQty is :" + str(Actual_TradeQty))
        # ------------------------------------------------------Get AvgPrice Value in OrderDetails ------------------------------------------------#
        AvgPrice_index_in_Array = array_from_copied_content.index("AvgPrice")  # Search string "AvgPrice" in Array
        AvgPrice_value_index_in_Array = Order_Detail_Array_Length + AvgPrice_index_in_Array  # Get AvgPrice value's index number
        AvgPrice_i = round(AvgPrice_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_AvgPrice = array_from_copied_content[AvgPrice_i]  # Get AvgPrice Value using Index
        print("Expected AvgPrice is :" + str(Expected_AvgPrice))
        print("Actual AvgPrice is :" + str(Actual_AvgPrice))
        # ------------------------------------------------------Get Transaction Type Value in OrderDetails ------------------------------------------------#
        TransType_index_in_Array = array_from_copied_content.index("BuySell")  # Search string "BuySell" in Array
        TransType_value_index_in_Array = Order_Detail_Array_Length + TransType_index_in_Array  # Get TransType value's index number
        TransType_i = round(TransType_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_TransType = array_from_copied_content[TransType_i]  # Get AvgPrice Value using Index
        print("Expected TransType is :" + str(Expected_TransType))
        print("Actual TransType is :" + str(Actual_TransType))
        # ------------------------------------------------------Get Order Type Value in OrderDetails ------------------------------------------------#
        OrderType_index_in_Array = array_from_copied_content.index("OrdType")  # Search string "OrdType" in Array
        OrdType_value_index_in_Array = Order_Detail_Array_Length + OrderType_index_in_Array  # Get OrdType value's index number
        OrdType_i = round(OrdType_value_index_in_Array, 0)  # Truncade Decimal value
        Actual_OrdType = array_from_copied_content[OrdType_i]  # Get OrderType Value using Index
        print("Expected OrderType is :" + str(Expected_OrdType))
        print("Actual OrderType is :" + str(Actual_OrdType))
        # ------------------------------------------------------Compare OrderDetails ------------------------------------------------#
        if str(Expected_Quantity) == str(Actual_Qty) and str(Expected_Price) == str(Actual_Price) and str(Expected_OpenQty) == str(Actual_OpenQty) and str(Expected_TradedQty) == str(Actual_TradeQty) and str(Expected_AvgPrice) == str(Actual_AvgPrice) and str(Expected_TransType) == str(Actual_TransType) and str(Expected_TriggerPrice) == str(Actual_TriggerPrice) and str(Expected_DiscQty) == str(Actual_DiscQty) and str(Expected_OrdType) == str(Actual_OrdType):
            Test_Status = "PASS"
            print("TC is Pass")
            if str(TypeOfOrder) == "LMT":
                LG.Write_log_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradedQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType,Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty)
            if str(TypeOfOrder) == "SL-LMT":
                LG.Write_log_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradedQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status,Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "MKT":
                LG.Write_log_ActualResult_MKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradedQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status,Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "SL-MKT":
                LG.Write_log_ActualResult_SLMKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradedQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)

        else:
            Test_Status = "FAIL"
            print("TC is Fail")
            if str(TypeOfOrder) == "LMT":
                LG.Write_log_ActualResult_LMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradedQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice, Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty)
            if str(TypeOfOrder) == "SL-LMT":
                LG.Write_log_ActualResult_SLLMTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradedQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status,Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "MKT":
                LG.Write_log_ActualResult_MKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty,Expected_Quantity, Actual_Price, Expected_Price, Actual_OpenQty,Expected_OpenQty, Actual_TradeQty, Expected_TradedQty,Actual_AvgPrice, Expected_AvgPrice, Test_Status,Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)
            if str(TypeOfOrder) == "SL-MKT":
                LG.Write_log_ActualResult_SLMKTOrder(Result_path, TC_NO, row_value, Test_Scenario, Actual_Qty, Expected_Quantity,Actual_Price, Expected_Price, Actual_OpenQty, Expected_OpenQty,Actual_TradeQty, Expected_TradedQty, Actual_AvgPrice, Expected_AvgPrice,Test_Status, Expected_TransType, Actual_TransType, Expected_TriggerPrice,Actual_TriggerPrice,Expected_DiscQty,Actual_DiscQty,Expected_OrdType,Actual_OrdType)

        # -----------------------------------------------Clear Filter in OrderBook  ---------------------------------------------------------#
        keyboard.send_keys('^F')



