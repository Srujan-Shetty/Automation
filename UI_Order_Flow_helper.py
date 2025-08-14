from openpyxl import load_workbook
import os
import time
from pywinauto.keyboard import send_keys
#-------------Helper Function to Get the path of File and focus on sheet -----------------------#
def Get_File_Path(FileName,sheet_index):                          # Helper Function to Activate Files Sheet
    script_dir = os.path.dirname(os.path.abspath(__file__))       #Get the path of python Package where this code is saved
    file_name = str(FileName)                                     # Get the File name to be fetched
    file_path = os.path.join(script_dir, file_name)               # Get the Path of the File
    #print("credential Path" + str(file_path))
    TestCaseFile = load_workbook(file_path)                       #Load the workbook
    TestCaseFile._active_sheet_index = int(sheet_index)           # Activate the sheet index
    TestInfo_sheet = TestCaseFile.active                          #Activate the sheet
    return TestInfo_sheet                                         #Return the sheet access
#--------------------Helper Function to select value_to_select from combo box------------------#
def Select_From_DropDown(combo_box, value_to_select):
    try:
        # Click to expand the dropdown
        combo_box.click_input()
        time.sleep(0.5)  # Let it expand

        # Send first few letters or navigate with arrows
        send_keys(value_to_select)  # Type the name
        time.sleep(0.5)

        # Hit Enter to select
        send_keys('{ENTER}')
        print(f"Selected using keyboard: {value_to_select}")

    except Exception as e:
        print(f"Failed to select from custom combo box: {e}")