import os
import subprocess
import pyautogui
import time

path = r"C:\UFT\JULY\JUL 18\1.0.8_18JUL_Portable\NT_Release_1.0.8_18JUL_Portable\Release"
os.chdir(path)

exe_name = "NorenTrader.exe"
subprocess.Popen(exe_name)


# Delay to give you time to open the login page
time.sleep(8)

# Fill in the coordinates below after using pyautogui.position()                    
username_field = (865, 477)
pyautogui.press('tab', presses=5)

# Select all text and clear the username field
pyautogui.hotkey('ctrl', 'a')     
pyautogui.press('backspace') 


# Move to username field and type username
pyautogui.click(username_field)
pyautogui.write("SRUJAN", interval=0.1)
pyautogui.press('tab')

# Move to password field and type password
pyautogui.write("Srujan@1234", interval=0.1)


# Click login button
pyautogui.press('enter')
time.sleep(2)
# Click on the 'OK' button in the login dialog
pyautogui.press('enter')
time.sleep(3)
#click on the continue button
pyautogui.press('enter')





