# 5. Calculator Automation
# - Perform 9 × 9
# - Validate the result is 81

from pywinauto import Application
import time

# Start Calculator and connect to it
app = Application(backend='uia').start("calc.exe")
calc = app.connect(title_re='^Calculator$', timeout=20)
time.sleep(2)

# Get the Calculator window
calc_window = calc.window(title_re='^Calculator$')

# Print control identifiers (optional, helps to confirm names)
calc_window.print_control_identifiers()

# Click buttons: 9*9 =
calc_window.child_window(title="Nine", auto_id="num9Button", control_type="Button").click_input()
calc_window.child_window(title="Multiply by", auto_id="multiplyButton", control_type="Button").click_input()
calc_window.child_window(title="Nine", auto_id="num9Button", control_type="Button").click_input()
calc_window.child_window(title="Equals", auto_id="equalButton", control_type="Button").click_input()

Expected_result = 81
result_text = calc_window.child_window(auto_id="CalculatorResults", control_type="Text").window_text()

if str(Expected_result) in result_text:
    print("Test Passed")
else:
    print("Test Failed")

# Optional: Wait and then close
time.sleep(3)
app.kill()


