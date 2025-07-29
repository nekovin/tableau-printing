import pyautogui
import time

print(pyautogui.position())

# Select all 4 pages (Page 1 to Page 4) using Shift+Click
print("Selecting all 4 pages...")
pyautogui.click(946, 1397)  # Click Page 1 tab
time.sleep(0.5)
pyautogui.keyDown('shift')  # Hold Shift
pyautogui.click(1165, 1402)  # Click Page 4 tab while holding Shift
pyautogui.keyUp('shift')  # Release Shift
time.sleep(5)

# Click on "Entire workbook" radio button
pyautogui.click(22, 40)
time.sleep(5)

# Click on "View PDF after printing" checkbox  
pyautogui.click(63, 438) 
time.sleep(5)

# Select "Selected sheets" radio button
pyautogui.click(520, 721)  # Updated to your correct coordinates
time.sleep(5)

# Click OK button
pyautogui.click(667, 812)
time.sleep(5)

# Type the PDF filename
pyautogui.write("Auburn - Walkin.pdf")

# Click Save button
pyautogui.click(1533, 1303)
time.sleep(5)