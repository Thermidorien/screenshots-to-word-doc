import time
import pyautogui
from pywinauto import application

# Path to the Word document
word_doc_path = "path_to_your_word_document.docx"

# Take a screenshot using Snipping Tool
time.sleep(2)  # Wait for 2 seconds before triggering the Snipping Tool
pyautogui.hotkey("winleft", "shift", "s")  # Trigger the Snipping Tool shortcut

# Wait for the Snipping Tool to open
time.sleep(2)

# Assuming you are using the rectangular snip mode
pyautogui.press("down")  # Select rectangular snip mode
pyautogui.press("enter")  # Take the snip

# Wait for the screenshot to be captured
time.sleep(2)

# Save the screenshot to a file
screenshot_path = "path_to_save_screenshot.png"
pyautogui.hotkey("ctrl", "s")  # Save the snip
pyautogui.typewrite(screenshot_path)
pyautogui.press("enter")  # Save the screenshot with the given path

# Open the Word document using pywinauto
app = application.Application()
app.start("WINWORD.EXE {}".format(word_doc_path))
time.sleep(2)  # Wait for Word to open

# Paste the screenshot into the Word document
pyautogui.hotkey("ctrl", "v")

# Save and close the Word document
time.sleep(1)
app.Word.ActiveDocument.Save()
app.Word.ActiveDocument.Close()

# Optional: If you want to quit Word after saving the document
app.quit()
