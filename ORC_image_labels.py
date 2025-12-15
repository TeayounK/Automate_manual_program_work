from pywinauto import Application
from pywinauto.findwindows import find_windows
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys
import psutil
import time
import os


IMAGE_PATH = "C:\FamilySearch\ImageViewer\TrainingImages"
filepath = [im for im in os.listdir(IMAGE_PATH)]

# Step 1: Connect to the already-open Chrome window
windows = find_windows(title_re=".*Document Recognition Suite.*")
if not windows:
    raise Exception("No matching Chrome window found")

app = Application(backend="uia").connect(handle=windows[0])
window = app.window(handle=windows[0])

for f in filepath:
    # Find all buttons in the window
    buttons = window.descendants(control_type="Button")

    print("Found buttons:")
    for b in buttons:
        # print(f"Title: '{b.window_text()}', AutomationId: '{b.automation_id()}', Class: '{b.class_name()}'")
        if b.window_text() == "Select Document":
            print("Clicking 'Select Document' button...")
            b.click_input()
            break

    time.sleep(1)

    edit_box = window.child_window(control_type="Edit")

    print(f"Typing file path: {f}")
    send_keys(f)
    send_keys('{ENTER}')

    proc = psutil.Process(window.process_id())

    completion = True
    while completion:
        cpu = proc.cpu_percent(interval=1)
        if cpu < 0.05: # assume idle if below 2%
            HyperLink = window.descendants(control_type="Hyperlink")
            for b in HyperLink:
                if b.window_text() == "Raw Data (JSON)":
                    print("Loading complete.")
                    completion = False
        print(f"Waiting... CPU {cpu:.1f}%")


    HyperLink = window.descendants(control_type="Hyperlink")

    print("Found Hyperlinks:")
    for b in HyperLink:
        print(f"Title: '{b.window_text()}', AutomationId: '{b.automation_id()}', Class: '{b.class_name()}'")
        if b.window_text() == "Raw Data (JSON)":
            print("Clicking 'Raw Data (JSON)' button...")
            b.click_input()
            break
    time.sleep(1)
    send_keys('{ESC}')


