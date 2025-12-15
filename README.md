# Automate_manual_program_work
This is a code automate manual process downloading training images from ImageViewer.exe provided from FamilySearch organization.

# `pyautogui`: GUI Automation Library

`pyautogui` is a cross-platform library that provides functions for automating mouse and keyboard actions. It allows you to control the mouse, keyboard, and perform various screen-related tasks such as taking screenshots and locating specific images on the screen.

## Key Features of pyautogui:

### Mouse Control:

Move the mouse to specific coordinates on the screen.

Click, double-click, or right-click.

Scroll the mouse wheel.

### Keyboard Control:

Type text or individual keys.

Press combinations of keys (like Ctrl+C).

### Screen Recognition:

Take screenshots.

Locate an image on the screen and get its position.

### Fail-safes:

Automatic safety feature that stops the program if the mouse is moved to a corner of the screen.

# `pywinauto`: Windows GUI Automation Library

`pywinauto` is a Python library that allows you to automate and control Windows applications. It interacts with native Windows GUI elements such as buttons, text boxes, and windows. It is much more suited for Windows-specific applications than pyautogui.

## Key Features of pywinauto:

### Window Management:

Bring windows to the foreground, minimize, or maximize them.

Interact with window elements using window titles, classes, or control IDs.

### GUI Interaction:

Automate clicks, typing, and drag-and-drop operations.

Set values for form elements (text boxes, checkboxes, radio buttons).

### Native Controls:

Works with native Windows controls (such as buttons, text fields, etc.) directly.

Unlike pyautogui, it can interact with Windows-based GUI elements by their class names or properties.

### Handling Dialogs and Pop-ups:

Automatically interact with Windows dialogs, such as pressing "OK" on a message box.
