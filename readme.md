# OCR Documentation

## Overview
The OCR Python script is designed to capture specific button areas on a screen, save the captured images, and repeatedly click these areas in an application window. It uses Optical Character Recognition (OCR) to extract text from selected regions within the application and save the data to an Excel file.

## Dependencies
The script requires the following libraries:
- tkinter
- PIL (Pillow)
- pyautogui
- json
- os
- numpy
- cv2 (OpenCV)
- pygetwindow
- threading
- subprocess
- easyocr
- pandas
- pickle
- sys

## Main Components

### Class: `App`
The `App` class encapsulates the main functionality of the application.

#### Methods

##### `__init__(self, root, reader)`
Initializes the main application window, loads preferences, and either initializes the app or shows the preferences window.

##### `load_preferences(self)`
Loads user preferences from a JSON file.

##### `save_preferences(self)`
Saves user preferences to a JSON file.

##### `show_preferences_window(self)`
Displays a window for the user to set preferences such as parameters, delay time, and target application details.

##### `browse_file(self)`
Allows the user to browse and select the target application path.

##### `save_pref_and_initialize(self)`
Saves preferences and initializes the application.

##### `initialize_app(self)`
Checks if the target application is open, loads the button pathway, or allows the user to capture button areas if no pathway is found.

##### `check_application_open(self)`
Checks if the target application is currently open.

##### `open_application(self)`
Opens the target application using the specified path.

##### `capture_area(self)`
Captures a user-defined area on the screen and saves it as a selected image.

##### `save_pathway(self)`
Saves the captured button pathway to a JSON file.

##### `load_pathway(self)`
Loads the saved button pathway from a JSON file.

##### `execute_pathway(self)`
Activates the target application and clicks the saved button areas, then runs OCR on the application window.

##### `run_ocr_and_save(self)`
Captures the active window, allows the user to select regions of interest (ROIs) for OCR, and saves the extracted data to an Excel file.

##### `repeat_execution(self)`
Repeats the pathway execution at intervals specified by the delay time.

##### `pause_program(self)`
Pauses the program and allows the user to download OCR data.

##### `resume_program(self)`
Resumes the program and continues OCR operations.

##### `download_ocr_data(self)`
Downloads the OCR data as an Excel file.

##### `reset_data(self)`
Resets all saved data and preferences.

### Function: `main()`
Initializes the Tkinter root window and starts the application.

### Entry Point
The script starts execution from the `if __name__ == '__main__': main()` block.

## Usage
1. Run the script.
2. Set preferences in the preferences window if not already set.
3. Capture button areas in the target application.
4. Save the button pathway.
5. The application will repeatedly click the captured areas and perform OCR on the target application window.
6. Pause the program to download OCR data or reset the application.

## Error Handling
- Handles missing files gracefully by checking their existence before attempting to load or save data.
- Provides user feedback through message boxes for various operations such as saving preferences, capturing areas, and downloading data.

## Notes
- Ensure all dependencies are installed.
- The target application should be open and accessible for the script to function correctly.
- The script uses threading to handle repetitive tasks without freezing the GUI.
- For ease, use the desktop application to perform the OCR operation 


### Why EasyOCR?
EasyOCR is selected for its simplicity and effectiveness in performing OCR on images. It supports multiple languages and provides accurate text extraction capabilities, making it suitable for our application, which involves extracting text from specific regions within an application window.

### Why a GUI?
A graphical user interface (GUI) is created to provide a user-friendly way to interact with the application. The GUI allows users to:

- Set preferences such as parameters, delay time, and target application details.
- Capture specific button areas on the screen.
- Save and load button pathways.
- Start, pause, and resume the program.
- Download extracted OCR data.
- Reset all saved data and preferences.
- The GUI simplifies the interaction process and makes the application accessible to users without requiring them to modify the script directly.

### Efficiency Improvements
Several improvements are made to ensure the program runs efficiently:

- Multithreading: Uses threading to handle repetitive tasks, such as executing pathways and performing OCR, without freezing the GUI.
- Loading Preferences and Pathways: Preferences and pathways are saved and loaded from JSON files, allowing the application to remember user settings and previously captured button areas.
- OCR on Selected Regions: The user can define specific regions of interest (ROIs) for OCR, reducing the amount of unnecessary data processing and improving the accuracy of text extraction.
- Error Handling: The application handles missing files gracefully and provides user feedback through message boxes, ensuring a smooth user experience.