## Knowledge Transfer: OCR_Script
### Objective
The objective of this script is to automate the interaction with a specified application by capturing specific button areas on the screen, saving these button areas, and clicking them repeatedly. Additionally, it performs Optical Character Recognition (OCR) on selected regions within the application window and saves the extracted text data to an Excel file.

### Components Overview
The script consists of the following main components:
1. **Imports**: Necessary libraries and modules.
2. **Class: `App`**: The core of the application, which handles the main functionality.
3. **Function: `main()`**: Initializes the Tkinter root window and starts the application.
4. **Entry Point**: The script execution starts from the `if __name__ == '__main__': main()` block.

### Detailed Breakdown

#### Imports
The script imports several libraries for various functionalities:
- **tkinter**: For creating the graphical user interface (GUI).
- **PIL (Pillow)**: For image processing tasks.
- **pyautogui**: For capturing screenshots and automating mouse clicks.
- **json**: For saving and loading preferences and pathways.
- **os**: For interacting with the operating system.
- **numpy**: For handling image data.
- **cv2 (OpenCV)**: For image processing and handling image formats.
- **pygetwindow**: For interacting with application windows.
- **threading**: For handling concurrent execution.
- **subprocess**: For opening the target application.
- **easyocr**: For performing OCR on captured regions.
- **pandas**: For managing and saving extracted data.
- **pickle**: For saving and loading regions of interest (ROIs).
- **sys**: For interacting with the Python runtime environment.

#### Class: `App`
The `App` class encapsulates the main functionality of the application, handling user interactions, capturing button areas, saving/loading preferences, performing OCR, and automating button clicks.

##### Initialization (`__init__`)
- **Parameters**: Initializes the main application window (`root`) and the OCR reader (`reader`).
- **Preferences**: Loads user preferences if available, otherwise prompts the user to set preferences.
- **Attributes**: Sets various attributes such as `selected_images`, `pathway_file`, `preferences_file`, `rois_file`, `parameters`, `delay_time`, `target_application_path`, `target_application_name`, and `running_event`.

##### Loading and Saving Preferences
- **`load_preferences`**: Loads user preferences from a JSON file.
- **`save_preferences`**: Saves user preferences to a JSON file.

##### GUI for Setting Preferences
- **`show_preferences_window`**: Displays a window for the user to set preferences such as parameters, delay time, and target application details.
- **`browse_file`**: Allows the user to browse and select the target application path.
- **`save_pref_and_initialize`**: Saves preferences and initializes the application.

##### Application Initialization
- **`initialize_app`**: Checks if the target application is open, loads the button pathway, or allows the user to capture button areas if no pathway is found.
- **`check_application_open`**: Checks if the target application is currently open.
- **`open_application`**: Opens the target application using the specified path.

##### Capturing Button Areas
- **`capture_area`**: Captures a user-defined area on the screen and saves it as a selected image.
- **`save_pathway`**: Saves the captured button pathway to a JSON file.
- **`load_pathway`**: Loads the saved button pathway from a JSON file.

##### Executing Pathways
- **`execute_pathway`**: Activates the target application and clicks the saved button areas, then runs OCR on the application window.

##### Performing OCR
- **`run_ocr_and_save`**: Captures the active window, allows the user to select regions of interest (ROIs) for OCR, and saves the extracted data to an Excel file.

##### Repeated Execution
- **`repeat_execution`**: Repeats the pathway execution at intervals specified by the delay time.

##### Pausing and Resuming
- **`pause_program`**: Pauses the program and allows the user to download OCR data.
- **`resume_program`**: Resumes the program and continues OCR operations.

##### Downloading and Resetting Data
- **`download_ocr_data`**: Downloads the OCR data as an Excel file.
- **`reset_data`**: Resets all saved data and preferences.

#### Function: `main()`
The `main()` function initializes the Tkinter root window and starts the application by creating an instance of the `App` class and calling its `mainloop()` method.

#### Entry Point
The script execution starts from the `if __name__ == '__main__': main()` block, ensuring that the `main()` function is called only if the script is run directly.

### Detailed Walkthrough

1. **Initialization**: 
   - When the script runs, it initializes the main Tkinter window (`root`) and creates an instance of the `App` class.
   - The `App` class constructor (`__init__`) attempts to load user preferences from a JSON file. If preferences are found, it initializes the application with these preferences. If not, it prompts the user to set preferences through a GUI.

2. **Setting Preferences**:
   - The user sets parameters, delay time, target application name, and path through a preferences window.
   - Preferences are saved to a JSON file for future use.

3. **Capturing Button Areas**:
   - The user captures button areas by selecting regions on the screen. These areas are saved as images.
   - The captured images are saved as a button pathway in a JSON file.

4. **Executing Pathways**:
   - The application loads the button pathway and repeatedly clicks the saved button areas in the target application.
   - The script uses PyAutoGUI for simulating mouse clicks on the captured button areas.

5. **Performing OCR**:
   - After clicking the button areas, the script performs OCR on specified regions within the application window.
   - The user can define regions of interest (ROIs) for OCR, which are saved for future use.
   - Extracted text data is saved to an Excel file, with new data appended if the file already exists.

6. **Pausing and Resuming**:
   - The user can pause the program to download OCR data or reset the application.
   - The program can be resumed to continue OCR operations.

7. **Downloading and Resetting Data**:
   - OCR data can be downloaded as an Excel file.
   - The user can reset all saved data and preferences, removing the JSON and pickle files.

### Efficiency Considerations
- **Multithreading**: Uses threading to handle repetitive tasks without freezing the GUI.
- **Saving and Loading**: Preferences and pathways are saved and loaded from JSON files, allowing the application to remember user settings and previously captured button areas.
- **OCR Efficiency**: The user defines specific ROIs for OCR, reducing unnecessary data processing and improving text extraction accuracy.
- **Error Handling**: The application handles missing files gracefully and provides user feedback through message boxes, ensuring a smooth user experience.


### Setting Up the Script on your System

To set up and run the "OCR_Script" on your system, follow the detailed instructions below:

#### Prerequisites
1. **Python**: Ensure that Python 3.6 or later is installed on your system. You can download the latest version of Python from [python.org](https://www.python.org/downloads/).

2. **Pip**: Python's package installer should be available. It usually comes with Python, but you can install it separately if needed by following the instructions [here](https://pip.pypa.io/en/stable/installation/).

3. **Required Libraries**: The script uses several external libraries. Install these using `pip`.

#### Step-by-Step Setup Instructions

1. **Clone or Download the Script**:
   - Clone the repository or download the script file to your local machine.

2. **Install Required Libraries**:
   Open a terminal or command prompt and navigate to the directory where the script is located. Run the following command to install all the required libraries:
   ```bash
   pip install tkinter Pillow pyautogui numpy opencv-python-headless pygetwindow easyocr pandas
   ```

3. **Set Up EasyOCR**:
   - EasyOCR requires additional data files to run. These will be downloaded automatically when you run the script for the first time.
   - Ensure you have an active internet connection during the first run to allow EasyOCR to download necessary files.

4. **Running the Script**:
   - Once all the dependencies are installed, you can run the script by executing the following command in the terminal or command prompt:
     ```bash
     python your_script_name.py
     ```
     Replace `your_script_name.py` with the actual name of your script file.

5. **Using the Application**:
   - The script will open a Tkinter window prompting you to set preferences if they are not already saved.
   - Follow the instructions provided by the application to capture button areas, save pathways, and perform OCR.

6. **Setting Up the Target Application**:
   - Ensure the target application you want to automate is installed on your system.
   - Provide the correct path and name of the target application in the preferences window of the script.

### Additional Details

#### Configuring the Script
- **Preferences**:
  - Set the parameters (comma-separated) that you want to monitor or extract using OCR.
  - Set the delay time in seconds for how often the script should repeat the pathway execution.
  - Provide the target application name and path for the application you wish to automate.

#### Capturing Button Areas
- **Capture Button Area**: Use the "Capture Button Area" button to select regions on the screen. These regions will be saved and clicked automatically during each execution cycle.

#### Saving and Loading Pathways
- **Save Pathway**: After capturing the button areas, save the pathway by clicking the "Save Pathway" button.
- **Load Pathway**: The script automatically loads the saved pathway when you run it the next time.

#### Performing OCR
- **Selecting ROIs**: When prompted, select the regions of interest (ROIs) for OCR. These regions are saved for future use.
- **OCR Execution**: The script captures the active window and performs OCR on the defined ROIs, saving the extracted data to an Excel file.

### Troubleshooting

1. **Module Not Found Error**:
   - If you encounter a "ModuleNotFoundError", ensure that all the required libraries are installed. Re-run the `pip install` command if necessary.

2. **Application Not Opening**:
   - Ensure that the path to the target application is correct. Check for any typos or incorrect paths in the preferences.

3. **OCR Accuracy**:
   - If OCR results are not accurate, adjust the regions of interest (ROIs) and ensure that the captured images are clear and well-defined.

By following these setup instructions, you should be able to run the "OCR_Script" on your system successfully. This script provides a robust solution for automating interactions with applications and performing OCR on specific regions, making it a valuable tool for repetitive tasks and data extraction.



### Conclusion
This script automates interactions with a specified application by capturing and clicking button areas and performing OCR on selected regions. The use of a GUI makes the application user-friendly, while efficient data handling and multithreading ensure smooth operation.