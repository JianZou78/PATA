Overview
This Python script automates the process of making a call using Microsoft Teams, interacting with an Android device via ADB, and managing test results. It includes functionalities for selecting folders, finding windows by title, and executing commands in the command prompt.
Dependencies
Before running the script, ensure you have the following Python libraries installed:
• pygetwindow
• pyautogui
• tkinter
• subprocess
• wexpect
• datetime
• pyperclip
• shutil
• os
• stat
• time
You can install the required libraries using pip:
Script Usage
1. Get User Input: The script prompts the user for the following inputs:
• Email address to be called.
• Scrcpy window title.
• Call duration in seconds.
• IP address of the Android DUT.
• Number of test loops.
2. Select Folders: The script provides functions to select the MPAT folder and the ResultSave folder using a folder dialog.
3. Find Window by Title: The script includes a function to continuously loop until it finds a window containing a given string in its title.
4. Open Command Prompt and Get Window ID: The script opens a command prompt in a specified directory and retrieves the window ID.
5. Automate Microsoft Teams Call: The script automates the process of making a call using Microsoft Teams, including starting a new chat, typing the email address, and initiating a video call.
6. Manage Test Results: The script includes functions to empty a folder, copy folder contents, and delete subfolders with a specific prefix.
7. MPAT Android Test: The script performs the MPAT Android test by executing commands in the command prompt, making a call, and saving the test results to a text file.
8. Open Notepad with Result File: The script opens Notepad with the result file after the test is completed.
Example Usage
1. Run the Script: Execute the script in your Python environment:
2. Provide User Input: Follow the prompts to provide the required input values.
3. Select Folders: Use the folder dialog to select the MPAT folder and the ResultSave folder.
4. Automate the Test: The script will automate the test process, manage the test results, and open the result file in Notepad.
Functions
• get_input(prompt, default_value): Prompts the user for input with a default value.
• select_MPAT_folder(): Opens a folder dialog to select the MPAT folder.
• select_ResultSave_folder(): Opens a folder dialog to select the ResultSave folder.
• find_window_containing_string(search_string): Continuously loops until it finds a window containing the given string in its title.
• open_cmd_and_get_window_id(directory): Opens a command prompt in the specified directory and retrieves the window ID.
• auto_WinDesk_call_MTRA(): Automates the process of making a call using Microsoft Teams.
• empty_folder(folder_path): Empties the specified folder by deleting all its contents.
• copy_folder_contents(source_folder, destination_folder): Copies the contents of the source folder to the destination folder.
• change_permissions_and_delete(folder_path): Changes the permissions of the folder and its contents, then deletes the folder.
• delete_subfolders_with_prefix(directory, prefix): Deletes subfolders with the specified prefix in the given directory.
• MPAT_Android_Test(ipaddr, mpat_path, loop): Performs the MPAT Android test and saves the results to a text file.
• open_notepad_with_file(file_name): Opens Notepad with the specified file.
Conclusion
This script automates the process of making a call using Microsoft Teams, interacting with an Android device, and managing test results. Follow the instructions to set up the environment and run the script. If you encounter any issues or need further assistance, please refer to the function descriptions and example usage provided in this guide.![image](https://github.com/user-attachments/assets/b19c63e4-f415-45bb-a70c-46b72a59b643)
