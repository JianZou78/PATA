# -*- coding: utf-8 -*-
"""
Created on Fri Mar  7 19:20:25 2025

@author: jianzou
"""

import pygetwindow as gw
import pyautogui
import time
import subprocess
import tkinter as tk
from tkinter import filedialog
import os
import shutil
import stat
import wexpect
from datetime import datetime
import pyperclip
 

def get_input(prompt, default_value):
    user_input = input(f"{prompt} (default: {default_value}): ")
    return user_input if user_input else default_value

call_to_address = get_input('Please input email address to be called, for example: sox_f65_3@soxlabcn.onmicrosoft.com\n','sox_f65_3@soxlabcn.onmicrosoft.com')

scrcpy_win_name = get_input('Please input scrcpy window title, for example: MeetingBoard 65 Pro\n','MeetingBoard 65 Pro')

call_Duration =int( get_input('Please input call duration expected in seconds , for example:10\n','10'))
dut_ipaddr = get_input('Please input the ip address of the Andriod DUT, for example: 10.172.206.65:5555\n','10.172.206.65:5555')
loopN = get_input('Please input the test loop, for example: 3\n','3')

def select_MPAT_folder():
    
    print (" select the MPAT folder path....\n" )
    root = tk.Tk()
    
    root.withdraw()
     
    # Open a folder dialog and get the selected folder path
    
    MPAT_path = filedialog.askdirectory()
     
    # Print the selected folder path
    
    print(f"Selected folder: {MPAT_path}")
    return MPAT_path

def select_ResultSave_folder():
    
    print (" select the ResultSave path....\n" )
    root = tk.Tk()
    
    root.withdraw()
     
    # Open a folder dialog and get the selected folder path
    
    ResultSave_path = filedialog.askdirectory()
     
    # Print the selected folder path
    
    print(f"Selected folder: {ResultSave_path}")
    return ResultSave_path
def find_window_containing_string(search_string):
    while True:
        # Get a list of all window titles
        windows = gw.getAllTitles()

        # Filter windows that contain the search string in their title
        matching_windows = [title for title in windows if search_string.lower() in title.lower()]

        if matching_windows:
            # Get the window object of the first matching window
            window = gw.getWindowsWithTitle(matching_windows[0])[0]
            print(f"Window ID: {window._hWnd}, Title: {window.title}")
            break

        # Wait for a short period before checking again
        time.sleep(1)
    return window

def open_cmd_and_get_window_id(directory):
    # Open the command prompt and change to the given directory
    subprocess.Popen(f'start cmd /K "cd /d {directory}"', shell=True)
    
    windowsList= find_window_containing_string('C:\WINDOWS\system32\cmd.exe')

    #windowsList = gw.getWindowsWithTitle('C:\WINDOWS\system32\cmd.exe')
    
    return windowsList
      


#MPAT_Android_Test()


def auto_WinDesk_call_MTRA():
    
    start_time = time.time()
           
    # Get all windows
    windows = gw.getAllTitles()
    
    # Print all window titles
    for title in windows:
        if 'Microsoft Teams' in title:
            print(f"Teams Window found in {title}\n")
    
    # Find the window by title
    window_list = gw.getWindowsWithTitle('Microsoft Teams')
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        
        print("Teams windows found\n")
        # Activate the window
        window.maximize()
        window.activate()
    else:
        print("No window with the title 'Microsoft Teams' found.")
    
    
    
    time.sleep(1)
    # Send Ctrl+N, start a new chat window
    pyautogui.hotkey('ctrl', 'shift','N')
    
    time.sleep(2)
    pyautogui.typewrite(call_to_address)
    time.sleep(2)
    pyautogui.press('enter')
    # time.sleep(1)
    # pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    
    time.sleep(1)
    # Send Alt+Shift+V, start video call
    pyautogui.hotkey('alt', 'shift', 'v')
    
    time.sleep(10)
    
    # Find the window by title
    window_list = gw.getWindowsWithTitle(scrcpy_win_name)
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        
        print(f" window with the title {scrcpy_win_name} found....\n")
        # Activate the window
        window.maximize()
        window.activate()
    else:
        print(f"No window with the title {scrcpy_win_name} found.")
    #===============================================================================start Call==============================
    time.sleep(1)
    # Define the PowerShell command you want to run
    command = "adb shell input keyevent KEYCODE_CALL"
    
    # Use subprocess to run the PowerShell command
    result = subprocess.run(["powershell", "-Command", command], capture_output=True, text=True)
    
    # Print the output of the command
    print(result.stdout)
    
    
    #===============================================================================End Call==============================
    time.sleep(call_Duration)  # call duration
    
    # Find the window by title
    window_list = gw.getWindowsWithTitle(scrcpy_win_name)
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        
        print(f"\n window with the title {scrcpy_win_name} found....\n")
        # Activate the window
        window.maximize()
        window.activate()
    else:
        print(f"No window with the title {scrcpy_win_name} found.")
    # Define the PowerShell command you want to run
    command = "adb shell input keyevent KEYCODE_ENDCALL"
    
    # Use subprocess to run the PowerShell command
    result = subprocess.run(["powershell", "-Command", command], capture_output=True, text=True)
    
    # Print the output of the command
    print(result.stdout)
    
    # Find the window by title
    window_list = gw.getWindowsWithTitle('Microsoft Teams')
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        
        print("Teams windows found\n")
        # Activate the window
        window.maximize()
        time.sleep(1)
        window.activate()
    else:
        print("No window with the title 'Microsoft Teams' found.")
        
    time.sleep(3)
    # Send Ctrl+Shift+H, end call
    pyautogui.hotkey('ctrl', 'shift', 'H')
    
    end_time = time.time()
    
    duration = int(end_time - start_time +3)
    
    print(f"Call and exectuion time around {duration} seconds...\n")
    return duration
######

####### ====================================================================
# Example usage for MPAT upload
def empty_folder(folder_path):
    # Check if the folder exists
    if not os.path.exists(folder_path):
        print(f"The folder '{folder_path}' does not exist.")
        return
    
    # Change permissions and empty the folder by deleting all its contents
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        os.chmod(item_path, stat.S_IWRITE)  # Change the permissions to writable
        if os.path.isdir(item_path):
            shutil.rmtree(item_path)
        else:
            os.remove(item_path)
    
    print(f"All contents of the folder '{folder_path}' have been deleted.")

    
def copy_folder_contents(source_folder, destination_folder):
    # Check if the source folder exists
    if not os.path.exists(source_folder):
        print(f"The source folder '{source_folder}' does not exist.")
        return
    
    # Create the destination folder if it does not exist
    os.makedirs(destination_folder, exist_ok=True)
    
    # Copy the contents of the source folder to the destination folder
    for item in os.listdir(source_folder):
        source_item = os.path.join(source_folder, item)
        destination_item = os.path.join(destination_folder, item)
        if os.path.isdir(source_item):
            shutil.copytree(source_item, destination_item)
        else:
            shutil.copy2(source_item, destination_item)
    
    print(f"All contents from '{source_folder}' have been copied to '{destination_folder}'.")

def change_permissions_and_delete(folder_path):
    # Change the permissions of the folder and its contents
    for root, dirs, files in os.walk(folder_path):
        for dir in dirs:
            os.chmod(os.path.join(root, dir), 0o777)
        for file in files:
            os.chmod(os.path.join(root, file), 0o777)
    # Delete the folder and its contents
    shutil.rmtree(folder_path)
    print(f"Deleted folder: {folder_path}")

def delete_subfolders_with_prefix(directory, prefix):
    # List all subfolders in the given directory
    for subfolder in os.listdir(directory):
        subfolder_path = os.path.join(directory, subfolder)
        # Check if the subfolder name starts with the given prefix and is a directory
        if os.path.isdir(subfolder_path) and subfolder.startswith(prefix):
            try:
                change_permissions_and_delete(subfolder_path)
            except PermissionError as e:
                print(f"PermissionError: {e}")
            except Exception as e:
                print(f"Error: {e}")



def MPAT_Android_Test(ipaddr='10.172.206.65:5555', mpat_path=r'C:\Users\jianzou\OneDrive - Microsoft\MPAT_1.2.1',loop=1):
    
    result_file = os.path.join(mpat_path,"result_url.txt")
    
    delete_subfolders_with_prefix(mpat_path,"result")
    # Empty the existing folder by deleting all its contents
   

   

    
    for ii in range(0,int(loop)):
        window_id = open_cmd_and_get_window_id(mpat_path)
        
        # existing_folder = os.path.join(mpat_path,"result")
        # if os.path.exists(existing_folder) and os.path.isdir(existing_folder):
        #     change_permissions_and_delete(existing_folder)
        
        
        if window_id:
            print(f"Command Prompt window ID: {window_id}\n")
            cmd_test = 'AudioToolClient.exe test --id ' + ipaddr + ' --app com.microsoft.skype.teams.ipphone --path ./result_'+str(ii)
            #window_id.maximize()
            window_id.activate()
            time.sleep(1)
            pyautogui.typewrite(cmd_test)
            time.sleep(1)
            
            pyautogui.press('enter')
            while True:
                # Send Ctrl +A copy all  the contents to clipboard
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(1)
                # Get the content from the clipboard
                clipboard_content = pyperclip.paste()
                if "Testing is about to begin" not in clipboard_content:
                    time.sleep(1)
                else:
                    break
            pyautogui.typewrite('y')
            
            duration = auto_WinDesk_call_MTRA()
            
            
            #window_id.maximize()
            window_id.activate()
            time.sleep(1)
            pyautogui.press('enter')
            while True:
                # Send Ctrl +A copy all  the contents to clipboard
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(1)
                # Get the content from the clipboard
                clipboard_content = pyperclip.paste()
                if "The test results will be uploaded" not in clipboard_content:
                    time.sleep(1)
                else:
                    break#time.sleep(10)

                
            pyautogui.typewrite('y')
            while window_id.isActive:
                time.sleep(3)
                print('waiting...\n')
            
            #window_id.maximize()
            window_id.activate()
            
            time.sleep(3)

            
            # Send Ctrl +A copy all  the contents to clipboard
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(1)
            # Get the content from the clipboard
            clipboard_content = pyperclip.paste()

            lines =  clipboard_content.splitlines(); 


            # Save the clipboard content to a text file
            # Check if the file already exists
            if os.path.exists(result_file):
                # Append the clipboard content to the file
                with open(result_file, "a") as file:
                    file.write(f"Test {ii}\n")
                    file.write(lines[-3] + "\n")
            else:
                # Create a new file and write the clipboard content
                with open(result_file, "w") as file:
                    file.write(f"Test {ii}\n")
                    file.write(lines[-3] + "\n")
                    
            time.sleep(1)
            
            window_id.close()
        
            # new_folder = os.path.join(mpat_path,'result_'+str(ii))

            # copy_folder_contents(existing_folder,new_folder)

                 
        else:
            print("Command Prompt window not found.")

    return result_file

def open_notepad_with_file(file_name):
    # Check if the file exists
    if os.path.exists(file_name):
        # Open Notepad with the specified file
        subprocess.run(['notepad.exe', file_name])
    else:
        print(f"The file '{file_name}' does not exist.")

if __name__ == "__main__":
    
    MPAT_path = select_MPAT_folder()
    result_file = MPAT_Android_Test(ipaddr=dut_ipaddr,mpat_path= MPAT_path, loop=loopN)
###### Display the result ==================================================================== 
    open_notepad_with_file(result_file) 