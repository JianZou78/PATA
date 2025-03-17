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
import win32gui
 



def get_input(prompt, default_value):
    user_input = input(f"{prompt} (default: {default_value}): ")
    return user_input if user_input else default_value
PATA_version = r"Release0 Version: March 17, 2025"    
print("\n\n---------------------------------PATA_Python_Auto_Teams_Android MPAT------------------------------------------------------\n")
print('Tool for development testing only. Provided "as is" with no guarantee of compatibility with your test environment.\n\n')
print('Prerequisite\n')
print('Install MPAT on the x64 Windows PC. \n')
print('Ensure Windows Teams is installed and can call DUT Android Teams \n')
print('Install Scrcpy on the Windows PC and can scrcpy to the DUT \n')
print('DUT Android Teams devices adb enabled and accessible from the Windows PC \n\n')
print(f"------------------------------------------{PATA_version}----------------------------------------------------\n\n")



call_to_address = get_input('Please input email address to be called, for example: sox_f65_12@soxlabcn.onmicrosoft.com\n','sox_f65_12@soxlabcn.onmicrosoft.com')

call_audio_video = get_input('audio call or video call , for example: video\n','video')

call_to_win_name = get_input('Please input call to window name keywords, for example: F65_12\n','F65_12')

inCall_win_name = get_input('Please input incall window name keywords, for example: Microsoft Teams meeting\n','Microsoft Teams meeting')

scrcpy_win_name = get_input('Please input scrcpy window title, for example: XBar V50\n','XBar V50')

call_Duration =int( get_input('Please input call duration expected in seconds , for example:10\n','10'))
dut_ipaddr = get_input('Please input the ip address of the Andriod DUT, for example: 10.172.206.175:5555\n','10.172.206.175:5555')
loopN = int(get_input('Please input the test loop, for example: 3\n','3'))


def window_exists_exact(title):
    def enum_windows_callback(hwnd, titles):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            if title in window_text:
                titles.append(hwnd)

    titles = []
    win32gui.EnumWindows(enum_windows_callback, titles)
    return len(titles) > 0


def scrcpy_run(scrcpyExePath = r"C:\scrcpy-win64-v2.1"):
    
    currPath = os.getcwd()
    
    #scrcpyExePath = r"C:\scrcpy-win64-v2.1"
    os.chdir(scrcpyExePath)
    
    subprocess.run(["powershell", "-Command", "Start-Process powershell"])

    time.sleep(3)
    
    cmdline1 = '.\scrcpy.exe -m 1920'
     
    
    commands = cmdline1;
     
    window_list = gw.getWindowsWithTitle('PowerShell')
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        window.activate()
        time.sleep(1)
        pyautogui.typewrite(commands)
        pyautogui.press('enter')

        #os.chdir(currPath);
        time.sleep(2)
        window.close()
    else:
        print("No window with the title 'PowerShell' found.")
        
    os.chdir(currPath)
        
    # window_list = gw.getWindowsWithTitle(scrcpy_win_name)
    
    # # Check if any windows were found
    # if window_list:
    #     # Select the first window from the list
    #     window = window_list[0]

    #     window.close()
    # else:
    #     print(f"No window with the title {scrcpy_win_name} found.\n")
        
    
def scrcpy_PS_accept(scrcpyExePath = r"C:\scrcpy-win64-v2.1"):
    
    currPath = os.getcwd()
    
    #scrcpyExePath = r"C:\scrcpy-win64-v2.1"
    os.chdir(scrcpyExePath)
    
    subprocess.run(["powershell", "-Command", "Start-Process powershell"])

    time.sleep(3)
    
    cmdline1 = 'adb shell input keyevent KEYCODE_CALL'
     
    
    commands = cmdline1;
     
    window_list = gw.getWindowsWithTitle('PowerShell')
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        window.activate()
        time.sleep(1)
        pyautogui.typewrite(commands)
        pyautogui.press('enter')

        #os.chdir(currPath);
        time.sleep(1)
        window.close()
    else:
        print("No window with the title 'PowerShell' found.")
        
    window_list = gw.getWindowsWithTitle(scrcpy_win_name)
    
    os.chdir(currPath)
    
    # # Check if any windows were found
    # if window_list:
    #     # Select the first window from the list
    #     window = window_list[0]

    #     window.close()
    # else:
    #     print(f"No window with the title {scrcpy_win_name} found.\n")    
        
    
def scrcpy_PS_end(scrcpyExePath = r"C:\scrcpy-win64-v2.1"):
    
    currPath = os.getcwd()
    
    #scrcpyExePath = r"C:\scrcpy-win64-v2.1"
    os.chdir(scrcpyExePath)
    
    subprocess.run(["powershell", "-Command", "Start-Process powershell"])

    time.sleep(3)
    
    cmdline1 = 'adb shell input keyevent KEYCODE_ENDCALL'
     
    
    commands = cmdline1;
     
    window_list = gw.getWindowsWithTitle('PowerShell')
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        window.activate()
        time.sleep(1)
        pyautogui.typewrite(commands)
        pyautogui.press('enter')

        #os.chdir(currPath);
        time.sleep(1)
        window.close()
    else:
        print("No window with the title 'PowerShell' found.")
        
    os.chdir(currPath)
    # window_list = gw.getWindowsWithTitle(scrcpy_win_name)
    
    # # Check if any windows were found
    # if window_list:
    #     # Select the first window from the list
    #     window = window_list[0]

    #     window.close()
    # else:
    #     print(f"No window with the title {scrcpy_win_name} found.\n")    

def select_Scrcpy_folder():
    
    print (" select the Scrcpy folder path....\n" )
    root = tk.Tk()
    
    root.withdraw()
     
    # Open a folder dialog and get the selected folder path
    
    Scrcpy_path = filedialog.askdirectory()
     
    # Print the selected folder path
    
    print(f"Selected folder: {Scrcpy_path}")
    return Scrcpy_path
    

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
    window_list = gw.getWindowsWithTitle('Microsoft Teams')
    
    # Print all window titles
    for item in window_list:
        if str.lower('Microsoft Teams')in item.title.lower():
            print(f"Teams Window found in {item.title}\n")
            if call_to_win_name.lower() in item.title.lower():
                item.close()
            elif inCall_win_name.lower() in item.title.lower():
                item.close()
            else:
                print('...')
                #item.activate()
                
                
            
    
    # Find the window by title
    window_list = gw.getWindowsWithTitle('Microsoft Teams')
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        
        print("Teams Main windows found\n")
        # Activate the window
        window.maximize()
    
        window.activate()
    else:
        print("No window with the title 'Microsoft Teams' found.")
    
   
    #time.sleep(3)
    #window.activate()
    pyautogui.hotkey('ctrl', 'shift','N')
    
    time.sleep(5)
    scrcpy_run(scrcpyExePath = Scrcpy_path)
    time.sleep(1)

    pyautogui.typewrite(call_to_address)
    time.sleep(1)
 
    pyautogui.press('enter')
    # time.sleep(1)
    while True:
        exists_call_to_win = window_exists_exact(call_to_win_name)
        exists_in_call_win = window_exists_exact(inCall_win_name)
        
        if (exists_call_to_win or exists_in_call_win) :
            window_list = gw.getWindowsWithTitle(call_to_win_name)
            window_call = window_list[0]
            time.sleep(1)
            #window_call.close()
            print(f"Window with title '{call_to_win_name}' or '{inCall_win_name}' exists.")
            break
        else:
            print(f"Window with title '{call_to_win_name}' and '{inCall_win_name}' does not exist.")
            pyautogui.press('enter')
            time.sleep(2)
    
    pyautogui.press('enter')
    
    time.sleep(1)
    # Send Alt+Shift+V, start video call
    if call_audio_video.lower() in "video":
        pyautogui.hotkey('alt', 'shift', 'v')
    elif call_audio_video.lower() in "audio":
        pyautogui.hotkey('alt', 'shift', 'a')
    else:
        print('Please check input if "audio" or "video')
        
    
        
         
    
    time.sleep(3)
    
    #===============================================================================start Call==============================
    scrcpy_PS_accept(scrcpyExePath = Scrcpy_path)
    
    #===============================================================================End Call==============================
    time.sleep(call_Duration)  # call duration
    #===============================================================================End Call==============================
    scrcpy_run(scrcpyExePath = Scrcpy_path)
    
    time.sleep(2)
    
    scrcpy_PS_end(scrcpyExePath = Scrcpy_path)
 
    
        
    # Print the output of the command
    #print(result.stdout)
    
    # Find the window by title
    window_list = gw.getWindowsWithTitle(call_to_win_name)
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        for item in window_list:
            print(f'try to End call on Windows Teams with "{call_to_win_name}" ...\n')
            print(item.title)
            if call_to_win_name.lower() in item.title.lower():
                item.activate()
                
                pyautogui.hotkey('ctrl', 'shift', 'H')
                pyautogui.hotkey('ctrl', 'shift', 'H')
                
                item.close()
            
    else:
        print(f"No window with the title '{call_to_win_name}' found...\n")
     
    time.sleep(1)
    
    window_list = gw.getWindowsWithTitle(inCall_win_name)
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        for item in window_list:
            print(f'try to End call on Windows Teams with "{inCall_win_name}" ...\n')
            print(item.title)
            if inCall_win_name.lower() in item.title.lower():
                item.activate()
                pyautogui.hotkey('ctrl', 'shift', 'H')
                pyautogui.hotkey('ctrl', 'shift', 'H')
                
                item.close()
    else:
        print(f"No window with the title '{inCall_win_name}' found...\n")
    # exists_call_to_win = window_exists_exact(call_to_win_name)
    # exists_in_call_win = window_exists_exact(inCall_win_name)
    
  
    # if exists_call_to_win :
    #     window_list = gw.getWindowsWithTitle(call_to_win_name)
    #     # Select the first window from the list
    #     window_call = window_list[0]
    #     window_call.activate()
    #     pyautogui.hotkey('ctrl', 'shift', 'H')
    #     window_call.close()
    # elif exists_in_call_win :
    #     window_list = gw.getWindowsWithTitle(inCall_win_name)
    #     # Select the first window from the list
    #     window_call = window_list[0]
    #     window_call.activate()
    #     pyautogui.hotkey('ctrl', 'shift', 'H')
    #     window_call.close()
        
    # else:
    #     print('Check inCall window title ...\n')
    #     time.sleep(1);
        
   
     
    end_time = time.time()
    
    duration = int(end_time - start_time)
    
    print(f"Call and exectuion time around {duration} seconds...\n")
    

   
    return duration

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

def activate_edge_and_search_keyword(winTitlewith='LogVisualizer',keyword="AAudio",keyword_2="OPENSLES"):
    # Get a list of all window titles
   
    
    while True:
        windows = gw.getAllTitles()

        # Filter windows that contain 'Microsoft Edge' in their title
        edge_windows = [title for title in windows if winTitlewith in title]
        

        if edge_windows:
            # Get the window object of the first matching window
            window = gw.getWindowsWithTitle(edge_windows[0])[0]
            print(f"Window ID: {window._hWnd}, Title: {window.title}")
        
            # Activate the window
            window.show()
            #window.maximize()
            window.activate()
        
            # Wait for the window to be activated
            time.sleep(2)
        
            # Send Ctrl+A to select all content
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(1)
        
            # Send Ctrl+C to copy the selected content to the clipboard
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(1)
        
            # Get the content from the clipboard
            clipboard_content = pyperclip.paste()
        
            # Count the occurrences of the keyword in the clipboard content
            keyword_count = clipboard_content.lower().count(keyword.lower())
            keyword_count_2 = clipboard_content.lower().count(keyword_2.lower())

            # Print the count of the keyword occurrences
            print(f"The keyword '{keyword}' appears {keyword_count} times in the page content.\n")
            print(f"The keyword '{keyword_2}' appears {keyword_count_2} times in the page content.\n")
            
            #window.close()
            break
        
        
        
    return window, keyword_count,keyword_count_2

def MPAT_Android_Test(ipaddr='10.172.206.65:5555', mpat_path=r'C:\Users\jianzou\OneDrive - Microsoft\MPAT_1.2.1',loop=1):
    
    result_file = os.path.join(mpat_path,"result_url.txt")
    
    delete_subfolders_with_prefix(mpat_path,"result")
    # Empty the existing folder by deleting all its contents
   
# print all the test config and current time to txt
    # Get the current date and time
    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if os.path.exists(result_file):
        # Append the clipboard content to the file
        with open(result_file, "a") as file:
            file.write(f"------------------------------------------{PATA_version}----------------------------------------------------\n")
            file.write(f"\nTest started on {current_datetime}\n\n")
            
            file.write(f"Call to address: {call_to_address}\n")
            
            file.write(f"MPAT Path: {MPAT_path}\n")
            
            file.write(f"Scrcpy exe Path: {Scrcpy_path}\n")
            
            file.write(f"Audio or video call: {call_audio_video}\n")
            
            file.write(f"Call to Windows name keyword: {call_to_win_name}\n")
            
            file.write(f"In call Windows name keyword: {inCall_win_name}\n")
            
            file.write(f"DUT scrcpy Windows name: {scrcpy_win_name}\n")
             
            file.write(f"Call duration: {call_Duration}\n")
            
            file.write(f"DUT ip address: {dut_ipaddr}\n")
            
            file.write(f"Test loops: {loopN}\n\n")

            #file.write(f"AAduio appears time: {times_AAudio}\n")
            #file.write(f"OpenSLES appears time: {times_OpenSLES}\n")
            #file.write(lines[-3] + "\n")
           
    else:
        # Create a new file and write the clipboard content
        with open(result_file, "w") as file:
            file.write(f"\nTest started on {current_datetime}\n")
            
            file.write(f"Call to address: {call_to_address}\n")
            
            file.write(f"MPAT Path: {MPAT_path}\n")
            
            file.write(f"Scrcpy exe Path: {Scrcpy_path}\n")
            
            
            file.write(f"Audio or video call: {call_audio_video}\n")
            
            file.write(f"Call to Windows name keyword: {call_to_win_name}\n")
            
            file.write(f"In call Windows name keyword: {inCall_win_name}\n")
            
            file.write(f"DUT scrcpy Windows name: {scrcpy_win_name}\n")
             
            file.write(f"Call duration: {call_Duration}\n")
            
            file.write(f"DUT ip address: {dut_ipaddr}\n")
            
            file.write(f"Test loops: {loopN}\n\n")

            #file.write(f"AAduio appears time: {times_AAudio}\n")
   

    
    for ii in range(1,loopN + 1):
        loopStart_time = time.time()
        
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
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.2)
                # Get the content from the clipboard
                clipboard_content = pyperclip.paste()
                if "Testing is about to begin" not in clipboard_content:
                    time.sleep(1)
                else:
                    break
            pyautogui.typewrite('y')
            
            auto_WinDesk_call_MTRA()
            
            
            #window_id.maximize()
            window_id = find_window_containing_string('C:\WINDOWS\system32\cmd.exe')
            window_id.activate()
            time.sleep(2)
            pyautogui.press('enter')
            
            while True:
                # Send Ctrl +A copy all  the contents to clipboard
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.2)
                # Get the content from the clipboard
                clipboard_content = pyperclip.paste()
                if "The test results will be uploaded" not in clipboard_content:
                    time.sleep(1)
                else:
                    time.sleep(1)
                    break#time.sleep(10)

                
            pyautogui.typewrite('y')
            
            while True:
                # Send Ctrl +A copy all  the contents to clipboard
                window_id.activate()
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.2)
                # Get the content from the clipboard
                clipboard_content = pyperclip.paste()
                if "https://" not in clipboard_content:
                    time.sleep(3)
                else:
                    time.sleep(0.2)
                    break#time.sleep(10)
                
                    
            # while window_id.isActive:
            #     time.sleep(3)
            #     print('waiting...\n')
            
            #window_id.maximize()
            window_id.activate()
            
            time.sleep(1)

            
            # Send Ctrl +A copy all  the contents to clipboard
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
            # Get the content from the clipboard
            clipboard_content = pyperclip.paste()

            lines =  clipboard_content.splitlines(); 

            #active Microsoft Edge Window and check AAudio and OpenSLES appears times
            #WinEdge,times_AAudio, times_OpenSLES = activate_edge_and_search_keyword(winTitlewith='LogVisualizer',keyword="AAudio",keyword_2="OPENSLES")
            time.sleep(3)  # wait enough seconds to get report full load to microsoft edge 
            #Edge_win_name='LogVisualizer';
            keyword="AAudio";
            keyword_2="OPENSLES"
  
            #windows = gw.getAllTitles()

            # Filter windows that contain 'Microsoft Edge' in their title
            #edge_windows = [title for title in windows if winTitlewith in title]
            
            window_list =gw.getWindowsWithTitle('LogVisualizer')
            # Check if any windows were found
            #if window_list:
            # Select the first window from the list
            window_IE = window_list[0]
            print(f"Window ID: {window_IE._hWnd}, Title: {window_IE.title}\n")
            
            
            print("MPAT result windows found\n")
            # Activate the window
            window_IE.show()
            #window.maximize()
            window_IE.activate()
        
            # Wait for the window to be activated
            time.sleep(1)
        
            # Send Ctrl+A to select all content
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.2)
        
            # Send Ctrl+C to copy the selected content to the clipboard
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
        
            # Get the content from the clipboard
            clipboard_content = pyperclip.paste()
        
            # Count the occurrences of the keyword in the clipboard content
            times_AAudio = clipboard_content.lower().count(keyword.lower())
            times_OpenSLES = clipboard_content.lower().count(keyword_2.lower())
      

           
            #time.sleep(1)
                #window.close()
                

            # Save the clipboard content to a text file
            # Check if the file already exists
            if os.path.exists(result_file):
                # Append the clipboard content to the file
                with open(result_file, "a") as file:
                    file.write(f"Test {ii}\n")
                    file.write(f"AAduio appears time: {times_AAudio}\n")
                    file.write(f"OpenSLES appears time: {times_OpenSLES}\n")
                    file.write(lines[-3] + "\n")
                   
            else:
                # Create a new file and write the clipboard content
                with open(result_file, "w") as file:
                    file.write(f"Test {ii}\n")
                    file.write(f"AAduio appears time: {times_AAudio}\n")
                    file.write(f"OpenSLES appears time: {times_OpenSLES}\n")
                    file.write(lines[-3] + "\n")
                    
            time.sleep(1)
            
            window_id.close()
        
            # new_folder = os.path.join(mpat_path,'result_'+str(ii))

            # copy_folder_contents(existing_folder,new_folder)

                 
        else:
            print("Command Prompt window not found.")
        loop_time = time.time() - loopStart_time
        print(f"loop #{ii} test time {loop_time:.1f} seconds\n")
    
    # print end time
    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if os.path.exists(result_file):
        # Append the clipboard content to the file
        with open(result_file, "a") as file:
            file.write(f"\nTest finished on {current_datetime}\n")
            

    else:
        # Create a new file and write the clipboard content
        with open(result_file, "w") as file:
            file.write(f"\nTest finished on {current_datetime}\n")
            
 
    return result_file

def open_notepad_with_file(file_name):
    # Check if the file exists
    if os.path.exists(file_name):
        # Open Notepad with the specified file
        subprocess.run(['notepad.exe', file_name])
    else:
        print(f"The file '{file_name}' does not exist.")
        
    return


if __name__ == "__main__":
    # print("test")

    MPAT_path = select_MPAT_folder()
    Scrcpy_path = select_Scrcpy_folder()
    
    result_file = MPAT_Android_Test(ipaddr=dut_ipaddr,mpat_path= MPAT_path, loop=loopN)
###### Display the result ==================================================================== 
    #result_file = r"C:\Users\jianzou\OneDrive - Microsoft\MPAT_1.2.1\result_url.txt"
    open_notepad_with_file(result_file) 