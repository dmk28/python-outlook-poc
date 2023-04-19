
import time 
import psutil
from pymem import Pymem, pymem
import ctypes
import subprocess
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
import os
import tempfile
import re
from pymem.ptypes import RemotePointer
import tkinter as tk
from tkinter import filedialog


global hasRun
hasRun = False
global process_address
process_address = b'0x0001'
file_exts = {
    'docx': ['Microsoft Word 2007+', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.1'],
    'pptx': ['Microsoft PowerPoint 2007+', 'PowerPoint 97-2003'],
    'xlsx': ['Microsoft Office 2007+', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.0', 'Composite Document File V2 Document, Little Endian, Os: Windows, Version 6.1'],
    # add more file types and their corresponding magic numbers here
}




def check_for_outlook():
    print("Check for outlook initialized")
    global hasRun
    outlook_found = False
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.name().lower() == 'outlook.exe':
                pid = proc.pid
                print("Read process")
                file_type = read_process(pid)
                if file_type:
                    root = tk.Tk()
                    root.withdraw()
                    default_ext = '.' + file_type
                    file_path = filedialog.asksaveasfilename(defaultextension=default_ext)
                outlook_found = True
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            print("Not reading process")
            pass
    if not outlook_found:
        print("Outlook process not found.")
 


def read_process(pid):
    global process_address
    print("Let's get a handle...")
    try:
        process = pymem.Pymem("outlook.exe")
        process.open_process_from_id(pid)
        print("Got handle... Getting base address")
        base_address = process.base_address
        print("Base address:", base_address)
        address_size = process.__sizeof__
        print("How big is this?", address_size)
        process_mods = list(process.list_modules())
        print("Module listing successful, printing to list...")
        control  = True
        while control == True:
            print("Searching pattern")
            pattern = [b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1', b'\x50\x4B\x03\x04', b'\x50\x4b\x03\x04\x14\x00\x06\x00',b'\x50\x4b\x03\x04\x14\x00\x06\x00',
            b'\x50\x4b\x03\x04\x14\x00\x06\x00',
            b'\x50\x4b\x03\x04\x14\x00\x06\x00']
            handle = process.process_handle
            search = pymem.memory.virtual_query(handle, base_address)
            time.sleep(0.1)
            for p in pattern:

                search_result = pymem.pattern.pattern_scan_all(handle, p)
                if not search_result:
                    print("Couldn't find attachment in process memory")
                    return
                else:
                    scan_bytes = pymem.memory.read_bytes(address=search_result, handle=handle, byte=len(p))
                    print("Pattern found, let's proceed")
                    file_extraction(handle, base_address)
                    

                
                control = False
                

           
    except pymem.exception.MemoryReadError:
        print("Exception found: MemoryReadError")
        
      
    except MemoryError as e:
        print("Couldn't access memory", e)

    return None


def file_extraction(handle, process_address):
    try:
        print("Extracting...")
        mem_info = pymem.memory.virtual_query(handle, address=process_address)
        bytes = pymem.memory.read_bytes(address=mem_info.BaseAddress, byte=mem_info.RegionSize, handle=handle)
        with tempfile.NamedTemporaryFile(delete=False) as f:
            f.write(bytes)
            f.flush()
            os.fsync(f.fileno())
            with open(f.name, 'rb') as f2:
                file_data = f2.read()
        save_file_dialog(file_data=file_data)
    except pymem.exception.MemoryReadError:
        print("Exception found: MemoryReadError")

 

def save_file_dialog(file_data):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(defaultextension='.docx')
    if file_path:
        with open(file_path, 'wb') as f:
            f.write(file_data)
    root.quit()
def test_run():
    outlook = win32.Dispatch("Outlook.Application")
    inbox = outlook.GetNameSpace("MAPI").GetDefaultFolder(6)
    email = inbox.Items[0]
    attachment = email.Attachments.Item(1)
    attachment.SaveAsFile(r"C:\savepath")

if __name__ == '__main__':
   check_for_outlook()

