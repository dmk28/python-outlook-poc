import time
import psutil
import pymem
from pymem import Pymem, pymem, memory
import tkinter as tk
from tkinter import filedialog
import os
import tempfile

file_exts = {
    'docx': [b'\x50\x4B\x03\x04'],
    'pptx': [b'\x50\x4B\x03\x04'],
    'xlsx': [b'\x50\x4B\x03\x04'],
    # add more file types and their corresponding magic numbers here
}

def check_for_outlook():
    print("Check for outlook initialized")
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
        control = True
        while control:
            print("Searching pattern")
            pattern_found = False
            handle = process.process_handle
            search = pymem.memory.virtual_query(handle, base_address)
            time.sleep(0.1)
            for file_type, patterns in file_exts.items():
                for p in patterns:
                    search_result = pymem.pattern.pattern_scan_all(handle, p)
                    if search_result:
                        scan_bytes = pymem.memory.read_bytes(address=search_result, handle=handle, byte=len(p))
                        print("Pattern found, let's proceed")
                        extracted_file_data = file_extraction(handle, base_address)
                        if extracted_file_data:
                            save_file_dialog(file_data=extracted_file_data, file_type=file_type)
                            pattern_found = True
                            break
                if pattern_found:
                    break
            control = False

    except pymem.exception.MemoryReadError:
        print("Exception found: MemoryReadError")

    except pymem.exception.MemoryWriteError as e:
        print("Couldn't access memory", e)

    return None


def file_extraction(handle, process_address):
    try:
        print("Extracting...")
        mem_info = pymem.memory.virtual_query(handle, address=process_address)
        bytes = pymem.memory.read_bytes(address=mem_info.BaseAddress, byte=mem_info.RegionSize, handle=handle)
        return bytes
    except pymem.exception.MemoryReadError:
        print("Exception found: MemoryReadError")
        return None


def save_file_dialog(file_data, file_type):
    root = tk.Tk()
    root.withdraw()
    default_ext = '.' + file_type
    file_path = filedialog.asksaveasfilename(defaultextension=default_ext)
    if file_path:
        with open(file_path, 'wb') as f:
            f.write(file_data)
    root.quit()
    root.destroy()


if __name__ == '__main__':
    check_for_outlook()
