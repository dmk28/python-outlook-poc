import psutil
import pymem
import pymem.exception
import os
import base64
import tkinter as tk
from tkinter import filedialog

def find_process_by_name(process_name):
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.name().lower() == process_name.lower():
                return proc
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return None

def read_process(pid):
    try:
        process = pymem.Pymem()
        process.open_process_from_id(pid)

        for module in process.list_modules():
            base_address = module.lpBaseOfDll
            module_size = module.SizeOfImage

            memory_region = process.read_bytes(base_address, module_size)
            search_pattern = {
                'docx': [b'\x50\x4B\x03\x04'],
                'pptx': [b'\x50\x4B\x03\x04'],
                'xlsx': [b'\x50\x4B\x03\x04'],
                'png': [b'\x89\x50\x4E\x47\x0D\x0A\x1A\x0A'],
                'pdf': [b'\x25\x50\x44\x46'],
                'jpg': [b'\xFF\xD8\xFF'],
                # add more file types and their corresponding magic numbers here
            }  # Magic number for MS Office files

            for file_type, magic_number_list in search_pattern.items():
                for magic_number in magic_number_list:
                    position = memory_region.find(magic_number)

                    if position != -1:
                        print(f"Pattern found at address: {hex(base_address + position)}")
                        file_name = filedialog.asksaveasfilename(defaultextension=f'.{file_type}', initialfile=f'attachment.{file_type}')
                        file_data = memory_region[position:]
                        attachment_byte_array = base64.b64encode(file_data)
                        print(f"Attachment {file_type} saved to attachment_byte_array")
                        return "OK"

    except pymem.exception.MemoryReadError:
        print("Exception found: MemoryReadError")

    except pymem.exception.MemoryWriteError as e:
        print("Couldn't access memory", e)

def main():
    process_name = 'mailbox_addin.exe'
    while True:
        proc = find_process_by_name(process_name)
        if proc:
            read_process(proc.pid)
            if read_process(proc.pid) == "OK":
                break
        else:
            print(f"Process '{process_name}' not found.")

if __name__ == "__main__":
    main()
