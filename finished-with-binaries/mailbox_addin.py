import os
import win32com.client
import tkinter as tk
from tkinter import filedialog
from tkinter import Listbox
import pyperclip
import tempfile
import base64
from multiprocessing.shared_memory import SharedMemory


attachment_dict = {}


def get_inbox_emails():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder
    messages = inbox.Items
    return messages

def get_attachments_from_emails(emails):
    attachments = []
    for message in emails:
        if message.Attachments.Count > 0:
            for attachment in message.Attachments:
                attachments.append(attachment)
    return attachments

def save_attachment_to_folder(attachment):
    file_name = filedialog.asksaveasfilename(defaultextension='.*',
                                             filetypes=[('All Files', '*.*')],
                                             initialfile=attachment.FileName)
    if file_name:
        attachment.SaveAsFile(file_name)
        print(f"Attachment {attachment.FileName} saved to {file_name}")

def save_attachment_to_byte_array(attachment):
    byte_array = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102")
    print(f"Attachment {attachment.FileName} saved to byte array")
    return byte_array


def on_save_as_button_click():
    attachment_file_name = attachments_listbox.get(tk.ACTIVE)
    attachment = attachment_dict.get(attachment_file_name)
    if attachment:
        save_attachment_to_folder(attachment)

def on_save_to_byte_array_button_click():
    global global_attachment_byte_array
    attachment_file_name = attachments_listbox.get(tk.ACTIVE)
    attachment = attachment_dict.get(attachment_file_name)
    if attachment:
        global_attachment_byte_array = save_attachment_to_byte_array(attachment)
        print(f"Attachment {attachment_file_name} saved to global_attachment_byte_array")



# Tkinter window
root = tk.Tk()
root.title("Outlook Attachments")
listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=100, height=35)
#listbox to display attachments
attachments_listbox = tk.Listbox(root, width=50)
attachments_listbox.pack(pady=10)

# Populate the listbox with attachment names
emails = get_inbox_emails()
attachments = get_attachments_from_emails(emails)
for attachment in attachments:
    attachment_dict[attachment.FileName] = attachment
    attachments_listbox.insert(tk.END, attachment.FileName)
# 'Save As...' button
save_as_button = tk.Button(root, text="Save As...", command=on_save_as_button_click)
save_as_button.pack(pady=5)

# 'Save to Byte Array' button
save_to_byte_array_button = tk.Button(root, text="Save to Byte Array", command=on_save_to_byte_array_button_click)
save_to_byte_array_button.pack(pady=5)

# Start the Tkinter main loop
root.mainloop()
