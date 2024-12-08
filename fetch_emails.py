import win32com.client
import pythoncom
from bs4 import BeautifulSoup
import re

def connect_to_outlook_app_windows():
    # Initialize COM library
    pythoncom.CoInitialize()

    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # This opens an Outlook (Classic)'s window in which to 
    # sign up with the desired address. Once done, it's not 
    # required any more.

    # Access the Inbox's subfolder called 'Newsletters'
    inbox = namespace.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder
    newsletters_folder = inbox.Folders["Newsletters"]  # Access the 'Newsletters' subfolder
    messages = newsletters_folder.Items

    pythoncom.CoUninitialize()
    return messages

def write_to_txt(messages):
    with open("all_week_news.txt", "w+", encoding="utf-8") as f:
        for idx in range(messages.Count):
            # Clean up HTML with BeautifulSoup
            message = BeautifulSoup(messages[0].Body, "html.parser")

            # Remove any addresses
            sanitized_content = re.sub(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', "[REDACTED]", message.get_text())

            # Save to file
            str_separator = "".join(("\n", "-"*50, "\n", "New document", "\n", "-"*50, "\n"))
            f.write(str_separator)
            f.write(sanitized_content)

def main():
    write_to_txt(connect_to_outlook_app_windows())