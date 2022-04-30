import win32com.client
import os
from win32com.client import Dispatch
import pandas as pd
from datetime import datetime
import os
import time
from datetime import datetime


path = r"C:\Users\User\Documents\Email Files"

# code to read only unread emails.
def save_file(messages,path):
    for message in messages:
        if message.Unread:
            attachments = message.Attachments
            sender_email_address =  message.SenderEmailAddress
            email_subject = message.Subject
            for attachment in attachments:
                print(f"The attachment in the email is {attachment.FileName}")
                print(f"The Email address of the sender is {sender_email_address}")                
                print(f"The Subject of the email is {email_subject}")                
                print("Please enter the attachment extension you want to save:")
                attachment_extension = str(input())
                if attachment_extension in attachment.FileName:  
                    attachment.SaveAsFile(os.path.join(path, str(attachment)))
                    print(f"attachment {attachment.FileName} saved")
                else:
                    print("We do not want to save other attachments")
            if message.Unread:
                message.Unread = False
            break

def main():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    # messages = messages.Restrict("[SenderEmailAddress] = ' '") #TODO: Filter based on sender if required
    messages = inbox.Items
    save_file(messages,path)
    print ("Sucessfully saved the file")


def extract_latest_file_and_push_to_DB():
    arr = os.listdir(path)
    # latest_date = datetime.now().strftime("%Y%m%d")
    #trying to exctarct only the latest file based on the timestamp given in the file
    latest_date = "20220329"
    for i in arr:
        name = i.split("_")[1]
        file_date = name.split(".")[0]
        print(file_date)
        if "(1)" in file_date:
            print("Do not consider the file with (1)")
        if "(1)" not in file_date and file_date > latest_date:
            df_file = pd.read_excel(f"{path}\{i}")
            print(f"{path}\{i}")
            df_file.to_excel(f"{path}\Output{file_date}.xlsx")
            
main()
# extract_latest_file_and_push_to_DB()



