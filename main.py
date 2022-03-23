import win32com.client
import os
from datetime import datetime, timedelta

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def extract_attachments(mailbox, restrictMessage, outputDir):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Mailbox: {mailbox}')  # Press Ctrl+F8 to toggle the breakpoint.
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")

    for root in mapi.Folders:
        try:
          if (mailbox in root.FolderPath):
             print ("FolderPath: " + root.FolderPath)
             for folder in root.Folders:
                 print("folder: " + folder.FolderPath)
                 messages = folder.Items
                 restrictedMessages = messages.Restrict(restrictMessage)
                 for message in list(restrictedMessages):
                    received = str(message.ReceivedTime)
                    print (message.SenderEmailAddress)
                    print (received[0:10])
                    receivedString = received[0:10] + "_"
                    for attachment in message.Attachments:
                       attachment.SaveASFile(os.path.join(outputDir, receivedString + attachment.FileName))
                       print(f"attachment {receivedString +  attachment.FileName}  saved")

        except:
            print("exception")

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    outputDir = r"f:\swissedu_attachments"
    extract_attachments('ed@leijnse.info', "[SenderEmailAddress] = 'helena.dimi@windowslive.com'", outputDir)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
