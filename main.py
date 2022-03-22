import win32com.client
import os
from datetime import datetime, timedelta

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")
    for account in mapi.Accounts:
        print(account.DeliveryStore.DisplayName)
        if (account.DeliveryStore.DisplayName == "ed@leijnse.info"):
            inbox = mapi.GetDefaultFolder(6)
            messages = inbox.Items
            for message in list(messages):
                print(message.Subject)
                print(message.Body)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
