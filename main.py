# Pseudocode
# Download file from email and place in directory with project (Look at interacting directly with outlook)
# Open xlsx file and delete first three rows of data and check if last row has data not needed
# Save edited file as both csv and xlsx and add the date in the format YYYY-MM-DD to the end of the name before the extension
# Verify that all Unique Identifiers have 2341 as the last 4 digits of column 1 in the CSV file
# Move new files to //LB/SharedSpace/Systems/Discovery/Prospector/Patron Records/
# Delete the active file from the directory

import datetime
import os
import win32com.client
path = os.path.expanduser("C:\\Users\\garrettthompson_a\\Downloads")
today = datetime.date.today()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)


def save_attachments(subject, messages):
    message = messages.GetFirst()
    print(str(message))
    if message.Subject == subject:
        # body_content = message.body
        attachments = message.Attachments
        attachment = attachments.Item(1)
        attachment.SaveAsFile(os.path.join(path, str(attachment)))
        if message.Subject == subject and message.Unread:
            message.Unread = False


