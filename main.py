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
import pandas as pd
import openpyxl
path = os.path.expanduser("C:\\Users\\garrettthompson_a\\Downloads\\Prospector\\ProspectorAutomatedProject")
today = datetime.date.today()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)


# This function converts the csv data that we have cleaned into something that we can use a bit
# easier and make it look like the data that we already have
def convert_xlsx_csv():
    read_file = pd.read_excel("ProspectorPatrons.xlsx")
    read_file.to_csv("ProspectorPatrons.csv", index=None, header=True)


# This function allows us to automatically space the columns within the excel file so the user no longer has to do this
# themselves this is done through the openpyxl library
def delete_rows(sheet):
    # iterate the sheet by rows
    for row in sheet.iter_rows():

        # all() return False if all of the row value is None
        if not all(cell.value for cell in row):
            # detele the empty row
            sheet.delete_rows(row[0].row, 1)

            # recursively call the remove() with modified sheet data
            delete_rows(sheet)
            return


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
    xfile = openpyxl.load_workbook("ProspectorPatrons.xlsx")
    sheet = xfile["Sheet1"]
    delete_rows(sheet)
    xfile.save("ProspectorPatrons.xlsx")
    convert_xlsx_csv()


save_attachments("prospector", messages)