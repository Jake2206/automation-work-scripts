'''
Author: Jake Alexander
Purpose: Checking DropBox folders for uploads.
'''

import dropbox as dropbox
from datetime import date
from dateutil.relativedelta import relativedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Authorize the API
scope = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file'
    ]
file_name = 'client_key.json'
creds = ServiceAccountCredentials.from_json_keyfile_name(file_name,scope)
client = gspread.authorize(creds)

#fetch the sheet
sheet = client.open('2021 - Monthly check lists').sheet1

#Get the next month's date and format the string for searching the sub folders
date = date.today() + relativedelta(months=+1)
month_num = date.month
month_str = date.strftime("%B")
year = date.strftime("%Y")
date = (str(date.month) + " - " + date.strftime("%B") + " " + str(date.strftime("%Y")))

#access dropbox
access_token = ''
dbx = dropbox.Dropbox(access_token)

#set up variables
non_empty_folders = []

#Look through the file system starting at the active clients directory
for folder in dbx.files_list_folder('/1 BB/Active BB Clients').entries:
    #look through each folder's sub folders 
    for monthly_folder in dbx.files_list_folder(folder.path_lower).entries:
        #check for the current month
        if monthly_folder.name == date + ' Ad Content':
            #check if the current month's subfolder has any entries
            if dbx.files_list_folder(monthly_folder.path_lower).entries:
                ind = folder.name.find(' - ')
                name = folder.name[:ind]
                print(name)
                non_empty_folders.append(name)

#find the correct column index
uploaded_col_num = sheet.find('Uploaded').col

#mark each cell that has uploaded files in the folder
for store in non_empty_folders:
    cell = sheet.find(store)
    sheet.update_cell(cell.row, uploaded_col_num, 'X')
        

