'''
Author: Jake Alexander
Purpose: Make links to dropbox folders in a DropBox directory and add them
to a google spreadsheet.
'''

import pandas as pd
import dropbox as dropbox
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

#Read the file into a PD DF
email_list = pd.DataFrame(sheet.get_all_records())

ad_rep_col = email_list['Ad Rep']

'''
#This is a way to separate my stores from other stores
my_stores = []

for idx in range(len(ad_rep_col)):
    if ad_rep_col[idx] == 'JA' and email_list['STORE NAME'][idx] != '':
        my_stores.append((email_list['STORE NAME'][idx]).lower())
'''

all_dropbox_links = email_list['Main']
stores = email_list['STORE NAME']
all_dropbox_links = email_list['Monthly']
'''UPDATE THIS'''
monthly_folder_name = '5 - May 2021 Ad Content'

access_token = ''
dbx = dropbox.Dropbox(access_token)
folders = []


for folder in dbx.files_list_folder('/1 BB/Active BB Clients').entries:
    folders.append(folder.path_lower)
    ''' Here is a way to create new subfolders
    new_folder = folder.path_lower + "/13 - 2020 folders"
    dbx.files_create_folder(new_folder, True)
    print(new_folder)
    '''
    
for sub_folder in folders:
    for monthly_folder in dbx.files_list_folder(sub_folder).entries:
        #Important!!! Update this for each month
        if monthly_folder.name == monthly_folder_name:
            for i in range(len(stores)):
                if stores[i] != '' and stores[i].lower() in sub_folder and all_dropbox_links[i] == '':
                    try:
                        shared_link_metadata = dbx.sharing_create_shared_link_with_settings(monthly_folder.path_lower)
                    except:
                        shared_link_metadata = dbx.sharing_get_shared_links(monthly_folder.path_lower).links[0]
                    all_dropbox_links[i] = shared_link_metadata.url
                    
                    #This will find and write in all of the links that already exist.
                    meta_data = dbx.sharing_list_shared_links(monthly_folder.path_lower).links
                    for entry in meta_data:
                        if entry.name == monthly_folder_name:
                            all_dropbox_links[i] = entry.url

current_db_link_col = sheet.find('Monthly')

#write entries in spreadsheet. ***THIS MIGHT CAUSE A PROBLEM***
for i in range(stores):
    sheet.update_cell(i + 2, current_db_link_col, all_dropbox_links[i + 2])
