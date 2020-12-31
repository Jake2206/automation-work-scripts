'''
Author: Jake Alexander
Purpose: Make links to dropbox folders in a DropBox directory and add them
to an excel spreadsheet.
Date: 9/6/20
'''

import pandas as pd
import dropbox as dropbox

'''IMPORTANT!! Put correct file name in here.'''
excel_file = "file name goes here"
email_list = pd.read_excel(excel_file, keep_default_na=False)
ad_rep_col = email_list['Ad Rep']
my_stores = []

for idx in range(len(ad_rep_col)):
    if ad_rep_col[idx] == 'JA' and email_list['STORE NAME'][idx] != '':
        my_stores.append((email_list['STORE NAME'][idx]).lower())

all_dropbox_links = email_list['Main']
stores = email_list['STORE NAME']
all_dropbox_links = email_list['Monthly']
    
access_token = 'DropBox access token goes here'
dbx = dropbox.Dropbox(access_token)
folders = []


for folder in dbx.files_list_folder('path goes here').entries:
    folders.append(folder.path_lower)
    
    #dbx.files_create_folder(new_folder, True)
    #print(new_folder)
    
    

for sub_folder in folders:
    for monthly_folder in dbx.files_list_folder(sub_folder).entries:
        #Important!!! Update this for each month
        if monthly_folder.name == 'folder name goes here':
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
                        if entry.name == 'entry name goes here':
                            all_dropbox_links[i] = entry.url
                            print(all_dropbox_links[i])
                            print(stores[i])

#Write the file - make sure this is the correct file path
with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
    email_list.to_excel(writer, sheet_name='Sheet1')