'''
Author: Jake Alexander
Purpose: Checking DropBox folders for uploads.
Date: 9/1/2020
'''

import pandas as pd
import dropbox as dropbox

#IMPORTANT!! Update this file name
email_list = pd.read_excel("excel file name/path goes here", keep_default_na=False)
ad_rep_col = email_list['Ad Rep']
my_stores = []

for idx in range(len(ad_rep_col)):
    if (ad_rep_col[idx] == 'JA' or ad_rep_col[idx] == 'LB') and email_list['STORE NAME'][idx] != '':
        my_stores.append((email_list['STORE NAME'][idx]).lower())

all_dropbox_links = email_list['Main']
    
access_token = 'access token goes here'
dbx = dropbox.Dropbox(access_token)
folders = []
monthly_folders = []
non_empty_folders = []

for folder in dbx.files_list_folder('DropBox path goes here').entries:
    folders.append(folder.path_lower)

for sub_folder in folders:
    for monthly_folder in dbx.files_list_folder(sub_folder).entries:
        #Important!!! Update this for each month
        if monthly_folder.name == 'subfolder name goes here':
            monthly_folders.append(monthly_folder.path_lower)

for folder in monthly_folders:
    if dbx.files_list_folder(folder).entries:
        non_empty_folders.append(folder)

for folder in non_empty_folders:
    for store in my_stores:
        if store in folder:
            print(store)     
            