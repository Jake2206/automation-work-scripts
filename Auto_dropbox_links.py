'''
Author: Jake Alexander
Purpose: Send links using email sending program to clients.
'''

import pandas as pd
from Email_Sender import send_email
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

# Read the file into a PD DF
email_list = pd.DataFrame(sheet.get_all_records())

# Get all the  Variables. Make sure these cell names match!
store_names = email_list['STORE NAME']
all_dropbox_links = email_list['Monthly']
already_uploaded = email_list['Uploaded']

subject = 'Next Dropbox folder for your BB service'
intro = 'I hope that you are doing well! Below is a link to your next Dropbox folder; please let me know if you have any trouble with it. We need 5-10 pics or one short video per ad, and remember that creativity and unique content will really help grab Brides\' attention!.\n'
def send_emails(store_idx):
    text = ( '{0}\n{1}'.format(intro, all_dropbox_links[store_idx]))

    send_email(store_idx, subject, text)
    return True

def main():
    not_sent = []
    for i in range(len(all_dropbox_links)):
        dropbox_link = all_dropbox_links[i]
        if dropbox_link[0:5] == 'https' and (already_uploaded[i] == '' or already_uploaded[i] == None):
            send_emails(i)
        elif store_names[i] != '':
            not_sent.append(store_names[i])
    for entry in not_sent:
        print(entry)

if __name__ == '__main__':
    main()
