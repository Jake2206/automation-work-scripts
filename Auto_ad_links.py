'''
Author: Jake Alexander
Purpose: This program accesses information from a spreadsheet to send emails 
formatted with links. It uses Email_Sender.py to send the emails.
'''

import pandas as pd
import datetime
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

#Read the file into a PD DF
email_list = pd.DataFrame(sheet.get_all_records())

# Get all the  Variables. Make sure these cell names match!
store_names = email_list['STORE NAME']
all_ad_links = email_list['Ad Link']
ad_rep = email_list['Ad Rep']
link_sent = email_list['Link']
sent_link_col = sheet.find('Link').col

#set the month
current_time = datetime.datetime.now()
current_date = str(current_time.month) + '/' + str(current_time.day)
subject = 'May Ad Link - Bridal Boutiques' #UPDATE ME!!

#get all links for a specific store
def get_links(possible_links, idx):
    links = []
    if possible_links[idx] == 'N/A' or possible_links[idx] == '':
        return ''
    links.append(possible_links[idx])
    
    if (idx == len(ad_rep)-10):
        return links
    
    for i in range(1, 3):
        if store_names[idx+i] != '':
            break
        else:
            links.append(possible_links[idx+i])
    
    return links

def main():
    # Loop through the links column
    for idx in range(len(ad_rep)):

        ad_link = get_links(all_ad_links, idx)
        rep = ad_rep[idx]
        
        if rep != 'JA' or ad_link == '' or store_names[idx] == '' or link_sent[idx] != '':
            if store_names[idx] != '':
                print('This store did not recieve an ad: ' + store_names[idx])
            continue
        
        # Create the email body to send
        if len(ad_link) == 1:
            body = ('I hope you are doing well! Here is your ad link for the upcoming month: \n{0}\n\nIt is scheduled to run on the 1st of the month. Let me know if you have questions, concerns, or edits.'.format(ad_link[0]))
        if len(ad_link) == 2:
            body = ('I hope you are doing well! Here are your ad links for the upcoming month:\n{0}\n{1}\n\nThey are scheduled to run on the 1st of the month. Let me know if you have questions, concerns, or edits.'.format(ad_link[0], ad_link[1]))
        if len(ad_link) == 3:
            body = ('I hope you are doing well! Here are your ad links for the upcoming month:\n{0}\n{1}\n{2}\n\nThey are scheduled to run on the 1st of the month. Let me know if you have questions, concerns, or edits.'.format(ad_link[0], ad_link[1], ad_link[2]))
        if len(ad_link) == 4:
            body = ('I hope you are doing well! Here are your ad links for the upcoming month:\n{0}\n{1}\n{2}\n{3}\n\nThey are scheduled to run on the 1st of the month. Let me know if you have questions, concerns, or edits.'.format(ad_link[0], ad_link[1], ad_link[2], ad_link[3]))
        
        #Use the send email function from send_emails.py
        send_email(idx, subject, body)
        
        #mark the date that the ad was sent to the client
        #note the offset by 2 for the 2 header lines (not recognized in the DF but are recognized in gspread)
        sheet.update_cell(idx + 2, sent_link_col, current_date) 

if __name__ == '__main__':
    main()
