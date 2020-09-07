'''
Author: Jake Alexander
Purpose: This program accesses information from a spreadsheet to send emails 
formatted with links. It uses Email Sender.py to send the emails.
Date: 9/3/20
'''

import pandas as pd
import datetime
from Email Sender.py import send_email

excel_file = "file name/path here"

# Read the file - make sure this is the correct file!! Put it in the folder.
email_list = pd.read_excel(excel_file, keep_default_na=False)

# Get all the  Variables. Make sure these cell names match!
store_names = email_list['STORE NAME']
all_ad_links = email_list['Ad Link']
ad_rep = email_list['Ad Rep']
link_sent = email_list['Link']
current_time = datetime.datetime.now()
current_date = str(current_time.month) + '/' + str(current_time.day)
subject = 'Monthly Bridal Boutiques Ad Link'

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
        
        if rep != 'JA' or ad_link == '' or store_names[idx] == '' or not link_sent[idx] == '':
            print(store_names[idx])
            continue
        
        # Create the email body to send
        if len(ad_link) == 1:
            body = ('Formatted message goes here using {0},{1},{2} as variable placeholders'.format(ad_link[0]))
        if len(ad_link) == 2:
            body = ('Read Above'.format(ad_link[0], ad_link[1]))
        if len(ad_link) == 3:
            body = ('Read Above'.format(ad_link[0], ad_link[1], ad_link[2]))
        if len(ad_link) == 4:
            body = ('Read Above'.format(ad_link[0], ad_link[1], ad_link[2], ad_link[3]))
        
        #Use the send email function from Email Sender.py
        send_email(idx, subject, body)
        
        #mark the date that the ad was sent to the client
        link_sent[idx] = current_date

#Write the file - make sure this is the correct file path
with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
    email_list.to_excel(writer, sheet_name='Sheet1')

if __name__ == '__main__':
    main()
