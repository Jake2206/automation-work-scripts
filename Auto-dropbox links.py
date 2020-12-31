'''
Author: Jake Alexander
Purpose: Send links using email sending program to clients.
Date: 9/4/20
'''

import pandas as pd
from email sender import send_email

'''Make sure to update this file!!!'''
excel_file = "file name goes here"

# Read the file - make sure this is the correct file and such!! Put it in the folder.
email_list = pd.read_excel(excel_file, keep_default_na=False)

# Get all the  Variables. Make sure these cell names match!
store_names = email_list['STORE NAME']
all_dropbox_links = email_list['Monthly']
already_uploaded = email_list['Uploaded']

subject = 'Subject goes here'
intro = 'I hope business is going well!\n I have included a link to your current monthly DropBox folder below for your convenience.\n Remember that creativity and personal touch are great ways to get Bride\'s attention!'
def send_emails(store_idx):
    text = ( '{0}\n{1}'.format(intro, all_dropbox_links[store_idx]))

    send_email(store_idx, subject, text)
    return True

def main():
    not_sent = []
    for i in range(len(all_dropbox_links)):
        dropbox_link = all_dropbox_links[i]
        if dropbox_link[0:5] == 'https' and already_uploaded[i] == '':
            send_emails(i)
        else:
            not_sent.append(store_names[i])
    print(not_sent)

if __name__ == '__main__':
    main()