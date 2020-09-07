'''
Author: Jake Alexander
Purpose: This is a program that takes in a spreadsheet and sends emails to clients
'''

import pandas as pd
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

'''
Change these to your credentials and name
'''
your_name = "Jake"
your_email = "email account"
your_password = "enter password"
excel_file = "file name and path"

# IMPORTANT! set this to true if you want to send the emails
sendNow = False
    
# Read the file - make sure this is the correct file and such!! Put it in the folder.
email_list = pd.read_excel(excel_file, keep_default_na=False)

# Get all the  Variables. Make sure these cell names match!
all_names = email_list['Name']
all_emails = email_list['Social Media Contact\'s']

#Formats the beginning of the email with all first names of client.
def compose(firstnames):
    numNames = len(firstnames)
    if numNames == 0:
        return("")
    if numNames == 1:
        return firstnames[0]
    if numNames == 2:
        return( "{0} and {1}".format(firstnames[0], firstnames[1]))
    if numNames == 3:
        return( "{0}, {1}, and {2}".format(firstnames[0], firstnames[1], firstnames[2]))
    if numNames == 4:
        return( "{0}, {1}, {2}, and {3}".format(firstnames[0], firstnames[1], firstnames[2], firstnames[3]))
    else:
        return("all")

def greetingGet(names):
    
    if pd.isnull(names):
        return "No name / blank cell"
    fullnames = [x.strip() for x in names.split(',')]
    firstnames = []
    for name in fullnames:
        fname = name.split()[0]
        firstnames.append(fname)
    hi_firstnames = compose(firstnames)
    return hi_firstnames

#Sends the email. Takes an index, subject, and body and creates an email.
def send_email(idx, subject, body):
    attachment = None
    
    if len(body) == 2:
        attachment = body[1]
        body = body[0]
        
    # Get/format name and email
    name = all_names[idx]
    
    # You will need to allow less secure apps to access your gmail account.
    # There's a setting you change in the gmail dashboard. It's annoying
    if sendNow == True:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(your_email, your_password)
    
    if pd.isnull(name) or name == '':
        return "Invalid store index"
    
    firstname = greetingGet(name)
    email = all_emails[idx]
    
    # Create the email to send
    message = MIMEMultipart()
    message["From"] = your_email
    message["To"] = email
    message["Subject"] = subject
    
    full_email = ("Hello {0},\n\n" 
                  "{1}\n\n"
                  "Best,\n"
                  "-Jake"
                  .format(firstname, body))
    
    message.attach(MIMEText(full_email, "plain"))
    
    if attachment != None:
       filename = attachment
    # Open PDF file in binary mode
    
    # We assume that the file is in the directory where you run your Python script from
    with open(filename, "rb") as attachment:
        # The content type "application/octet-stream" means that a MIME attachment is a binary file
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    
        encoders.encode_base64(part)
        
        part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {filename}",
                    )
    # Add attachment to your message and convert it to string
    message.attach(part)
        
    text = message.as_string()
    
    
    # This prints all your emails so you can see if they look ok.
    print(full_email)
    print("\n\n")

    # This sends email:
    if sendNow == True:
        try:
            server.sendmail(your_email, email.split(','), text)
            
            print('Email to {} successfully sent!\n\n'.format(email))
        except Exception as e:
            print('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))

    # Close the smtp server
    if sendNow == True:
        server.close()

