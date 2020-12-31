'''
Author: Jake Alexander
Purpose: Make a pdf monhtly service report out of information on two 
spreadhseets. 
Date: 9/7/20
'''

from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd
from send_emails import send_email
import datetime
import calendar

#Get the month
current_time = datetime.datetime.now()
current_month = calendar.month_name[current_time.month-1] + ' 2020 Ad Results'

#IMPORTANT!! Update these with the correct files

#Also make sure to delete stores that quit and add new stores from the old file!
old_monthly_excel = 'file path goes here'
monthly_excel = 'file path goes here'
ad_stats_excel = 'file path goes here'

old_monthly = pd.read_excel(old_monthly_excel, keep_default_na=False)
monthly = pd.read_excel(monthly_excel, keep_default_na=False)
ad_stats = pd.read_excel(ad_stats_excel, keep_default_na=False)

#Get all of the individual info for a store using the index on the current monthly sheet
def get_info(idx):
    report_idx = None
    store = monthly['STORE NAME'][idx]
    
    #get the index in the ad results spreadsheet of the ad for matching store
    for i in range(len(ad_stats['Ad Name'])):
        if store in ad_stats['Ad Name'][i]:
            report_idx = i
    
    #return null if the ad is missing
    if report_idx == None:
        return None
    
    #Get the owner's name, exclude any other names
    owner = monthly['Name'][idx]
    owner = [x.strip() for x in owner.split(',')]
    owner = owner[0]
    
    #Go through the spreadhseet from the month that the ad is for to get the ad link
    for j in range(len(old_monthly)):
        if old_monthly["STORE NAME"][j] == monthly['STORE NAME'][idx]:
            ad_link = old_monthly['Ad Link'][j]
    
    reach = ad_stats['Reach'][report_idx]
    frequency = round(ad_stats['Frequency'][report_idx], 2)
    impressions = ad_stats['Impressions'][report_idx]
    bb_page = monthly['Link to BB'][idx]
    
    return store, owner, ad_link, reach, frequency, impressions, bb_page

def make_report(store, owner, ad_link, reach, frequency, impressions, BB_page):
    reach = f'{reach:,}'
    impressions = f'{impressions:,}'
    
    #load fonts
    pdfmetrics.registerFont(TTFont('Good Vibes', 'font file path goes here'))
    pdfmetrics.registerFont(TTFont('CMU Serif', 'font file path goes here'))
    
    c = canvas.Canvas("report.pdf")
    c.setPageSize((768, 1024))
    c.setFont('Good Vibes', 18)
    
    styles = getSampleStyleSheet()
    style = styles['Title']
    style.fontName = 'CMU Serif'
    style.fontSize = 32
    style.textColor = '#67999E'
    
    # Drawing the image
    c.drawInlineImage(r'name the file and the path location here', 0, 0, preserveAspectRatio=True)
    
    #Create the Store name text
    items = []
    items.append(Paragraph(store, style))
    
    # Create a Frame for the store name text
    f = Frame(0, inch*12.95, inch*10.66, inch, showBoundary=0)
    f.addFromList(items, c)
    
    #Create the owner name text
    items = []
    items.append(Paragraph(owner, style))
    
    # Create a Frame for the owner name text
    f = Frame(0, inch*12.45, inch*10.66, inch, showBoundary=0)
    f.addFromList(items, c)
    
    
    #Create the ad link
    items = []
    ad_link_link = ('<link href={}>Click here to view your BB ad</link>').format(ad_link)
    items.append(Paragraph(ad_link_link, style))
    
    # Create a Frame for the ad link
    f = Frame(0, inch*10.1, inch*10.66, inch, showBoundary=0)
    f.addFromList(items, c)
    
    """Change font back here"""
    style.fontSize = 46
    style.textColor = '#80BDC4'
    
    #Create the month text
    items = []
    items.append(Paragraph(current_month, style))
    
    # Create a Frame for the month text
    f = Frame(0, inch*10.9, inch*10.66, inch, showBoundary=0)
    f.addFromList(items, c)
    
    """Change font back here"""
    style.fontSize = 28
    style.fontName = 'Good Vibes'
    style.textColor = 'Black'
    
    #Create the reach text
    items = []
    reach = str(reach)
    items.append(Paragraph(reach, style))
    
    # Create a Frame for the reach text
    f = Frame(inch*1.4, inch*6.9, inch*2, inch, showBoundary=0)
    f.addFromList(items, c)
    
    #Create the frequency text
    items = []
    frequency = str(frequency)
    items.append(Paragraph(frequency, style))
    
    # Create a Frame for the frequency text
    f = Frame(inch*4.6, inch*6.9, inch*2, inch, showBoundary=0)
    f.addFromList(items, c)
    
    #Create the impressions text
    items = []
    impressions = str(impressions)
    items.append(Paragraph(impressions, style))
    
    # Create a Frame for the frequency text
    f = Frame(inch*7.99, inch*6.9, inch*2, inch, showBoundary=0)
    f.addFromList(items, c)
    
    """Change the font back here"""
    style.fontSize = 22
    style.fontName = 'CMU Serif'
    style.textColor = '#67999E'
    
    #Create the facebook link
    facebook = []
    facebook_link = '<link href="https://www.facebook.com/BridalBoutiques.us/">Link to our FaceBook page</link>'
    facebook.append(Paragraph(facebook_link, style))
    
    # Create a Frame for the Facebook link
    f = Frame(0, inch*5.5, inch*10.66, inch/2, showBoundary=0)
    f.addFromList(facebook, c)
    
    #Create the Instagram link
    instagram = []
    instagram_link = '<link href="https://www.instagram.com/bridalboutiques.us/">Link to our Instagram page</link>'
    instagram.append(Paragraph(instagram_link, style))
    
    # Create a Frame for the Instagram link
    f = Frame(0, inch*5, inch*10.66, inch/2, showBoundary=0)
    f.addFromList(instagram, c)
    
    #Create the Blog link
    blog = []
    blog_link = '<link href="https://bridalboutiques.us/blog/">Link to BB.US Blog</link>'
    blog.append(Paragraph(blog_link, style))
    
    # Create a Frame for the Blog link
    f = Frame(0, inch*4.5, inch*10.66, inch/2, showBoundary=0)
    f.addFromList(blog, c)
    
    #Create the BB page link
    page = []
    page_link = ('<link href={}>Link to your BB.US page</link>').format(BB_page)
    page.append(Paragraph(page_link, style))
    
    # Create a Frame for the BB page link
    f = Frame(0, inch*4, inch*10.66, inch/2, showBoundary=0)
    f.addFromList(page, c)
    
    #Create the announcements link
    announcements = []
    announcements_link = '<link href="https://www.facebook.com/groups/1606839309617584/announcements/">Link to Important Announcements</link>'
    announcements.append(Paragraph(announcements_link, style))
    
    # Create a Frame for the announcements link
    f = Frame(0, inch*3.5, inch*10.66, inch/2, showBoundary=0)
    f.addFromList(announcements, c)
    
    c.save()
    return('report.pdf')
        
    
def main():
    subject = 'Monthly Bridal Boutiques Service report'
    body = 'I hope you are doing well!\nAttached is your November ad report!\n\nPlease let me, Jacqui, or Linda know if you want to go over your results!\n'
    no_ads = []
    for i in range(len(monthly['STORE NAME'])):
        if monthly['Ad Rep'][i] != 'NA' and monthly['Report'][i] == '':
            info = get_info(i)
            if info == None:
                no_ads.append(monthly['STORE NAME'][i] + ' does not have an ad or the ad name is entered incorrectly in Facebook Ads Manager')
                continue
            report = make_report(info[0], info[1], info[2], info[3], info[4], info[5], info[6])
            #print(monthly['STORE NAME'][i])
            body_sent = (body, report)
            send_email(i, subject, body_sent)
            #IMPORTANT!! Update when you are ready to send the emails
            monthly['Report'][i] = 'Done'
        else:
            no_ads.append(monthly['STORE NAME'][i])
            
    #Write the file - make sure this is the correct file path and that you are sending the emails
    """with pd.ExcelWriter(monthly_excel, engine='xlsxwriter') as writer:
        monthly.to_excel(writer, sheet_name='Sheet1')"""
    
    print('These stores did not recieve a report')
    for ad in no_ads:
        print(ad)
    
if __name__ == '__main__':
    main()
    
    
    
    
    