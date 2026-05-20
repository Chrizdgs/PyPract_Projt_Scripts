# Read the emails in a folder of an outlook account and send an email with the count of mails from last week
# with a table of those mails with this format (Date, Subject, Recipients, Body)
import win32com.client
import datetime
import pandas as pd
import pytz
import re

def read_outlook_emails(folder_name):
    #Init the outlook client
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    account = outlook.Folders["Your_account_name"]
    inbox = account.Folders[folder_name]
    #set the dates to check
    last_week = datetime.datetime.now(pytz.utc) - datetime.timedelta(days=8)
    emails = []
    for email in inbox.Items:
        if email.ReceivedTime > last_week:
            emails.append({
                'Date received': email.ReceivedTime.replace(tzinfo=None).astimezone(pytz.utc), #set a format to the dates
                'Subject': email.Subject,
                'Recipients': email.To,
                'Body':  re.sub(r'\r\n|\r|\n', ' ', re.sub(r'<[^>]+>', '',email.Body)), #clean the mail body
            })
    return emails

def send_email_with_count(recipient, emails):
    message = win32com.client.Dispatch("Outlook.Application").CreateItem(0) #create the mail
    #assign recipients, subject and body
    message.To = recipient
    message.Subject = "Business Summary Errors of Last week"
    message.Body = f"There are {len(emails)} emails from last week.\n\n -Python AutoEmail"
    # Paste the table of emails in the body
    table = pd.DataFrame(emails)    
    message.HTMLBody = message.HTMLBody + table.to_html(index=False)
    message.Send()

if __name__ == "__main__":
    emails = read_outlook_emails("Your_forder_Name")
    send_email_with_count("Your_recepient_account_address", emails)