# Read the mails in a folder of an outlook account and send
# an email with the caount of last week
import win32com.client
import datetime
import pytz

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
account = outlook.Folders["Your_Account_Name"]
inbox = account.Folders["Your_Forder"]

# count emails from last week
one_week_ago = datetime.datetime.now(pytz.utc) - datetime.timedelta(days=14)
emails_from_last_week = len(list(filter(lambda x: x.ReceivedTime > one_week_ago, inbox.Items)))

# send email with the count of email from last week
recipient = "Your_recipients_accounts_address"
subject = "Count of emails from last week"
body = f"There are {emails_from_last_week} emails from last week."

# Send email
message = win32com.client.Dispatch("Outlook.Application").CreateItem(0)
message.To = recipient
message.Subject = subject
message.Body = body
message.Send()