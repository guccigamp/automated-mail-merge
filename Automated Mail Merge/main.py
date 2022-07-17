import smtplib
import os
import pandas as pd
import automation
from os.path import basename
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

from_addr = 'guccigamp@gmail.com'

data = pd.read_csv('input/team_directory.csv')

subject = open(file='input/subject.txt', mode='r').read()


#checking for any duplicate entry
set_of_emails = set()
for index, row in data.iterrows():
    if row["Email"] in set_of_emails:
        data["Flag"][index] = 'sent(duplicate)'
    else:
        if row["Flag"] == "unsent":
            set_of_emails.add(row["Email"])
data.to_csv('input/team_directory.csv')
print(set_of_emails)
list_of_emails = list(set_of_emails)

#attaches the from_addr, to_addr, subject, and the body of the mail
text_content = open(file="input/text.txt", mode='r').read()
for email in list_of_emails:
    to_addr = email
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Subject'] = subject
    body = MIMEText(text_content, 'plain')
    msg.attach(body)
#if the attachments directory is not empty
    if os.listdir('input/attachments') is not []:
        #attaching multiple attachments
        filenames = os.listdir('input/attachments')
        file_directories = []
        for file_dir in filenames:
            file_directories.append(os.path.join('input/attachments', file_dir))
        for filename in file_directories:
            with open(file=filename, mode='rb') as f:
                part = MIMEApplication(f.read(), basename(filename))
                part['Content-Disposition'] = 'attachment; filename="{}"'.format(basename(filename))
            msg.attach(part)


    #setting up the SMTP
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('guccigamp@gmail.com','njkknjwfqahhyoiw')
    server.send_message(msg, from_addr=from_addr, to_addrs=[to_addr])

    #flagging the newly sent emails
    for index,row in data.iterrows():
        if row["Email"] == to_addr and row["Flag"] == "unsent":
            data["Flag"][index] = "sent"
    data.to_csv('input/team_directory.csv')

server.quit()

automation.automate()