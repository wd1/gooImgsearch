#Please read Readme first
from distutils.debug import DEBUG
import os.path
import pandas as pd
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
#get contact data
file = input("Enter the CONTACT URL CSV FILE PATH :  ")

df_contact = pd.read_csv(file,engine="c")
contact_index = df_contact.columns

#get url file path
file = input("Enter the URL EXCEL FILE PATH :  ")

load_dotenv()

sender_address = os.getenv('EMAIL')
sender_pass = os.getenv('PWD')
mail_content = os.getenv('MAIL_CONTENT')
mail_subject = os.getenv('MAIL_SUBJECT')
#message Setup
message = MIMEMultipart()
message['From'] = sender_address
message['Subject'] = mail_subject
message.attach(MIMEText(mail_content, 'plain'))

df = pd.read_excel(file,engine="openpyxl")
#get corresponding contact and send Email
#Create session
# session = smtplib.SMTP("smtp.gmail.com", 587) 
# #start session
# session.starttls() 
# session.login(sender_address, sender_pass) #log in
index = df.columns
for i in range(0,len(df.index)):
    #get row data
    row = df.loc[i]
    #get url in COL_A
    url = row[index[0]].strip()
    if(url.startswith('http://')):
        first_slash = url.find('/',8)
        url = url[8:first_slash]
    elif(url.startswith('https://')):
        first_slash = url.find('/',9)
        url = url[9:first_slash]
    else:
        first_slash = url.find('/')
        url = url[0:first_slash]
    #get corresponding contact 
    result_row = df_contact[df_contact[contact_index[0]].str.contains(url)]
    contact = result_row[contact_index[1]]
    print(contact)
    #send EMail
    message.attach(MIMEText(mail_content, 'plain'))    
    text = message.as_string()
    #session.sendmail(sender_address, contact, text)
    print('EMail has been sent to :' + contact)
#session.quit()
print('End')