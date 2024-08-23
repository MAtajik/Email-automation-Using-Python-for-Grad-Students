import os
from mimetypes import guess_type

import pandas as pd
import smtplib, ssl, email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import docx2txt

email_df = pd.read_excel('Mother_of_all_emails.xlsx', sheet_name='email')
cv_df = pd.read_excel('Mother_of_all_emails.xlsx', sheet_name='cv')
my_email = {}
my_cv = {}
cv_path = []
email_path = []
receiver_ids = []
subjects = []
transcript_path = "E:\\desktop\\USA AUto\\Academic Transcript.pdf"

# Creating CV Path
for index, row in email_df.iterrows():
    for colum_name in email_df:
        my_email[colum_name] = row[colum_name]
    receiver_ids.append(my_email['Email'])
    subject = f"Prospective Graduate Student {my_email['intended_program']} for Fall'23"
    subjects.append(subject)
    directory = 'Dr.' + my_email['last_name_of_professor']
    parent_dir = r"E:\desktop\USA AUto\Final Email"
    path_cv = os.path.join(parent_dir + '\\' + directory + '\\' + 'Curriculum Vitae.pdf')
    path_email = os.path.join(parent_dir + '\\' + directory + '\\' + 'Email.docx')
    cv_path.append(path_cv)
    email_path.append(path_email)


def attach_files(file):
    # file = os.path.basename(file_path)
    # Open PDF file in binary mode
    with open(file, "rb") as attachment:
        # Add file as application/octet-stream
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)

    #file_name
    file_name = file.split('\\')[-1]

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {file_name}",
    )
    # Add attachment to message and convert message to string
    # message.attach(part)
    # final_message = message.as_string()
    return part


# send email with this function
def send_email(receiver, sub, body, transcript, cv):
    try:
        sender_email = 'minhajcuet13@gmail.com'
        password = 
        context = ssl.create_default_context()
        print('Connecting Server....')
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls(context=context)
        print(f'Server Connected.\nLogging in to Gmail.....')
        server.login(sender_email, password)
        print(f'Log in Successful!\nSending Email to {receiver}.')
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver
        message['Subject'] = sub
        message.attach(MIMEText(body, "plain"))
        message.attach(attach_files(transcript))
        message.attach(attach_files(cv))
        message.as_string()
        server.send_message(message)
        print(f"Email has been sent to {receiver}.")
    except Exception as e:
        print(e)


for i in range(len(cv_path)):
    text = docx2txt.process(email_path[i])
    send_email(receiver=receiver_ids[i], sub=subjects[i], body=text, transcript=transcript_path, cv=cv_path[i])
    # print(attach_files(cv_path[i]))


