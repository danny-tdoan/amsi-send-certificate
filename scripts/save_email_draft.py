import csv
import imaplib
import time
import getpass
import base64
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def get_recipients(file):
    recipients = list(csv.reader(open('recipients.csv')))
    return recipients[1:]


def get_template(template):
    return open(template).read()


def save_email_to_draft(subject, message_template, data, from_addr, password):
    """Actual function to compose the draft"""

    details = {}
    details['first_name'] = data[1]
    details['last_name'] = data[2]
    details['email'] = data[3]
    details['attachment'] = data[4]

    message = message_template.format(**details)

    # email
    toaddr = details['email']
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = toaddr
    msg['Subject'] = subject

    body = message
    msg.attach(MIMEText(body, 'plain'))

    attachment = open(details['attachment'], 'rb')
    filename = details['attachment']
    filename = filename[filename.index('/') + 1:]

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(part)

    server = imaplib.IMAP4_SSL('imap-mail.outlook.com', port=993)

    server.login(from_addr, password)
    # msg = {'raw': base64.urlsafe_b64encode(msg.as_string().encode('utf-8'))}
    server.append('Drafts', '\\Draft', imaplib.Time2Internaldate(time.time()), msg.as_bytes())

    print("Email saved to draft for " + details['email'])


def save_email_in_list_to_draft(template, recipients, subject):
    """Compose the email with contents (based on template), to/from, subject and attachment and save the email into draft
    box, awaiting for approval from sender before sending out
    Input:
        1. template as txt
        2. recipients as filename linked to the pre-processed csv containing name, address, attachment for each recipient
        3. subject provided by user

    Output:
        Compose all email with corresponding details, save to Drafts
    """

    message_template = get_template(template)
    recipient_data = get_recipients(recipients)

    # get password
    try:
        fromaddr = input("Enter your full email address: ")
        password = getpass.getpass()
    except Exception as error:
        print('ERROR', error)
        exit()

    for recipient in recipient_data:
        save_email_to_draft(subject, message_template, recipient, fromaddr, password)

    return "All done!"


if __name__ == "__main__":
    # get the message body and recipients
    message_template = sys.argv[1]
    recipients = sys.argv[2]

    email_subject = input("Enter the email subject: ")
    save_email_in_list_to_draft(message_template, recipients, email_subject)
