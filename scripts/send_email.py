import csv
import smtplib
import imaplib
import time
import getpass
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


def send_email(subject, message_template, data, fromaddr, password):
    details = {}
    details['first_name'] = data[1]
    details['last_name'] = data[2]
    details['email'] = data[3]
    details['attachment'] = data[4]
    details['note1'] = data[5]
    details['note2'] = data[6]
    details['note3'] = data[7]
    details['note4'] = data[8]
    details['note5'] = data[9]

    message = message_template.format(**details)

    # email
    toaddr = details['email']
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = subject

    body = message
    msg.attach(MIMEText(body, 'plain'))

    attachment = open(details['attachment'], 'rb')
    filename = details['attachment']
    filename = filename[filename.index('/') + 1:]

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(part)

    # prepare to send
    server = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
    server.starttls()
    server.login(fromaddr, password)

    # prepare to add the email to Sent folder, for the records
    imap_server = imaplib.IMAP4_SSL('imap-mail.outlook.com', port=993)
    imap_server.login(fromaddr, password)

    try:
        # save the email first, then send
        imap_server.append("Sent Items", "", imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        imap_server.logout()

        server.sendmail(fromaddr, toaddr, msg.as_string())
        server.quit()
        print("Email sent to " + details['email'])

    except Exception:
        print("Some error occured ", Exception)
        print("This email might not be sent even if it's present in Sent Inbox")


def send_email_in_list(template, recipients, subject):
    message_template = get_template(template)
    recipients = get_recipients(recipients)

    # get password
    try:
        fromaddr = input("Enter full email address: ")
        password = getpass.getpass()
    except Exception as error:
        print('ERROR', error)
        exit()

    for recipient in recipients:
        send_email(subject, message_template, recipient, fromaddr, password)

    return "All done!"


if __name__ == "__main__":
    # get the message body and recipients
    message_template = sys.argv[1]
    recipients = sys.argv[2]

    email_subject = input("Enter the email subject: ")
    # send_email_in_list('message_template.txt','recipients.csv',email_subject)
    send_email_in_list(message_template, recipients, email_subject)
