import smtplib
import csv
import ssl
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# User configuration
sender_email = '' #enter your email address
password = '' #enter your password
sender_name = '' #enter sender name

receiver_emails = []
receiver_names = []

# For email - ids
with open('ICAP_members_irex_e-6-ID.csv', 'r') as file: #update your file csv file name that contains email IDs
    reader = csv.reader(file, delimiter=',')
    for row in reader:
        receiver_emails.extend(row)
print(receiver_emails)

# For email - name
with open('ICAP_members_irex_e-6-name.csv', 'r') as file: #update your file csv file name that contains names
    reader = csv.reader(file, delimiter=',')
    for row in reader:
        receiver_names.extend(row)
print(receiver_names)

# Email body
email_html = open('irex.html') #write your html file name that contains the email's body
email_body = email_html.read()

filename = 'payload.pdf' #write pdf file name that you want to send as attachment

# Counter to keep track of the number of emails sent
email_count = 0

for receiver_email, receiver_name in zip(receiver_emails, receiver_names):
    print("Sending the email...")
    # Configurating user's info
    msg = MIMEMultipart()
    msg['To'] = formataddr((receiver_name, receiver_email))
    msg['From'] = formataddr((sender_name, sender_email))
    msg['Subject'] = 'E-Commerce Web Solutions | Digital Marketing | ERP | Startup Solutions'

    msg.attach(MIMEText(email_body, 'html'))

    try:
        # Open PDF file in binary mode
        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        # Encode file for ASCII characters
        encoders.encode_base64(part)
        # Add header to the email as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )
        msg.attach(part)
    except Exception as e:
        print(f'Oh no! We didn\'t find the attachment!\n{e}')
        break

    try:
        # Creating a SMTP session using port 587 with TLS for Titan email
        server = smtplib.SMTP('smtp.titan.email', 587)
        # Encrypts the email
        context = ssl.create_default_context()
        server.starttls(context=context)
        # We log in to our email account
        server.login(sender_email, password)

        # from sender to receiver with the email attachment of html body
        server.sendmail(sender_email, receiver_email, msg.as_string())
        email_count += 1
        print(f'Email sent! Total emails sent: {email_count}')
    except Exception as e:
        print(f'Oh no! Something bad happened!\n{e}')
    finally:
        print('Closing the server...')
        if 'server' in locals():
            server.quit()
