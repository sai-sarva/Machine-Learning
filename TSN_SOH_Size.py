import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# define email parameters
FROM_ADDRESS = "supplychainOceania@brightstar.com"
EMAIL_PASSWORD = "W98h4XSNaa"
TO_ADDRESS = ["analyticsandinsights.au@likewize.com", "alice.nguyen@likewize.com",
              "mark.hymers@likewize.com"]
CC_ADDRESS = ["rhea.ocoma@likewize.com"]
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

# The path of your directory
directory_path = 'C:/Users/sayan/Likewize/Operations Data Warehouse - Documents/CUBE_TSN Extract/SOH/'

# Find the most recent file in the directory
files = os.listdir(directory_path)
paths = [os.path.join(directory_path, basename) for basename in files]
latest_file = max(paths, key=os.path.getctime)

# Check the file size
file_size = os.path.getsize(latest_file) / 1024  # size in KB

# If the size is less than 4KB
if file_size < 10:
    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = FROM_ADDRESS
    message['To'] = ', '.join(TO_ADDRESS)
    message['CC'] = ', '.join(CC_ADDRESS)
    message['Subject'] = 'File Size Alert'

    # The body and the attachments for the mail
    mail_content = f'The size of the most recent file {os.path.basename(latest_file)} Stock On Hand is less than 4KB. Please check if the Cube Data Refreshed!!'
    message.attach(MIMEText(mail_content, 'plain'))

    # Create SMTP session for sending the mail
    session = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)  # Use smtp server for Office 365
    session.starttls()  # Enable security
    session.login(FROM_ADDRESS, EMAIL_PASSWORD)  # Login with mail_id and password
    text = message.as_string()
    session.sendmail(FROM_ADDRESS, TO_ADDRESS, text)
    session.quit()