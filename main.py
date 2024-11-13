import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

data = pd.read_excel(r"C:\Users\bvenk\Downloads\Mail Sender\contacts.xlsx", engine='openpyxl')

SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587  
EMAIL_ADDRESS = 'your mail id'
EMAIL_PASSWORD = 'your password'  

server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
server.starttls() 

# Login to the email server using your email and app password
server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

# Loop through each row in the Excel file
for index, row in data.iterrows():
    recipient_email = row['Email']
    recipient_name = row['Name']
    custom_message = row['Message']
    
    # Compose email
    message = MIMEMultipart()
    message['From'] = EMAIL_ADDRESS
    message['To'] = recipient_email
    message['Subject'] = f"Welcome Message for {recipient_name}"

    # Customize email body
    body = f"Dear {recipient_name},\n\n{custom_message}\n\nBest regards,\nYours Venkataramana\nKL University"
    message.attach(MIMEText(body, 'plain'))
    
    # Send email
    server.sendmail(EMAIL_ADDRESS, recipient_email, message.as_string())
    print(f"Email sent to {recipient_name} at {recipient_email}")

server.quit()
print("All emails sent successfully.")
