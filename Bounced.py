import imaplib
import email
from email.header import decode_header
import pandas as pd

# Function to check bounced emails for a given account
def check_bounced_emails(email_account, password, bounce_subjects):
    try:
        # Hostinger IMAP server details
        imap_server = 'imap.hostinger.com'
        imap_port = 993

        # Connect to the Hostinger email server and login
        mail = imaplib.IMAP4_SSL(imap_server, imap_port)
        mail.login(email_account, password)
        
        # Select the inbox
        mail.select('inbox')
        
        # Search for unread emails
        status, messages = mail.search(None, '(UNSEEN)')
        mail_ids = messages[0].split()
        
        bounced_emails = []
        
        for mail_id in mail_ids:
            # Fetch the email message by ID
            status, msg_data = mail.fetch(mail_id, '(RFC822)')
            
            # Get the email content
            msg = email.message_from_bytes(msg_data[0][1])
            
            # Decode the email subject
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else 'utf-8')
            
            # Check if the subject matches any of the bounce subjects
            if any(bounce_subject.lower() in subject.lower() for bounce_subject in bounce_subjects):
                # Mark email as read
                mail.store(mail_id, '+FLAGS', '\\Seen')
                
                # Extract the bounced email address from the email body
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == "text/plain":
                            body = part.get_payload(decode=True).decode()
                            if "Final-Recipient" in body:
                                bounced_email = body.split("Final-Recipient: rfc822;")[1].split('\r\n')[0]
                                bounced_emails.append(bounced_email)
                else:
                    body = msg.get_payload(decode=True).decode()
                    if "Final-Recipient" in body:
                        bounced_email = body.split("Final-Recipient: rfc822;")[1].split('\r\n')[0]
                        bounced_emails.append(bounced_email)

        # Close the connection and logout
        mail.close()
        mail.logout()
        
        return bounced_emails
    
    except Exception as e:
        print(f"Failed to check {email_account}: {str(e)}")
        return []

# Load the Excel file with email credentials
excel_file = 'bounce_check.xlsx'  # Replace with your Excel file path
credentials_df = pd.read_excel(excel_file)

# List of subjects indicating bounced emails
bounce_subjects = ["Mail Delivery Failed", "Undelivered Mail Returned to Sender", "Mail delivery failure notice", "Delivery Status Notification (Failure)"]

# DataFrame to store bounced emails with additional information
bounced_df = pd.DataFrame(columns=["Email Account", "Bounced Email", "Source Email Account"])

# Iterate through each email account in the Excel file
for index, row in credentials_df.iterrows():
    email_account = row['Email']
    password = row['Password']
    
    print(f"Checking bounced emails for: {email_account}")
    
    # Check for bounced emails
    bounced_emails = check_bounced_emails(email_account, password, bounce_subjects)
    
    # Append results to DataFrame with the source email account
    for bounced_email in bounced_emails:
        bounced_df = bounced_df.append({"Email Account": email_account, 
                                        "Bounced Email": bounced_email,
                                        "Source Email Account": email_account}, 
                                       ignore_index=True)

# Save bounced emails to a new Excel file with additional information
bounced_df.to_excel("bounced_emails_report_with_source.xlsx", index=False)
print("Bounced emails have been saved to 'bounced_emails_report_with_source.xlsx'")
