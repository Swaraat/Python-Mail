import pandas as pd
import smtplib
from time import sleep
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import random

# Load email and credential data from Excel
email_list_path = "email_list1.xlsx"  # Path to your target email Excel file
credentials_path = "credentials1.xlsx"  # Path to your credentials Excel file

# Read Excel sheets
df1 = pd.read_excel(email_list_path)
df2 = pd.read_excel(credentials_path)

# Check if 'status' column exists; if not, create it
if 'status' not in df1.columns:
    df1['status'] = ''  # Initialize 'status' column with empty strings

# Create a separate DataFrame with only unsent emails for processing
unsent_emails_df = df1[df1['status'] != 'sent']

# Declare variables
total_email_to_send = len(unsent_emails_df)
email_sent = 0
credential_index = 0
emails_per_credential = 4
email_by_each_cred = emails_per_credential

# Fixed Details
subject = "Enhance Your Contract Management with XEqualTo Analytics"
company_website = "https://xequalto.com/"
company_name = "XEqualTo"
your_phone = "+91 9147059111"

# Email body template with HTML formatting
body_template = """\
Hi {name},<br><br>
XEqualTo Analytics offers a Contract Analytics Solution to help your legal team streamline contract management, reduce risks, and improve decision-making.<br><br>
Here’s how we can assist:<br>
<ul>
    <li><strong>30% Faster Contract Reviews:</strong> Automate routine tasks, so your team can focus on critical issues.</li>
    <li><strong>25% Better Risk Management:</strong> Identify legal risks early with advanced clause detection and risk scoring.</li>
    <li><strong>40% Improved Compliance:</strong> Automatically track contract terms and ensure obligations are met.</li>
    <li><strong>20% Increased Contract Visibility:</strong> Organize contracts in a searchable database for easier access.</li>
    <li><strong>15% Faster Negotiations:</strong> Identify bottlenecks and speed up deal closures.</li>
</ul>
Take control of your contracts and mitigate risks with our solution. Book your free consultation <a href="https://calendly.com/ad-xequalto">here</a>.<br><br>

Best regards,<br>
{signature}<br>
{your_name}<br>
{your_position} | {company_name}<br>
ph: {your_phone}
"""

def send_mail(to_name, to_email, from_email, from_password, your_name, your_position, signature):
    try:
        # Set up the SMTP server for Hostinger
        smtp_server = "smtp.hostinger.com"  # Hostinger SMTP server
        smtp_port = 587  # Hostinger SMTP port

        # Connect to the server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(from_email, from_password)

        # Create the email
        msg = MIMEMultipart()
        msg['From'] = f"{your_name} <{from_email}>"
        # msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        
        # Format the body with HTML content
        body = body_template.format(
            name=to_name,
            your_name=your_name,
            your_position=your_position,
            company_name=company_name,
            your_phone=your_phone,
            signature=signature
        )
        msg.attach(MIMEText(body, 'html'))  # Use 'html' instead of 'plain'

        # Send the email
        server.send_message(msg)
        server.quit()

        print(f"Email sent to {to_email} from {from_email}")
        return "sent"
    except Exception as e:
        print(f"Failed to send email to {to_email} from {from_email}: {str(e)}")
        return "failed"

# Main execution
for index, row in unsent_emails_df.iterrows():
    if total_email_to_send > 0 and email_sent < total_email_to_send:
        # Get recipient details
        to_name = row['name']
        to_email = row['email']

        # Get sender details from credentials
        creds = df2.iloc[credential_index]
        from_email = creds['email']
        from_password = creds['password']
        your_name = creds['your_name']
        your_position = creds['your_position']
        signature = creds['signature']

        # Send the email
        status = send_mail(to_name, to_email, from_email, from_password, your_name, your_position, signature)

        # Update the original DataFrame (df1)
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Get current time
        df1.at[index, 'status'] = status
        df1.at[index, 'sent_from'] = from_email
        df1.at[index, 'sent_time'] = current_time  # Add time of sending email
        df1.to_excel(email_list_path, index=False)  # Save original DataFrame back to Excel

        # Print the result with time
        print(f"Status updated: {to_email} - {status} from {from_email} at {current_time}")

        # Increment counters
        email_sent += 1
        email_by_each_cred -= 1

        # Check if the current email limit per credential is reached
        if email_by_each_cred == 0:
            credential_index += 1
            email_by_each_cred = emails_per_credential
        
        # Wait for a random time between 60 to 300 seconds before sending the next mail
        sleep_time = random.randint(10, 30)
        print(f"Sleeping for {sleep_time} seconds before sending the next email.")
        sleep(sleep_time)

print('All mail sent')
