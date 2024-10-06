import pandas as pd
import smtplib
from time import sleep
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Load email and credential data from Excel
email_list_path = "email_list1.xlsx"  # Path to your target email Excel file
credentials_path = "credentials1.xlsx"  # Path to your credentials Excel file

# Read Excel sheets
df1 = pd.read_excel(email_list_path)
df2 = pd.read_excel(credentials_path)

# Check if 'status' column exists; if not, create it
if 'status' not in df1.columns:
    df1['status'] = ''  # Initialize 'status' column with empty strings

# Filter rows where status is not 'sent'
df1 = df1[df1['status'] != 'sent']

# Declare variables
total_email_to_send = len(df1)
email_sent = 0
credential_index = 0
emails_per_credential = 150
email_by_each_cred = emails_per_credential

# Fixed Details
subject = "We Can Enhance Your Data Strategy"
company_website = "https://xequalto.com/"
#your_position = "Sr. Business Consultant"
company_name = "XEqualTo"
your_phone = "+91 9147059111"

# Email body template
body_template = """\
Dear {name},

I hope you’re doing well.

I’m {your_name}, {your_position} at {company_name}. Our team is dedicated to helping organizations like yours turn data into valuable insights that drive better decision-making and strategic growth. At {company_name}, we simplify data into actionable insights with clear and effective tools and solutions.
We offer:

* Data Modelling: Designing advanced data warehouses and implementing data mesh strategies to enhance data quality and accessibility.
* Data Analytics & Visualization: Transforming data into clear, actionable insights.
* Snowflake Consulting: Providing expertise in Snowflake for scalable, secure, and efficient data warehousing and analytics.

I would love to discuss how we can support your data initiatives. Could we schedule a meeting? Please let me know your availability.
Looking forward to connecting with you.

Best,
{your_name}
{your_position}
{company_name}
ph: {your_phone}
"""

def send_mail(to_name, to_email, from_email, from_password, your_name):
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
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        
        # Format the body
        body = body_template.format(
            name=to_name,
            your_name=your_name,
            company_name=company_name,
            your_position=your_position,
            company_website=company_website,
            your_phone=your_phone
        )
        msg.attach(MIMEText(body, 'plain'))

        # Send the email
        server.send_message(msg)
        server.quit()

        print(f"Email sent to {to_email} from {from_email}")
        return "sent"
    except Exception as e:
        print(f"Failed to send email to {to_email} from {from_email}: {str(e)}")
        return "failed"

# Main execution
for index, row in df1.iterrows():
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

        # Send the email
        status = send_mail(to_name, to_email, from_email, from_password, your_name)

        # Update the DataFrame e next email
        #sleep(60)and Excel sheet
        df1.at[index, 'status'] = status
        df1.at[index, 'sent_from'] = from_email
        df1.to_excel(email_list_path, index=False)

        # Print the result
        print(f"Status updated: {to_email} - {status} from {from_email}")

        # Increment counters
        email_sent += 1
        email_by_each_cred -= 1

        # Check if the current email limit per credential is reached
        if email_by_each_cred == 0:
            credential_index += 1
            email_by_each_cred = emails_per_credential
        
        # Wait for 1 minute before sending the next mail
        sleep(180)
print('All mail sent')