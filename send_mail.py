#test
import smtplib
from email.message import EmailMessage
from tabulate import tabulate


def send_email_with_attendance(attendance, email_to,team, login_count,logout_count,subject=None):    # Email configuration
    smtp_host = 'mail.egyptpost.org'
    smtp_port = 25  # Update the port if necessary
    sender_email = 'W_Abdelrahman.Ataa@EgyptPost.Org'
    recipient_emails = email_to
    if subject == None:
        subject = 'Attendance Report'

    # Convert attendance data to a tabulated HTML table with custom CSS styling
    table = tabulate(attendance, headers="keys", tablefmt="html")
    table_with_custom_styles = f'''
    <style>
        table {{
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
            margin: 0 auto;
        }}
        th, td {{
            border: 1px solid #dddddd;
            text-align: center;
            padding: 8px;
        }}
        th {{
            background-color: #f2f2f2;
            color: #333;
        }}
        tr:nth-child(even) {{
            background-color: #f2f2f2;
        }}
    </style>
    <table>
        {table}
    </table>
    '''
    additional_content = f'''
    <p><strong>Login Count:</strong> {login_count}</p>
    <p><strong>Logout Count:</strong> {logout_count}</p>
    '''
    
    # Create the email message with the HTML table and additional content
    body_message = f'''
    <p>Dear {team},</p>
    {additional_content}
    <p>Here's the attendance report:</p>{table_with_custom_styles}
    <p>Best regards,</p>
    <p>Abdelrahman Ataa</p>
    <p>Soc Engineer</p>
    '''
    message = EmailMessage()
    message.set_content(body_message, subtype='html')
    message['Subject'] = subject
    message['From'] = sender_email
    message['To'] = ', '.join(recipient_emails)

    # Create the SMTP connection
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.set_debuglevel(1)  # Enable debug output for troubleshooting
        server.ehlo()
        server.sendmail(sender_email, recipient_emails, message.as_string())



email_to_soc = []#recever mails here 

email_to_noc = [] #recever mails here 

email_to_helpdesk = [ ] ##recever mails here 
