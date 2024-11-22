import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def send_mail(to, to_a, to_b, sender, subject, text_msg, attach_name, attach_mime, attach_clob, smtp_host, smtp_port=25, smtp_username=None, smtp_password=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = to
        msg['Cc'] = to_b
        msg['Subject'] = subject
        msg.attach(MIMEText(text_msg, 'html'))

        if attach_name and attach_mime and attach_clob:
            # Assuming attach_clob is the content of the attachment
            attachment = MIMEApplication(attach_clob, _subtype=attach_mime)
            attachment.add_header('Content-Disposition', 'attachment', filename=attach_name)
            msg.attach(attachment)

        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.sendmail(sender, [to, to_a, to_b], msg.as_string())

        
        print("Email sent successfully!")
    
    except smtplib.SMTPException as e:
        print(f"Failed to send email. Error: {e}")
    
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

def main():
    send_mail('lumbini@slt.com.lk', 'sasith@slt.com.lk', 'sasith@slt.com.lk', 'sasith@slt.com.lk', 'This is TEST Report', '<p>Test HTML Body</p>', 'attachment.txt', 'plain', 'Attachment Content', '172.25.2.207', 25, 'sasith_rpt', 'Sasith#rpt')

if __name__ == "__main__":
    main()