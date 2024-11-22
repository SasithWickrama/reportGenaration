import cx_Oracle
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def connect_oracle(username, password, host, port, service_name):
    dsn_tns = cx_Oracle.makedsn(host, port, service_name=service_name)
    return cx_Oracle.connect(username, password, dsn_tns)

def execute_query(connection, query):
    cursor = connection.cursor()
    cursor.execute(query)
    return cursor.fetchall()

def write_excel(data, col_count, path):
    # Example: Using pandas to write data to Excel
    import pandas as pd

    df = pd.DataFrame(data[1:], columns=data[0])
    df.to_excel(path, index=False)

def send_email(to_addresses, subject, body, attachment_path):
    from_email = "your_email@gmail.com"  # Your email address
    from_password = "your_email_password"  # Your email password

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = ", ".join(to_addresses)
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    # Attach Excel file
    attachment = open(attachment_path, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={attachment_path}')
    msg.attach(part)

    # Connect to Gmail's SMTP server
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, from_password)

    # Send email
    server.sendmail(from_email, to_addresses, msg.as_string())
    server.quit()

def main():
    host = "your_oracle_host"
    username = "sasith_rpt"
    password = "Sasith#rpt"
    port = 1521
    service_name = "your_service_name"

    array_to = ["sasith@slt.com.lk"]

    connection = connect_oracle(username, password, host, port, service_name)

    # Execute Oracle queries and process data
    query = "SELECT * FROM your_table"  # Replace with your actual query
    data = execute_query(connection, query)

    # Generate Excel
    excel_path = "path_to_excel_file.xlsx"
    write_excel(data, len(data[0]), excel_path)

    # Prepare email content
    email_body = "Your email body content here."
    subject = "Pending Fiber Planned Event Task Details Report"

    # Send email
    send_email(array_to, subject, email_body, excel_path)

if __name__ == "__main__":
    main()
