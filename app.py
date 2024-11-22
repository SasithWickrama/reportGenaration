import cx_Oracle
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta  # Import datetime and timedelta

def send_mail(to, sender, subject, text_msg, attach_name, attach_mime, attach_clob, smtp_host, smtp_port=25):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = to
        msg['Cc'] = to_b
        msg['Subject'] = subject
        msg.attach(MIMEText(text_msg, 'html'))

        if attach_name and attach_mime and attach_clob:
            attachment = MIMEApplication(attach_clob.read(), _subtype=attach_mime)
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
    print("@1")
    try:
        dsn_tns = cx_Oracle.makedsn("172.25.1.172", 1521, service_name="clty")
        connection = cx_Oracle.connect("sasith_rpt", "Sasith#rpt", dsn_tns)
    except Exception as e:
        print("Exception @1 : %s" % traceback.format_exc())
        logger.info("Exception : %s" % traceback.format_exc())
        return traceback.format_exc()

    cursor = connection.cursor()
    cursor.execute("SELECT COUNT(*) FROM CLARITY_ADMIN.SO_INFO_REPORT")
    v_ren_CNT = cursor.fetchone()[0]
    print("Exception @2 : %s" % v_ren_CNT)
    cursor.close()

    if v_ren_CNT == 0:
        get_ip_of_database = 'clarityn1'

        if get_ip_of_database in ['clarityn1', 'clarityn2']:
            cursor = connection.cursor()
            current_date = datetime.now()
            start_date = current_date - timedelta(days=2)
            end_date = current_date - timedelta(days=1)
            cursor.callproc("CLARITY_ADMIN.SO_REPORT_DATA", [start_date, end_date])
            cursor.close()

            #cursor = connection.cursor()
            #cursor.execute("DELETE FROM CLARITY_ADMIN.SO_INFO_REPORT")
            #cursor.close()

            l_body = '<html><head><title>Service Order Detail Report' + \
                '</title></head><body>' + \
                "\n" + \
                "\n" + \
                '<table width="100%" cellpadding="10" cellspacing="0">' + \
                "\n" + \
                "\n" + \
                '<tr>' + \
                "\n" + \
                "\n" + \
                '<span style=''color:Blue;font-family:Calibri,Arial,Helvetica,sans-serif''>' + \
                "\n" + \
                "\n" + \
                '<p>Dear ' + \
                'Sir' + \
                '/' + \
                'Madam' + \
                ',' + \
                "\n" + \
                "\n" + \
                '<br><br>' + \
                'Today Service Order Detail Report contain ' + \
                str(v_ren_CNT) + \
                ' records.' + \
                '<br>' + \
                'Please find the attached Service Order Detail Report.' + \
                "\n" + \
                "\n" + \
                "\n" + \
                "\n" + \
                '</table></body></html>'

            send_mail('sasith@slt.com.lk', 'sasith@slt.com.lk', 'This is TEST Report', l_body, 'SO_Report.xls', 'xls', dest_lob, '172.25.2.207')

if __name__ == "__main__":
    main()
