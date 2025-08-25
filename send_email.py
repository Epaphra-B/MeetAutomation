import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

def send_excel_email(sender_email, receiver_email, password, excel_buffer, dates, subject="Failed Meetings Data Log", body="Please find the attached Excel file."):

    try:
        # Create the email
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = subject

        fileName = 'meetingsLog_' + dates[0] + '_' + dates[1]

        if excel_buffer == None:
            body = f"No Falied meetings Recorded in this period form {dates[0]} to {dates[1]}"
        else:
            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            excel_buffer.seek(0)
            payload = excel_buffer.read()
            part.set_payload(payload)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{fileName}"')
            msg.attach(part)
            msg.attach(MIMEText(f"{body} that has falied meetings log from {dates[0]} to {dates[1]} with filename: {fileName}"))

        

        # Connect to Gmail SMTP server and send email
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email,receiver_email,msg.as_string())
        server.quit()

        print("Email sent successfully!")

    except Exception as e:
        print(f"Error sending email: {e}")
