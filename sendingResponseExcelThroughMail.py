import smtplib
from email.message import EmailMessage
import getpass


def sendMail(recipientMail,excelFileName):
    recipient_email = recipientMail
    sender_email = "laxminayanan546@gmail.com"
    app_password = "yqjbteinioazgqgt"  # Using the 16-digit app password from Google

    subject = "Response Sheet Of The Math Escape Room"
    body = "Please find the attached Excel file."

    msg = EmailMessage()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.set_content(body)


    excel_file_path = excelFileName  
    with open(excel_file_path, 'rb') as f:
        file_data = f.read()
        file_name = f.name

    msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, app_password)
            smtp.send_message(msg)
        return 0
    except Exception as e:
        return -1
    
