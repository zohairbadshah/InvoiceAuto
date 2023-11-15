import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


class Mail:
    sender_email = ''
    sender_password = ''
    receiver_email = ''
    subject = 'Mismatch in data'

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject


    @staticmethod
    def send_mail(message):
        Mail.msg.attach(MIMEText(message, 'plain'))

        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()

            server.login(Mail.sender_email, Mail.sender_password)

            # Send the email
            server.sendmail(Mail.sender_email, Mail.receiver_email, Mail.msg.as_string())
            print('Email sent successfully')
            server.quit()

        except Exception as e:
            print(f'Error: {str(e)}')

