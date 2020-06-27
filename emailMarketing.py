# --------------------------------------------------------------------------------------
#
# @Name         : Bulk Email Sender (Python)
# @Author       : Mario Oliveira (mariojgmaster@yahoo.com.br)
# @Description  : This script sends bulk emails using only one excel email list.
# @Version      : 1.0.4
# @Repository   : https://github.com/mjRibeiroTriscal/Bulk-Email-Sender
# @keywords     : digital marketing, bulk email sender, Email Marketing, python, smtp
#
# --------------------------------------------------------------------------------------

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas
import smtplib

e = pandas.read_excel("Emails.xlsx")
emails = e['Emails'].values

gmail_user = "<user_email>"
gmail_pwd = "<user_password>"

msg = MIMEMultipart()
msg['Subject'] = "HTML + StyleSheet + Automation"

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(gmail_user, gmail_pwd)
msg.attach(MIMEText(open('page.html').read(), 'html'))

for email in emails:
    server.sendmail(gmail_user, email, msg.as_string())

server.quit()
print('Emails sent!')
