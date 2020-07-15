import openpyxl
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

wb=openpyxl.load_workbook("ex.xlsx")
s=wb["sheet1"]
data=s['A']
sender_email = input("Enter your Gmail: ")
password =input("Enter your password: ")
message = MIMEMultipart("alternative")
message["Subject"] = input("Enter subject: ")
mesgConant=input("Enter message: ")

for i in range(len(data)):
    receiver_email = data[i].value
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Bcc"] = receiver_email
    html = """\
    <html>
      <body>
        <h3 style="color:green;">"""+mesgConant+"""
        </h3>
	<b style="color:red;">Best Regards<br>Eng.Ayah Alrifai</b>
      </body>
    </html>
    """
    part = MIMEText(html, "html")
    message.attach(part)

    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(
            sender_email, receiver_email, message.as_string()
        )
