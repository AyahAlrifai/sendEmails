
password ="HaYa.IaFiRlA.97"
import openpyxl
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

wb=openpyxl.load_workbook("ex.xlsx")
s=wb["sheet1"]
data=s['A':'C']

sender_email = "alrefayayah@gmail.com"
names=data[0]
grades=data[1]
emails=data[2]

for i in range(len(names)):
    if grades[i].value>=3.75:
        emoji="&#x1F600;"
    else:
        emoji="&#x1F622;"
    receiver_email = emails[i].value
    message = MIMEMultipart("alternative")
    message["Subject"] = "your grade"
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Bcc"] = receiver_email
    html = """\
    <html>
      <body>
        <h3 style="color:red;background-color:green">
        Hi """+names[i].value+"""<br>
        your grade is """+str(grades[i].value)+emoji+"""<br>
        good lock!<br>
        Ayah Al-refai
        </h3>
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
