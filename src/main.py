import os 
from dotenv import load_dotenv
from email.message import EmailMessage
import ssl
import smtplib
load_dotenv()
email_sender = os.getenv("EMAIL_SENDER")
email_reciver = os.getenv("EMAIL_RECIVER")
password = os.getenv("PASSWORD")
subject = "REPORTE SEMANAL"
body = "Adjunto el reporte semanal"

em = EmailMessage() 
em["from"]= email_sender
em["to"] = email_reciver
em["subject"] = subject
em.set_content(body)

context = ssl.create_default_context()

ruta_report = os.path.join("data", "report", "reporte_prueba.xlsx")
with open (ruta_report, "rb") as f:
    contenido = f.read()
em.add_attachment(
    contenido,
    maintype="application",
    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    filename=os.path.basename(ruta_report),
)

with smtplib.SMTP_SSL("smtp.gmail.com", 465, context = context ) as smtp:
    smtp.login(email_sender, password)
    smtp.send_message(em)
