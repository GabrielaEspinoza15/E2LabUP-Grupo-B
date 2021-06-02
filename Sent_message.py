#Google creado con Google.py
from Google import Create_Service
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import docx2txt
import schedule
import time


CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://mail.google.com/']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

running = True
def test():


    emailMsg = docx2txt.process(r'''C:\Users\Usuario\Documents\GMAIL_API\Archivo\Mensaje.docx''')

    mimeMessage = MIMEMultipart()
    mimeMessage['to'] = '<gabrielamylen15@gmail.com>, <gabriela.espinoza@aiesec.net>'
    mimeMessage['subject'] = 'Mensaje enviado desde Python'
    mimeMessage.attach(MIMEText(emailMsg, 'plain'))
    raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()

    message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
    
    global running
    running = False

    print(message)


    return schedule.CancelJob
#Hora programada
schedule.every().day.at("20:19").do(test)
while running:
    schedule.run_pending()
    time.sleep(1)
