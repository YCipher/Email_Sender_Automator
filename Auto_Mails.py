#XCipher
import smtplib
from typing import List

import openpyxl as exel
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import getpass
from email.message import EmailMessage
import os

wb = exel.load_workbook(r"C:\Users\Cipher\Desktop\companies.xlsx")
sheet1 = wb.active

username = str(input("Enter your username please : "))
password = str(input("Enter your password please : "))


for i in sheet1['A']:
    msg = EmailMessage()
    mail = i.value
    mail = str(mail)
    if mail.__contains__('@'):
        msg['subject'] = "Demande d'emploi"
        msg['to'] = mail
        msg['from'] = username
        msg.set_content("Bonjour !")
        files = [r"C:\Users\Cipher\Desktop\CV.pdf"]
        for file in files:
            with open(file , 'rb') as cv:
                file_data = cv.read()
                file_name = cv.name
                msg.add_attachment(file_data , maintype = 'image' , subtype = 'octet-stream' ,  filename = file_name)

        server = smtplib.SMTP('smtp.gmail.com' , 587)
        server.starttls()
        server.login(username , password)
        server.send_message(msg)
        print("Message sent to " , str(mail) )
        server.quit()
print("Your message has been sent to all emails in excel list !")











