import PyPDF2
import os
import csv
from string import Template
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import time  
import config

def readTemplate(filename):
    with open(filename, encoding='utf-8', errors='ignore') as templateFile:
        templateFileContent = templateFile.read()
    return Template(templateFileContent)

def readPDFAndCreateEmail():
    emailAddress = config.address
    password = config.password
    directory = config.localDirectory
    emailFormater = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
    emailFormater.starttls()
    emailFormater.login(emailAddress, password)
    data = []

    for file in os.listdir(directory):
        if not file.endswith(".pdf"):
            continue
        with open(os.path.join(directory,file), 'rb') as pdfFileObj:  
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            pdfReader.numPages
            pageObj = pdfReader.getPage(0)
            pageString = pageObj.extractText()
            splitString = pageString.split('\n')
            try:
                fileName = file
            except:
                fileName = 'N/A'
            try:
                indexNet30 = splitString.index("Net 30")
                name =  splitString[indexNet30-2]
            except:
                name = 'N/A'
            try:
                indexMoney =  splitString.index('Embroidery') 
                money = splitString[indexMoney+2]
                if money == "Waved":
                    money = splitString[indexMoney+3] 
            except: 
                money = 'N/A'
            try:
                indexEmail =  splitString.index('E-mail:') 
                email = splitString[indexEmail+1]
                print(email)  
            except:
                email = 'N/A'
            data.append((fileName, name, email, money))
            
            message_template = readTemplate('invoiceEmail.txt') 
            msg = MIMEMultipart()    
            message = message_template.substitute( PERSON_NAME = name, MONEY_AMOUNT = money )   
            print(message)

            #formats the message sent
            msg['From']=emailAddress
            msg['To']=email
            msg['Subject']=fileName
            msg.attach(MIMEText(message, 'plain'))
            
            #Code below adds the attachment 
            fileAttached = file
            attachment = open(config.localDirectory + file, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', "attachment; filename= %s" % fileAttached)
            msg.attach(part)
            emailFormater.send_message(msg)
            del msg
        
def writeToCSV(data):        
    with open('Invoices.csv', 'a') as csv_file:
        writer = csv.writer(csv_file, lineterminator='\n')
        writer.writerow(['Invoice Name','Name', 'Email Address', 'Amount Owed'])
        for fileName, name, email, money in data:
            writer.writerow([fileName, name, email, money])

def main():
    data = readPDFAndCreateEmail()
    writeToCSV(data)

if '__name__' == '__main__':
    main()
