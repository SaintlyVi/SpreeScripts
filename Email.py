# -*- coding: utf-8 -*-
"""
Created on Wed Oct 29 15:15:51 2014

@author: Wiebke.Toussaint
"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import date
from email import Encoders

def send_message( doc_name, message, part, maillist ): #send email function

    msg = MIMEMultipart()
            
    today = date.today()
    urlFile = open(maillist, "r+") 
    maillist = [i.strip() for i in urlFile.readlines()]    
        
    fromEmail = 'spreewarehouse@gmail.com' 
    #create message
    msg['Subject'] = str(doc_name) + str(today)
    msg['From'] = fromEmail
    #msg['To'] = ', '.join(MailList)
    body = message
    content = MIMEText(body, 'plain')
    msg.attach(content)
        
    #create attachment        
    filename = str(part)
    f = file(filename)
    attachment = MIMEText(f.read())
    attachment.set_payload(open(part, 'rb').read())
    Encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', 'attachment', filename=filename)           
    msg.attach(attachment)
        
    #call server and send email      
    mailServer = smtplib.SMTP('smtp.gmail.com', 587)
    mailServer.set_debuglevel(1)
    mailServer.ehlo()
    mailServer.starttls()
    mailServer.ehlo()
    mailServer.login('spreewarehouse@gmail.com', 'spreeapp')
    mailServer.ehlo()
    mailServer.sendmail(fromEmail, maillist, msg.as_string())
    mailServer.quit()
        
    print "Mail sent successfully"
    return 