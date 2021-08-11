# -*- coding: utf-8 -*-
"""
Created on Wed Aug 11 14:40:49 2021

@author: ccalvo
"""

SERVER = "smtp.outlook.com"
FROM = "ccalvo@ciren.cl"
TO = ["farrospide@ciren.cl"] # must be a list

SUBJECT = "Hello!"
TEXT = "This is a test of emailing through smtp of example.com."

# Prepare actual message
message = """From: %s\r\nTo: %s\r\nSubject: %s\r\n\

%s
""" % (FROM, ", ".join(TO), SUBJECT, TEXT)

# # Send the mail
# import smtplib
# server = smtplib.SMTP(SERVER)
# server.login("ccalvo", "Ciren2021")
# server.sendmail(FROM, TO, message)
# server.quit()

import win32com.client
key_recipients = {'Felipe' : ['Doctor', 'Estimado'], 'María' : ['Señora', 'Estimada']}
outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'ccalvo@ciren.cl'
mail.Recipients.Add("mvargas@ciren.cl")

mail.Subject = 'Email test'
mail.HTMLBody = '<h3>This is HTML Body</h3>'
for key in key_recipients.keys():
    mail.Body = key_recipients[key][1]+' '+ key_recipients[key][0] +' '+ key
    print(mail.Body)
    # mail.CC = 'mvargas@ciren.cl'
    mail.Send()