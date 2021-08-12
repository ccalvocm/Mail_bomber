# -*- coding: utf-8 -*-
"""
Created on Wed Aug 11 14:40:49 2021

@author: ccalvo
"""

import win32com.client
import pandas as pd
# diccionario de recipientes
key_recipients = {'Felipe' : ['Estimado', 'Doctor', 'ccalvo@ciren.cl'], 'María' : ['Estimada', 'Señora', 'ccalvo@ciren.cl']}

# iniciar la aplicacion
outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

# agregar destinatario
# mail.Recipients.Add("mvargas@ciren.cl")
# se pueden agregar personas en copia

# agregar adjunto
# mail.Attachments.Add(r'C:\Users\ccalvo\Pictures\header_2.png')

# asunto del correo
mail.Subject = 'Invitación Taller Participativo: Estudio de erosión Macrozona centro - sur de Chile para la región de O\'Higgins'
# asunto ñuble
mail.Subject = 'Invitación Taller Participativo: Estudio de erosión Macrozona centro - sur de Chile para la región de Ñuble '

# cuerpo ohiggins
mail.HTMLBody = "<html><body> <a href=""https://teams.microsoft.com/registration/aTK-N4lK20SRzUUSwSZIIA,Ao2jz4WK_0y1EcbLWEiQ5A,R4vhq1LzJk2JAKriSPOAqg,_uXfQjtChkqCSR37jQIN9A,kWmJt4SFKUKUtuXEnWnAYw,vV8fuKyy5kyeuN2GvcU-ww?mode=read&tenantId=37be3269-4a89-44db-91cd-4512c1264820""> <img src=""https://raw.githubusercontent.com/ccalvocm/Mail_bomber/main/invitacionOhiggins.jpg"" width=""931.3264346190028"" height=""500""></a></body></html>"

# cuerpo ñuble
mail.HTMLBody = "<html><body> <a href=""https://teams.microsoft.com/registration/aTK-N4lK20SRzUUSwSZIIA,Ao2jz4WK_0y1EcbLWEiQ5A,R4vhq1LzJk2JAKriSPOAqg,1-D3sndFIk-iLBYQ4WP0rw,q68dddF2LUCZftToYPNnsg,UZ0_fdeXMEC9HlsESIUkcA?mode=read&tenantId=37be3269-4a89-44db-91cd-4512c1264820""> <img src=""https://raw.githubusercontent.com/ccalvocm/Mail_bomber/main/invitacionOhiggins.jpg"" width=""931.3264346190028"" height=""500""></a></body></html>"

# destinatarios

excel_mails = pd.read_excel('MAPA DE ACTORES REGION DE OHIGGINS.xlsx', skiprows = 2)

# iterar
for row, col in excel_mails.iterrows():
    print(excel_mails.loc[row, 'Correo electrónico'])
    mail.Recipients.Add(excel_mails.loc[row, 'Correo electrónico'])

# con copia
mail.CC = 'mvargas@ciren.cl'

# enviar
mail.Send()
