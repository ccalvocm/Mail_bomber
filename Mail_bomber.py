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

mail.HTMLBody = "<html><body> <a href=""https://teams.microsoft.com/registration/aTK-N4lK20SRzUUSwSZIIA,Ao2jz4WK_0y1EcbLWEiQ5A,R4vhq1LzJk2JAKriSPOAqg,_uXfQjtChkqCSR37jQIN9A,kWmJt4SFKUKUtuXEnWnAYw,vV8fuKyy5kyeuN2GvcU-ww?mode=read&tenantId=37be3269-4a89-44db-91cd-4512c1264820""> <img src=""https://southcentralus1-mediap.svc.ms/transform/thumbnail?provider=spo&inputFormat=jpg&cs=fFNQTw&docid=https%3A%2F%2Fcirencl-my.sharepoint.com%3A443%2F_api%2Fv2.0%2Fdrives%2Fb!2FmpUux3nUeDqASyP8Dent5gMFQiUKxCivwO3LmDskdC7mrrdFGAQ44vb8TMFOm6%2Fitems%2F01NZ4S5E6H4H5PV5PKTJFY6CL4YMNVLKHV%3Fversion%3DPublished&access_token=eyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvY2lyZW5jbC1teS5zaGFyZXBvaW50LmNvbUAzN2JlMzI2OS00YTg5LTQ0ZGItOTFjZC00NTEyYzEyNjQ4MjAiLCJpc3MiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAiLCJuYmYiOiIxNjI4NzkxMjAwIiwiZXhwIjoiMTYyODgxMjgwMCIsImVuZHBvaW50dXJsIjoiaVlsTGJsd2hQOFNnbjNFUTl3TDg0YzVtbjJSN1hBT0JId2FtVXY2clUzQT0iLCJlbmRwb2ludHVybExlbmd0aCI6IjExNyIsImlzbG9vcGJhY2siOiJUcnVlIiwidmVyIjoiaGFzaGVkcHJvb2Z0b2tlbiIsInNpdGVpZCI6Ik5USmhPVFU1WkRndE56ZGxZeTAwTnpsa0xUZ3pZVGd0TURSaU1qTm1ZekJrWlRsbCIsIm5hbWVpZCI6IjAjLmZ8bWVtYmVyc2hpcHx1cm4lM2FzcG8lM2Fhbm9uI2VkOGVlYjUwYWRkZmU5OTA4YzI2YTEyMmJkODJjNmI2MjExNjMzOTFkNTBjMDRkZjRkNmI2ZGI5OTk1NzdlN2YiLCJuaWkiOiJtaWNyb3NvZnQuc2hhcmVwb2ludCIsImlzdXNlciI6InRydWUiLCJjYWNoZWtleSI6IjBoLmZ8bWVtYmVyc2hpcHx1cm4lM2FzcG8lM2Fhbm9uI2VkOGVlYjUwYWRkZmU5OTA4YzI2YTEyMmJkODJjNmI2MjExNjMzOTFkNTBjMDRkZjRkNmI2ZGI5OTk1NzdlN2YiLCJzaGFyaW5naWQiOiJUbURjbUZsUmVFV0ZUdGt6bzV4cmpnIiwidHQiOiIwIiwidXNlUGVyc2lzdGVudENvb2tpZSI6IjIifQ.MTZmMkVKVGRXL3YxZU9xeHZGcXBBMGtoL01DMHdDM2ZBNGVBczYxZ3A0OD0&cTag=%22c%3A%7BFAFAE1C7-EAF5-4B9A-8F09-7CC31B55A8F5%7D%2C1%22&encodeFailures=1&width=1920&height=893&srcWidth=&srcHeight="" width=""931.3264346190028"" height=""500""></a></body></html>"

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
