"""
Created on Thu Feb 17 07:38:05 2022
Desarrollo envio  calculo AENC actuualizado
para cada uno de los sectores
@author: Sebastian
"""

import pandas as pd
import pyodbc
from datetime import datetime
import win32com.client
from conexion_bd import con_data, consulta_sql

# ============================================================
# VARIABLES CALCULO
# ============================================================

mes = 11
anio = 2022

dia = str(datetime.date(datetime.today()))
nombre_archivo = f'CALCULO_AENC_{dia}.xlsx'
# ============================================================
# CONEXION BD
# ============================================================
conexion = con_data()
# ============================================================
# LISTADO DE CLIENTES A ENVIAR
# ============================================================
listado = 'rposada@celsia.com;hbriceno@celsia.com;mmoya@celsia.com;dgomezg@celsia.com;'
# ============================================================
# SQL A CONSULTAR
# ============================================================
sql = consulta_sql(mes=mes, anio=anio)

data = pd.read_sql_query(sql, conexion)
data.to_excel(
    f'C:\\Users\\Sebastian\\Documents\\Python Scripts\\AENC\\{nombre_archivo}', index=False)
adjunto = f'C:\\Users\\Sebastian\\Documents\\Python Scripts\\AENC\\{nombre_archivo}'

def envios_aenc(to, copia, adjunto, dia):
    outlook = win32com.client.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    html_mesagge = open('home.html').read()
    email.Attachments.Add(adjunto)
    email.To = to
    email.CC = copia
    #email.Subject = 'CALCULO AENC VERSION TXF OCTUBRE 2022'
    email.Subject = f'CALCULO AENC VERSION TX2 - NOVIEMBRE {dia}'
    email.HTMLBody = html_mesagge
    email.Display(False)
    email.Send()
    print('ok')


envios_aenc('sfherrera@celsia.com', listado, adjunto, dia)
