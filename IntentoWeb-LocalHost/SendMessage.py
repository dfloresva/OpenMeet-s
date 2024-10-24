import pandas as pd
import pyautogui as pg
import time
import webbrowser as web


def sendMessages(num,textSend):
    #ENVIO DE MENSAJES
    url=f"https://web.whatsapp.com/send?phone={num}&text={textSend}"
    web.open(url)
    time.sleep(13)
    pg.leftClick(1711,973)
    
    time.sleep(2)
    pg.hotkey('ctrl', 'w')
    pg.press('enter')
    

#df = pd.read_excel('data.xlsx', sheet_name='ABC', usecols='B:E',skiprows=2) # // nombre del archivo // nombre de la hoja // numero de columnas //salto de fila
df = pd.read_excel('data.xlsx', sheet_name='MSJ-TEST', usecols='A:H',skiprows=1) # // nombre del archivo // numero de columnas //salto de fila
print(df)

   
def create_MSJ(mensaje):
    msj=f"{mensaje}"
    return msj 

for key, row in df.iterrows():
    nombre = row['NOMBRE']
    numero = row['NUMERO']
    mensaje=row['CONCATENAR']
    msj=create_MSJ(mensaje)
    #sendMessages(numero,msj)
    sendMessages(numero,msj)