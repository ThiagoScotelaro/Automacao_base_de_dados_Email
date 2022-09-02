# Automação de processos com python

#importar blibiotecas

import pyautogui
import time
import pyperclip
from tkinter import E, messagebox
import pandas as pd
import os
import smtplib
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders
pyautogui.PAUSE = 1

#fazendo downloads da base de dados.

pyautogui.press('winleft')
pyautogui.write('chrome')
pyautogui.press('enter')
time.sleep(5)
messagebox.showinfo(title= 'Automation', message=' The automation will start, press ok and wait for the process to finish.')
pyautogui.hotkey('ctrl', 't')

link = ('https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos/vendas-de-derivados-de-petroleo-e-biocombustiveis')

pyperclip.copy(link)
pyautogui.hotkey('Ctrl', 'v')
pyautogui.press('enter')
time.sleep(5)
pyautogui.click(450, 700)
time.sleep(10)
pyautogui.hotkey('alt', 'f4')
time.sleep(10)

# Base de dados

df = pd.read_csv(r'C:\Users\THIAG\Downloads\vendas-derivados-petroleo-etanol-m3-1990-2022.csv', delimiter = ';')

# conversão da base de dados csv para xlsx!

print(r'C:\Users\THIAG\Downloads\vendas-derivados-petroleo-etanol-m3-1990-2022.csv')
df.to_excel(r'C:\Users\THIAG\Downloads\vendas-derivados-petroleo-etanol-m3-1990-2022.xlsx', index= True, header= True)
print("Arquivo convertido!")

df_xlsx = pd.read_excel(r'C:\Users\THIAG\Downloads\vendas-derivados-petroleo-etanol-m3-1990-2022.xlsx')
time.sleep(10)


#envio da informação por email.

sender = "thiago.scotelaro.da.silva@gmail.com"
recipient = "thiago.scotelaro.da.silva@gmail.com"
   
msg = MIMEMultipart() 
msg['From'] = sender 
msg['To'] = recipient 
msg['Subject'] = " Vendas de derivados de petroleo e biocombustiveis"
body = "Segue em anexo a planilha com os dados da venda em formato xlsx. "
msg.attach(MIMEText(body)) 
filename = "vendas-derivados-petroleo-etanol-m3-1990-2022.xlsx"
attachment = open(r'C:\Users\THIAG\Downloads\vendas-derivados-petroleo-etanol-m3-1990-2022.xlsx', "rb") 
p = MIMEBase('application', 'octet-stream') 
p.set_payload((attachment).read()) 
encoders.encode_base64(p) 
   
p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
msg.attach(p) 
s = smtplib.SMTP('smtp.gmail.com', 587) 
s.starttls() 
s.login(sender, "afrgnksnqwrenyfx") 
text = msg.as_string() 
s.sendmail(sender, recipient, text) 
s.quit()
print('email enviado com sucesso')