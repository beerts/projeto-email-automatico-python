#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# Para controle de custos, todos os dias, seu chefe pede um relatório com todas as compras de mercadorias da empresa.
# O seu trabalho, como analista, é enviar um e-mail para ele, assim que começar a trabalhar, com o total gasto, a quantidade de produtos compradas e o preço médio dos produtos.
# 
# E-mail do seu chefe: para o nosso exercício, coloque um e-mail seu como sendo o e-mail do seu chefe<br>
# Link de acesso ao sistema da empresa: https://pages.hashtagtreinamentos.com/aula1-intensivao-sistema
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[42]:


import pyautogui as py
import pyperclip as pc
import pandas as pd
import time



# pyautogui.click -> clicar com o mouse
# pyautogui.write -> escrever um texto
# pyautogui.pres -> apertar uma tecla
# pyautogui.hotkey -> apertar uma combinação de teclas
# pyautigui.dragto -> clicar e arrastar  


py.PAUSE = 3


#1- ACESSAR SISTEMA DA EMPRESA

py.hotkey("ctrl", "t")
py.write("https://pages.hashtagtreinamentos.com/aula1-intensivao-sistema")
py.press("enter")

time.sleep(3)

#2- FAZER LOGIN NO SISTEMA

py.click(x=882, y=467,  clicks= 1, button ="left") #clica para digitar o login
py.write('bernardort32@outlook.com') #login
py.click(x=894, y=550,  clicks= 1, button ="left") #clica na senha
py.write('senha123321') #senha
py.click(x=894, y=550, clicks= 1, button ="left") #clica no fazer login
py.click(x=971, y=657, clicks= 1 ,button ="left") 

#3- BAIXAR A BASE DE DADOS

py.click(x=709, y=442, clicks= 1, button ="left")
py.click(x=831, y=852, clicks = 1, button ="left")

#4- CALCULAR INDICADORES

#importar base de dados

tabela = pd.read_csv(r"C:\Users\berna\Downloads\Compras.csv", sep = ";")
# display(tabela)

#CALCULO INDICADORES
#total gasto -> somar valor final

total = tabela["ValorFinal"].sum()

#quantidade -> somar quantidade

quantidade = tabela["Quantidade"].sum()

#preço medio ->total gasto / quantidade

media = total / quantidade
      

#5- ENVIAR O EMAIL PARA A DIRETORIA/PARA O CHEFE

#entrar no email (OUTLOOK)

py.PAUSE = 3
py.hotkey("ctrl", "t")
py.write("https://outlook.live.com/mail/0/")
py.press("enter")  #entra no email

py.click(x=261, y=257) #clica em novo email
py.write("pythonimpressionador@gmail.com") #digita o email
py.press("tab") #confirma destinatário
py.press("tab") #passa para o assunto


pc.copy("Relatório de compras") #assunto
py.hotkey("ctrl", "v") #copia o assunto

py.press("tab") #passa para o texto

#TEXTO
texto = f"""

Prezados,

Segue aqui o relatório de compras do mês de dezembro:


____________________________

TOTAL = R${total:,.2f}
________________________________________

QUANTIDADE = {quantidade:.0f} UNIDADES
________________________________________

MÉDIA = R${media:,.2f}
_____________________________



Att.,
Bernardo Teixeira

"""

pc.copy(texto)  #copia o texto
py.hotkey("ctrl", "v") #cola o texto na aba texto

#anexar arquivo 
py.click(x=1598, y=250) 
py.click(x=1450, y=307)
py.click(x=1157, y=432)

#enviar
py.click(x=466, y=328) #clica em enviar


# In[39]:


import time
print(py.position(time.sleep(5)))


# In[ ]:




