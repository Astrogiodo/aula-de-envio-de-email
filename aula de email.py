# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[119]:

#sexo a 3


#instalar uma única vez
get_ipython().system('pip install pyautogui')
get_ipython().system('pip install pyperclip')


# In[120]:


import pyautogui
import pyperclip
import time


pyautogui.PAUSE=1

# escrever o passo a passo em português
# passo 1 - entrar no sistema da empresa (no caso o drive)
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")

time.sleep(3)

# passo 2 - navegar no sistema até encontrar a base de dados
time.sleep(5)
pyautogui.click(x=422, y=296, clicks=2)
time.sleep(2)

#passo 3 - exportar a base de vendas
pyautogui.click(x=429, y=394, clicks=1)
pyautogui.click(x=1714, y=187, clicks=1)
pyautogui.click(x=1499, y=595, clicks=1)
time.sleep(5)

# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[121]:


#passo 4 - calcularia os indicadores (faturamento e quantidade de produtos vendidos)
import pandas as pd

tabela = pd.read_excel(r"C:\Users\Yuri\Downloads\Vendas - Dez.xlsx")

display(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
time.sleep(3)


# In[130]:


#passo 5 - enviar em email para a diretoria com os indicadores
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=wm#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(3)

pyautogui.click(x=78, y=207, clicks=1)
time.sleep(5)

pyautogui.write("yuri.dias.de.souza.123@gmail.com")
pyautogui.press("tab")

pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl","v")
time.sleep(1)

pyautogui.press("tab")
texto = f"""
Prezados clientes, bom dia
o faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de:{quantidade:,}

abs
Yuri"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

pyautogui.hotkey("ctrl","enter")


# In[ ]:




#### Use esse código para descobrir qual a posição de um item que queira clicar
# pyautogui.position()
- Lembre-se: a posição na sua tela é diferente da posição na minha tela





