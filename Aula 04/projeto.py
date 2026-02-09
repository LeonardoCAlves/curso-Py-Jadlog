#!/usr/bin/env python
# coding: utf-8

# <h5 style="color:#2dede4;">
#   🤖 Projeto – Automação de Relatórios Administrativos
# </h5>
# 
# <p style="font-size: 14px">
# Neste exercício, você irá desenvolver um programa em <strong>Python</strong> que simula
# uma <strong>automação básica de processos administrativos</strong>, muito comum em empresas
# que lidam com relatórios operacionais e financeiros.
# </p>
# 
# <p style="font-size: 14px">
# O projeto será desenvolvido de forma <strong>incremental</strong>, iniciando pela
# <strong>organização automática de arquivos</strong> e evoluindo para o
# <strong>envio automatizado de relatórios por e-mail</strong> e
# <strong>notificações via WhatsApp</strong>, simulando um fluxo real de automação corporativa.
# </p>
# 
# <hr>
# 
# <p style="font-size: 18px; color:#2dede4">
#   🎯 Objetivo
# </p>
# <p style="font-size: 14px">Criar um programa que:</p>
# <ul style="font-size: 14px">
#   <li>Verifique a existência de um arquivo Excel de contas processadas</li>
#   <li>Crie automaticamente uma pasta chamada <code>base</code></li>
#   <li>Mova o arquivo para dentro dessa pasta de forma segura</li>
#   <li>Evite erros em execuções repetidas do script</li>
#   <li>Utilize funções para organizar cada etapa da automação</li>
#   <li>Prepare a base para envio automatizado de relatórios</li>
# </ul>
# 
# <hr>
# 
# <p style="font-size: 18px; color:#2dede4">
#   📥 Dados de Entrada
# </p>
# <p style="font-size: 14px">
# O programa deverá trabalhar com um arquivo Excel chamado
# <code>contas_processadas.xlsx</code>, localizado inicialmente no mesmo diretório
# do script ou em uma pasta de dados do projeto.
# </p>
# 
# <p style="font-size: 14px">
# Este arquivo contém informações de contas financeiras já processadas, como:
# </p>
# <ul style="font-size: 14px">
#   <li>ID da conta</li>
#   <li>Tipo (Pagar ou Receber)</li>
#   <li>Descrição</li>
#   <li>Valor original e valor final</li>
#   <li>Datas de vencimento</li>
#   <li>Status da conta</li>
#   <li>Data de processamento</li>
# </ul>
# 
# <hr>
# 
# <p style="font-size: 18px; color:#2dede4">
#   🗂️ Regras de Funcionamento
# </p>
# 
# <table border="1" cellpadding="8" cellspacing="0">
#   <thead style="background-color:#e0f2f1;">
#     <tr style="font-size: 14px; color:#000">
#       <th>Regra</th>
#       <th>Descrição</th>
#     </tr>
#   </thead>
#   <tbody style="font-size: 14px">
#     <tr>
#       <td>Verificação do arquivo</td>
#       <td>O programa deve verificar se o arquivo Excel existe antes de qualquer ação</td>
#     </tr>
#     <tr>
#       <td>Criação de pasta</td>
#       <td>A pasta <code>base</code> deve ser criada automaticamente, se não existir</td>
#     </tr>
#     <tr>
#       <td>Movimentação segura</td>
#       <td>O arquivo deve ser movido para a pasta <code>base</code> sem sobrescrever arquivos</td>
#     </tr>
#     <tr>
#       <td>Execução repetida</td>
#       <td>O script não deve gerar erro se for executado mais de uma vez</td>
#     </tr>
#   </tbody>
# </table>
# 
# <br>
# 
# <p style="font-size: 18px; color:#2dede4">
#   ⚙️ Processamento dos Dados
# </p>
# 
# <table border="1" cellpadding="8" cellspacing="0">
#   <thead style="background-color:#e0f2f1;">
#     <tr style="font-size: 14px; color:#000">
#       <th>Etapa</th>
#       <th>Descrição</th>
#     </tr>
#   </thead>
#   <tbody style="font-size: 14px">
#     <tr>
#       <td>Validação</td>
#       <td>Checar a existência do arquivo de origem</td>
#     </tr>
#     <tr>
#       <td>Organização</td>
#       <td>Criar a estrutura de pastas necessária para o projeto</td>
#     </tr>
#     <tr>
#       <td>Automação</td>
#       <td>Mover o arquivo utilizando funções do sistema operacional</td>
#     </tr>
#     <tr>
#       <td>Feedback</td>
#       <td>Exibir mensagens claras informando o status da automação</td>
#     </tr>
#   </tbody>
# </table>
# 
# <hr>
# 
# <p style="font-size: 18px; color:#2dede4">
#   📤 Saída Esperada
# </p>
# <p style="font-size: 14px">Ao final da execução, o programa deve:</p>
# <ul style="font-size: 14px">
#   <li>Criar a pasta <code>base</code>, caso ela não exista</li>
#   <li>Mover o arquivo <code>contas_processadas.xlsx</code> para essa pasta</li>
#   <li>Informar se o arquivo já foi movido anteriormente</li>
#   <li>Exibir mensagens de sucesso ou alerta durante o processo</li>
# </ul>
# 
# <p style="background:#f1f8e9; padding:2px; border-left:6px solid #000000; color:#000051; font-size:14px">
# 💡 <strong>Dica:</strong> utilize os módulos <strong>os</strong> e <strong>shutil</strong>
# para manipular arquivos e diretórios, organize o código em
# <strong>funções</strong> e pense neste projeto como a base de uma
# <strong>automação corporativa completa</strong>.
# </p>
# 
# <p style="text-align:center; color:#00ff37;">
# 🚀 Bora automatizar processos de verdade! 👽
# </p>
# 
# 
# 

# Importando bibliotecas

# In[1]:


import os
import smtplib
import shutil
import pyautogui
import pywhatkit as kit
from time import sleep
from datetime import datetime
from email.message import EmailMessage
import win32com.client as win32


# Organizando arquivos

# In[2]:


arquivo_excel = 'contas_processadas.xlsx'
pasta_base = 'base'


# Verificando se o arquivo existe

# In[3]:


if not os.path.exists(arquivo_excel):
    raise FileNotFoundError("Arquivo não encontrado, verifique se o download está concluído")

print("Arquivo encontrado")


# Criando a pasta base dentro da pasta Aula 04

# In[4]:


if not os.path.exists(pasta_base):
    os.makedirs(pasta_base)
    print('Pasta criada com sucesso.')

else:
    print('Pasta Base já foi existe.')


# Movendo arquivo para dentro da pasta Base

# In[5]:


destino = os.path.join(pasta_base, os.path.basename(arquivo_excel))

if not os.path.exists(destino):
    shutil.move(arquivo_excel, destino)
    print('Arquivo movido com sucesso!')

else:
    print('Arquivo já existe na pasta Base')


# Criando corpo do Email

# In[6]:


hoje = datetime.now().strftime("%d . %m . %Y | %H:%M")

corpo_email = f"""
Prezados,

Segue em anexo o relatorio de contas processadas do mês de janeiro.
Data de processamento: {hoje}

Atenciosamente

👽may the 4th b with u
"""       


# Configurando OUTLOOK

# In[7]:


outlook = win32.Dispatch('Outlook.Application')

email = outlook.CreateItem(0)

email.To = 'leon4rdoalvess@gmail.com'
email.Subject = 'Relatório de contas processadas de janeiro'
email.Body = corpo_email

email.attachments.Add(
    r"C:\Users\leon4\Documents\Turmas\Python\03_JadLog T1\_Alunos\Aula 04\base\contas_processadas.xlsx"
)

email.Send()

print('Email enviado com sucesso!')


# Alerta WhatsApp

# In[8]:


telefone = '+5511979714423'

mensagem = "Relatório enviado lá jão!!!"

kit.sendwhatmsg_instantly(
    telefone,
    mensagem,
    wait_time=15,
    tab_close=False
)

sleep(2)

pyautogui.press("enter")
print('Mensagem enviada com sucesso!')


# Criando executável

# In[ ]:


# !pip install pyinstaller 

get_ipython().system('jupyter nbconvert --to script projeto.ipynb')

