# coding=utf-8:

import sys,os,datetime
from urllib3 import request
from win32com.client import Dispatch
from openpyxl import load_workbook

def enviar_email(nomes,setor,dia,mes):
  outlook = Dispatch('outlook.application') #Obtendo uma instancia do outlook
  mail = outlook.CreateItem(0) #Criando objeto de amil
  mail.To = 'gabriel.gomes@edpbr.com.br'#"o365_edpbr.com.br@edponcloud.onmicrosoft.com" #Email do destinatário
  mail.sentOnBehalfOfName = "projeto-construcao@edpbr.com.br"
  mail.CC = ""#Copia para alguem
  mail.Subject = f"Feliz Aniversário! Aniversariantes de {dia}/{mes}" #Assunto do E-mail

  # mail.Body = "Corpo do email"
  # mail.Body += '''
  #     Eu nao sei
  # '''

  colaboradores_area = ''
  for nome,area in zip(nomes,setor):
    colaboradores_area += "<br>"+nome+" - "+area+"<br> "

  mail.HTMLbody += '''
  <!DOCTYPE html>
  <html>
  <head>
  <meta charset='utf-8'>
  <meta http-equiv='X-UA-Compatible' content='IE=edge'>
  <meta name='viewport' content='width=device-width, initial-scale=1'>

  <style>
    body{
      font-family: 'EDP Preon';
      color: white;
      background-color:#151c22;
    }

    .main{
      border: 5px;
      border-collapse: collapse;
      border-color: white;
    }
  </style>
</head>
<body style="margin: 0; padding: 0;">
  <table cellpadding="0" cellspacing="0" width="100%">
    <table align="center" cellpadding="0" cellspacing="0" width="600">
      <tr>
        <td bgcolor="#151c22">
          <img src="https://images.unsplash.com/photo-1464349153735-7db50ed83c84?ixlib=rb-1.2.1&ixid=eyJhcHBfaWQiOjEyMDd9&auto=format&fit=crop&w=831&q=80" alt="Criando Mágica de E-mail" width="600" height="300" style="display: block;" />
        </td>
      </tr>
      <tr>
        <td bgcolor="#151c22" style="padding: 25px 30px 40px 30px;">
          <table cellpadding="0" cellspacing="0" width="100%">
           <tr>
            <td style="text-align: center;font-size: 30px;">
             <b>FELIZ ANIVERSÁRIO</b>
            </td>
           </tr>
           <tr>
            <td style="padding: 20px 0 30px 0; text-align: center;">'''

  mail.HTMLbody += f'''A equipe DEP deseja antecipadamente um feliz aniversário ao(s) colaborador(es):<br><b>{colaboradores_area}</b><br>
             É um prazer ter você na nossa equipe, e esperamos que compartilhe muitos
             anos de vida com a gente!!!'''#.format(colaboradores_area)
             
  mail.HTMLbody +='''
            </td>
           </tr>
           <tr>
            <td style="text-align: center;">
             <b>EQUIPE DEP - Projetos e Construções</b>
            </td>
           </tr>
          </table>
         </td>
      </tr>
      <!-- <tr>
        <td bgcolor="#ee4c50">
          Linha 3
        </td>
      </tr> -->
     </table>
  </table>
</body>
</html>
  '''
  # mail.Attachments.Add("Caminho do seu arquivo!")
  # mail.Attachments.Add("C:\\Users\\7100746\\Desktop\\irineu.txt")
  mail.Send() #Envia o email

# enviar_email()