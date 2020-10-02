# <b>Email de Aniversário Personalizado<b> #

# Requisitos

  - Python 3.7.x 

  - Libs: urllib3, pywin32, openpyxl

  - Outlook 2016

# Como usar?
  Antes de utilizar a função é necessário inserir na mesma o email do destinatário. Por padrão, o envio vai ser a partir do email registrado no outlook. Mas caso queria enviar em um email diferentes, deverá ser colocado na função também.

  A função enviar_email() recebe 4 paramêtros, sendo 2 deles vetores: nomes,setor e os outros dois numeros: dia, mês.

  Exemplo:

  nomes = ['Gabriel','Douglas'] <br>
  setor = ['DEPC','DDPM'] <br>
  dia = 25 <br>
  mes = 10 <br>

  enviar_email(nomes,setor,dia,mes)

  Após essa chamada, será enviado o email.

  OBS: É importante que o Outlook esteja aberto para o envio funcionar normalmente.

# Template do Email
<b>Imagem retirada do site https://unsplash.com/</b>

<img src="Template_Email.PNG" alt="template_email_birthday">
