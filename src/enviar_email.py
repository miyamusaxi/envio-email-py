import smtplib
import re
import pathlib as pb

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def tratar_lista_emails(lista, remover):
    items = []
    for i in lista:
        item = i
        string = str(item).strip(remover)
        string = string.replace("'","")
        items.append(string)    
    return items

def enviar_email(assunto, destinatario,corpo_email, remetente, psw, arq = [], host = 'smtp.gmail.com', port = '587'):
        
    server = smtplib.SMTP(host= host, port= port)
    server.ehlo()
    server.starttls()
    
    server.login(remetente, psw)
    
    msg = MIMEMultipart()
    
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto
    
    msg.attach(MIMEText(corpo_email, 'html'))
    
    for i in arq:
        ident, caminho, file = i
        ident = str(ident).lower()
        if caminho and file: 
            #Nesse if ele verifica se o identificador bate com o destinatario pra anexar o arquivo certo
            if ident in str(destinatario):   
            #Abrindo o arquivo pelo path  
                attchment = open(caminho, 'rb')
    
                #Lendo o arquvio e transformando ele em base64
                att = MIMEBase('application', 'octet-stream')
                att.set_payload(attchment.read())
                encoders.encode_base64(att)
    
                #adicionando o nome do arquivo
                att.add_header('Content-Disposition',f'attachment; filename= {file}')
                attchment.close()
    
                #anexando o arquivo
                msg.attach(att)
            else:
                pass
        
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()
    
def verificar_email(email):
    r = re.compile(r'^[\w-]+@(?:[a-zA-Z0-9-]+\.)+[a-zA-Z]{2,}$')
    email = r.match(email)
    return email


if __name__ == '__main__':
    # lista de destinatários
    destinatarios = ['teste.suporte.prox@gmail.com','277353.git@gmail.com']
    #Lista com o identificador dos arquivos e o caminho dele no sistema
    cam_arq = []
    # Para anexar documentos
    nome_arquivo = 'teste1.pdf' # define um nome do arquivo tem que ter o tipo pra ser definido também na hora do envio do e-mail
    caminho = pb.Path(f'..\Docs\{nome_arquivo}') #le o caminho
    ident = 'prox' #identifica pra qual  email você quer enviar, pode ser o nome ou o servidor, você quem define esse filtro
    tupla_arq = (ident, caminho, nome_arquivo) # cria uma tupla pra depois anexar 
    cam_arq.append(tupla_arq) #adiciona a 

    nome_arquivo = 'teste2.pdf'
    caminho = pb.Path(f'..\Docs\{nome_arquivo}')
    ident = 'git'
    tupla_arq = (ident, caminho, nome_arquivo)
    cam_arq.append(tupla_arq)
    
    remet = '' #Quem envia o arquivo
    psw = '' # Senha de app que você pode criar no gmail e usar pra enviar o e-mail pelo python
    assunto = 'TESTE ENVIO MAIS DE UM ARQUIVO' # Assunto do e-mail
    #Corpo do E-mail
    corpo_email = 'Olá, este email esta sendo enviado de forma automática e com <b>HTML</b> no corpo do email<p><b>SACOU?</b></p>' 
    
    if cam_arq:
        for i in destinatarios:
            destinatario = i
            if verificar_email(destinatario):    
                enviar_email(assunto, destinatario, corpo_email, psw =psw, remetente= remet, arq= cam_arq)
    else:
        for i in destinatarios:
            destinatario = i
            if verificar_email(destinatario):    
                enviar_email(assunto, destinatario, corpo_email, psw =psw, remetente= remet)

