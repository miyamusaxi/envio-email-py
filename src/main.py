import smtplib
import re
import pandas as pd
import numpy as np

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

def verificar_email(email):
    r = re.compile(r'^[\w-]+@(?:[a-zA-Z0-9-]+\.)+[a-zA-Z]{2,}$')
    email = r.match(email)
    return True

def enviar_email(assunto, destinatario,corpo_email, remetente, psw, cc = None, arq = [], host = 'smtp.gmail.com', port = '587'):
        
    server = smtplib.SMTP(host= host, port= port)
    server.ehlo()
    server.starttls()
    
    server.login(remetente, psw)
    
    msg = MIMEMultipart()
    
    # Para o email ficar mais bonito e organizado você consegue definir um nome de usuario
    msg['From'] = f'Usuario da Silva <{remetente}>'
    msg['To'] = destinatario
    msg['Subject'] = assunto
    if cc:
        msg['Cc'] = cc
    else: 
        msg['Cc']= None
    
    msg.attach(MIMEText(corpo_email, 'html'))

    #Anexo de arquivos em pdf
    for data in arq:
        files = [data['doc'],data['doc2']]
        for f in files:
            if f: 
            #Nesse if ele verifica se o identificador bate com o destinatario pra anexar o arquivo certo
                if str(data['ident']).lower() in str(destinatario):   
            #Abrindo o arquivo pelo path  
                    attchment = open(r'C:\py_projects\automatizacao\envio-email-py\doc\{}.pdf'.format(f), 'rb')
    
                #Lendo o arquvio e transformando ele em base64
                    att = MIMEBase('application', 'octet-stream')
                    att.set_payload(attchment.read())
                    encoders.encode_base64(att)
    
                #adicionando o nome do arquivo
                    att.add_header('Content-Disposition',f'attachment; filename= {f}.pdf')
                    attchment.close()
                #anexando o arquivo
                    msg.attach(att)
            else:
                pass 
    if msg['Cc']:
        to = [msg['To']]+ msg['Cc'].split(';')
        #to = [verificar_email(email) for email in to_emails]
        server.sendmail(msg['From'], to, msg.as_string())
    else:
        server.sendmail(msg['From'], msg['To'], msg.as_string())
    print('Success to send mail!')
    server.quit()

if __name__ == '__main__':
    arq = []
    #email que sera enviado o email
    remetente = 'mail@mail.com'
    #senha de app
    psw = ''
    #nome da planilha com os contatos
    nomePlanilha = 'modelo'
    assunto = 'Assunto escolhido | Assunto X'
    itens = pd.read_excel(
        f'path\{nomePlanilha}.xlsx', sheet_name='Enviar').replace({np.nan: None}).to_dict(orient='records')
    
    for row in itens:
        arquivos = {'ident':row['ident'],'doc':row['doc'],'doc2':row['doc2']}
        arq.append(arquivos)
        corpo_email = f"""
            Corpo em HTML ou apenas uma mensagem padrão
        """
        # Verifica se esta tudo preenchido para enviar os emails um com pessoas em cópia e anexos
        
        if verificar_email(row['email']) and arq and row['cc']:
            enviar_email(remetente= remetente, psw= psw, destinatario= row['email'], arq= arq, assunto= assunto, corpo_email= corpo_email, cc=row['cc'])
        
        elif verificar_email(row['email']) and arq:
            enviar_email(remetente= remetente,arq=arq, psw= psw, destinatario= row['email'], assunto= assunto, corpo_email= corpo_email)
        
        elif verificar_email(row['email']):
            enviar_email(remetente= remetente, psw= psw, destinatario= row['email'], assunto= assunto, corpo_email= corpo_email)