"""
Código produzido por Felipe Whitaker em 14/04/2018 com ajuda do StackOverFlow
"""

from EMAIL_FUNCOES import pip_install

_packages_ = ('pip', 'pandas', 'xlrd', 'openpyxl')
pip_install(_packages_)

import pandas
import time, smtplib, xlrd, openpyxl
from os import getcwd
from time import sleep, strftime, localtime
from email.mime.multipart import MIMEMultipart
from EMAIL_FUNCOES import info_linha, cria_corpo, text_html, format_subst_txt,\
     cria_txt, cria_xlsx

cria_txt()
print('Iniciando o programa...\nO modelo para pegar informações foi criado. ' + \
      'Atualize-o e coloque na pasta principal ')

email, assunto, txt, xlsx = info_linha(input('.txt com as informações: ') + '.txt')

base_text = open(txt + '.txt', 'r').read()
nome_pdf  = (base_text.split('pdf_')[1] + '.pdf'
             if 'pdf_' in base_text else False)

cria_xlsx(base_text, xlsx)
print('Modelo de Planilha criado! Atualize-o e coloque na pasta principal ')

dfPlanilha  = pandas.read_excel(xlsx + '.xlsx')
dfIter      = dfPlanilha.loc[dfPlanilha['bool'] == False]

qtd     = len(dfIter.index)
perc    = 100 / qtd

print('Total de %d destinatários. '%qtd)

server  = smtplib.SMTP('smtp.gmail.com', 587)
senha   = input('Senha de %s: '%email)
server.starttls()
server.login(email, senha)
print('Login em %s feito com sucesso! '%email)

print('Iniciando iteração... ')
for i in range(qtd):

    para_email = dfIter.iloc[i].loc['email']

    msg_email  = cria_corpo(msg         = MIMEMultipart(),
                            From        = email,
                            Subject     = assunto,
                            To          = para_email,
                            Text        = base_text,
                            df          = dfIter,
                            idx         = i,
                            anexo_pdf   = nome_pdf)
    
    server.sendmail(email, para_email, msg_email.as_string())

    if i == 0 and not bool(input('Está correto? ')):
        print('Iteração interrompida ')
        break

    print('%.2f/100%% feito! '%(perc * (i + 1)), end = '\r')
    
    sleep(2)

server.quit()

dfPlanilha['bool']                                      = True
dfPlanilha.loc[dfPlanilha['empresa'] == 'SIEng','bool'] = False
dma = strftime('%d-%m-%Y', localtime())

dfPlanilha.to_excel(excel_writer    = getcwd() + r'\Antigos\%s'%xlsx + '.xlsx',
                    sheet_name      = 'Atualizado %s'%dma,
                    index           = False)
