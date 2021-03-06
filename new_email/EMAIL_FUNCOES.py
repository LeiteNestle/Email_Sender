"""
Código produzido por Felipe Whitaker em 14/04/2018 com ajuda do StackOverFlow
"""

import sys
import pandas as pd
from os import getcwd
from email import encoders
from time import strftime, localtime
from email.mime.text import MIMEText
from email.mime.base import MIMEBase

""" Instalação dos pacotes necessários """

def pip_install(packages):
    print('Importando os pacotes necessários, por favor aguarde ')
    from importlib import import_module as imp_mod
    for pack in packages:
        try:
            imp_mod(pack)
        except ImportError:
            print('Instalando %s. Pode demorar... '%pack)
            from pip import main
            main(['install','--upgrade', pack])
        finally:
            imp_mod(pack)
            print('%s importado com sucesso '%pack)
    return

"""Retiran informações do arquivo de texto """

def info_linha(origem):
    """Divide o .txt em [email, assunto, txt, xlsx] """
    return open(origem, 'r').read().split('\n')

"""Criação do objeto """

def cria_corpo(msg, From, Subject, To, Text, df, idx, anexo_pdf):
    msg['From']     = From
    msg['Subject']  = Subject
    msg['To']       = To

    new_text = text_html(format_subst_txt(Text, df, idx))
    msg.attach(MIMEText(new_text, 'html'))
    
    if anexo_pdf:
        anexo       = open(anexo_pdf, 'rb').read()
        attach      = MIMEBase('application', 'octet-stream').set_payload(anexo)
        encoders.encode_base64(attach)
        attach.add_header('content-disposition',
                          'attachment; filename = %s'%anexo_pdf)
        msg.attach(attach)

    return msg
    

"""Alterações no texto """

def text_html(text):
    return '<br>'.join(text.split('\n'))

def format_subst_txt(text, df, idx, dic =   {'#':   ('<b>','</b>'),
                                             '$':   ('<i>','</i>'),
                                             '%':   ('<u>','</u>')}):
    text_list = text.split('_')
    for i,n in enumerate(text_list):
        if '@' in n:
            text_list[i] = df.iloc[idx].loc[n.split('@')[0]] + n.split('@')[1]
        for tipo in dic:
            if tipo in n:
                text_list[i] = dic[tipo][0] + (dic[tipo][1] + ' ').join(n.split(tipo))
    return ''.join(text_list)

"""Criação dos arquivos base """

def cria_txt(nome = 'origem.txt', interior = 'de_email\nassunto\ntexto_base\nplanilha'):
    open(getcwd() + r'\Modelo\%s'%nome, 'w').write(interior)
    return

def add_col(text):
    lista = [n_col.split('@')[0] for n_col in (col for col in text.split('_') if '@' in col)]
    return lista

def cria_xlsx(text = '',
              arquivo_XLSX = 'planilha.xlsx',
              dic = {'email': None, 'nome': None, 'bool': False}):
    [dic.setdefault(key) for key in add_col(text)]
    return pd.DataFrame(dic, index = [0]).to_excel(getcwd() + r'\Modelo\%s'%arquivo_XLSX + '.xlsx',
                                                   sheet_name = 'Criado em %s'%strftime('%d-%m-%Y', localtime()),
                                                   index = False)
