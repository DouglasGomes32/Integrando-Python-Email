import pandas as pd
import win32com.client as win32


outlook = win32.Dispatch("outlook.application")
gerentes_df = pd.read_excel('Enviar E-mails.xlsx')
print(gerentes_df)

for i, email in enumerate(gerentes_df['E-mail']):
    gerente = gerentes_df.loc[i, 'Gerente']
    area = gerentes_df.loc[i, 'Relatório']
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = f'Relatório de {area}'
    mail.Body = f'''
    Prezado {gerente},
    Segue em anexo o Relatório de {area} solicitado.
    Qualquer duvida estou a disposição.
    Att,
    '''
    attachment = r'C:\Users\USER\Documents\Python\Conteudo\Integração Python - E-mail\{}.xlsx'.format(area)
    mail.Attachments.Add(attachment)
    mail.Send()
    