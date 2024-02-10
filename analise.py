import pandas as pd
import win32com.client as win32

#importar a base de dados
#ler a base de dados
tabela = pd.read_excel('Vendas.xlsx')

print(tabela)

#faturamento
faturamento = tabela[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#quantidade 

quantidade = tabela[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)


#ticket medio
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

#enviar relatório por email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'shototodoroki163@gmail.com'
mail.Subject = ' Relatório das Vendas'
mail.HTMLBody = f'''
<p>Bom dia Prezados</p>,

<p>Segue o relatório das vendas de todas as lojas:</p>

<p><strong>Faturamento :</strong></p>:
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2F}'.format})}

<p><strong>Quantidade :</strong></p>:
{quantidade.to_html()}

<p><strong>Ticket Médio: </strong></p>
{ticket_medio.to_html()}

Qualquer dúvida estou a disposição

Att Jones

'''
mail.Send()
print('email send!')