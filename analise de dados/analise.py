import pandas as pd
import win32com.client as win32

#objetivos
#fazer uma analise com a base de dados e mandar por email os resultados de faturamento por loja, quantidade vendida por loja e ticket médio por loja


#importar a base de dados
#ler a base de dados
tabela = pd.read_excel('Vendas.xlsx')

#visualizando a base de dados
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
