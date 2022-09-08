import pandas as pd

tabela_vendas = pd.read_excel("Vendas.xlsx")
pd.set_option('display.max_columns', None)
print(tabela_vendas)
faturamento = tabela_vendas[["ID Loja","Valor Final"]].groupby("ID Loja").sum()
print(faturamento)
quantidade_produtos = tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()
print(quantidade_produtos)
ticket_medio = (faturamento["Valor Final"] / quantidade_produtos["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'bruno-coelho@hotmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p> Prezados, </p>
<p> Segue um relatório de vendas por cada loja </p>

<p> Faturamento: </p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p> Quantidade Vendida: </p>
{quantidade_produtos.to_html()}

<p> Ticket médio dos produtos em cada loja: </p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}


<p> Qualque dúvida estou á disposição. </p>
<p> Att; </p>
<p> Bruno Coelho </p>
'''
mail.Send()

print('-'* 50)
print ('E-mail enviado')