import pandas as pd
import win32com.client as win32

tabela_vendas = pd.read_excel('Vendas.xlsx')

pd.set_option('display.max_columns', None)
print(tabela_vendas)

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'laisloboteixeira@gmail.com'
mail.Subject = 'Relatório de Vendas'
mail.HTMLBody = f'''
<p>Prezados (as),</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p><strong>Faturamento:</strong></p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p><strong>Quantidade de Venda:</strong></p>
{quantidade.to_html()}

<p><strong>Ticket Médio dos Produtos em cada Loja:</strong></p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Atenciosamente,</p>
<p>Laís</p>
'''

mail.Send()

print('E-mail enviado')
