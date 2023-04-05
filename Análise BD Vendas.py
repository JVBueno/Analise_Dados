import pandas as pd
import openpyxl
import win32com.client as win32

# Importação do banco de dados e sua devida extensão, neste caso um arquivo xlsx
# O arquivo deve ser anexado ao projeto
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualização completa da base de dados no Python
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print('-'*50)

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print (faturamento)
print('-'*50)

# Quantidade de produto vendido por loja
quantidadeporloja = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidadeporloja)
print('-'*50)

# Ticket médio por produto em cada loja
ticketmedio = (faturamento['Valor Final'] / quantidadeporloja['Quantidade']).to_frame()
ticketmedio = ticketmedio.rename(columns={0:'Valor do Ticket'})
print(ticketmedio)
print('-'*50)

# Enviar email com relatório

# O outlook deve estar instalado como app, não funciona com navegador
outlook = win32.Dispatch('Outlook.application')
mail = outlook.CreateItem(0)
mail.To = '>>> Inserir Email destino <<<'
mail.Subject = '>>> Inserir Assunto <<<'

# Corpo do email - Formatado como um arquivo html
mail.HTMLBody = f''' <p>Prezados,</p>

<p>Segue o Relatório de vendas por cada loja:</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>Quntidade Vendida:</p>
{quantidadeporloja.to_html(formatters={'Quantidade':'R${:,.2f}'.format})}

<p>Ticket Médio dos produtos em cada loja:</p>
{ticketmedio.to_html(formatters={'Valor do Ticket':'R${:,.2f}'.format})}

<p>Qualqual dúvida estou à disposição.</p>

<p>Att, João Victor</p>
'''

mail.Display()     #mail.Display() para visualizar/testar
                   #mail.Send() para enviar

print('''
Email Enviado com Sucesso ! ''')