from tkinter import *
from tkinter import ttk
import pandas as pd
import win32com.client as win32

def main():
    # pegando as variaveis
    arq = file.get()
    colGrup = colGrupamento.get()
    colQtd = colQtde.get()
    colTot = colTotal.get()
    desti = dest.get()
    sub = subject.get()

    print(arq + "\n" + colGrup+ "\n" + colQtd + "\n" + colTot + "\n" + desti + "\n" + sub +"\n")
    # importar base de dados
    tabela = pd.read_excel(arq)

    # realizar o faturamento
    faturamento = tabela[[colGrup, colTot]].groupby(colGrup).sum()

    # realizar contagem de produtos
    qtdProdutos = tabela[[colGrup, colQtd]].groupby(colGrup).sum()

    # realizar o ticket médio por produto em cada loja
    ticketMedio = (faturamento[colTot]/qtdProdutos[colQtd]).to_frame()
    ticketMedio = ticketMedio.rename(columns={0: 'Média'})

    # enviar email
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    email.To = desti
    email.Subject = sub
    email.HTMLBody = f'''
    <p>Este é um teste de envio de email feito em python</p>

    <p>Faturamento</p>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format}).replace(',','_').replace('.',',').replace('_','.')}
    
    <p>Quantidade Vendida</p>
    {qtdProdutos.to_html()}
    <br>
    <p>Ticket Médio Vendido Em cada Loja</p>
    {ticketMedio.to_html(formatters={'Média': 'R${:,.2f}'.format}).replace(',','_').replace('.',',').replace('_','.')}
    
    '''

    email.Send()
    print("Email enviado")
    result['text'] = "Email enviado"

# Interface com usuário
janela = Tk()
janela.title("Anlise de Base de Dados")

text_informativo = ttk.Label(janela, text="\tPegando dados de uma base, processando e mandando pelo email\n"
                                          "\tNeste Progrma você irá dar os nomes das colunas de um arquivo Excel\n"
                                          "\t\tpara receber o processamento dos dados delas.")
text_informativo.grid(column=0, row=0, padx=10, pady=15)

# pega o nome do arquivo que será usado
lblFile = ttk.Label(janela, text="Digite o nome do arquivo ")
lblFile.grid(column=0, row=1, padx=10, pady=10)
file = ttk.Entry(janela)
file.grid(column=1, row=1, padx=10, pady=10)

# pega o nome da coluna que sera usada para agrupar os valores
lblColGrupamento = ttk.Label(janela, text="Digite o nome da coluna que será usada para agrupar os valores ")
lblColGrupamento.grid(column=0, row=2, padx=10, pady=10)
colGrupamento = ttk.Entry(janela)
colGrupamento.grid(column=1, row=2, padx=15, pady=10)

# pega o nome da coluna das quantidades de produtos vendidos
lblColQtde = ttk.Label(janela, text="Digite o nome da coluna de quantidades ")
lblColQtde.grid(column=0, row=3, padx=10, pady=10)
colQtde = ttk.Entry(janela)
colQtde.grid(column=1, row=3, padx=15, pady=10)

# pega o nome da coluna dos valores totais dos produtos vendidos
lblColTotal = ttk.Label(janela, text="Digite o nome da coluna de valores totais ")
lblColTotal.grid(column=0, row=4, padx=10, pady=15)
colTotal = ttk.Entry(janela)
colTotal.grid(column=1, row=4, padx=15, pady=15)


# pega o email do destinatário
lblEmail = ttk.Label(janela, text="Digite o Email do destinatário ")
lblEmail.grid(column=0, row=5, padx=10, pady=10)
dest = ttk.Entry(janela)
dest.grid(column=1, row=5, padx=15, pady=10)

# pega o assunto do email
lblSubject = ttk.Label(janela, text="Digite o assunto do email ")
lblSubject.grid(column=0, row=6, padx=10, pady=10)
subject = ttk.Entry(janela)
subject.grid(column=1, row=6, padx=15, pady=15)

btnEnviar = ttk.Button(janela, text='Enviar', command=main)
btnEnviar.grid(column=0, row=7, padx=25, pady=15)

result = ttk.Label(janela, text="")
result.grid(column=0, row=8, padx=25, pady=20)

janela.mainloop()
