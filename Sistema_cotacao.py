import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import requests
import pandas as pd
from datetime import datetime
import numpy as np




requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moeda = requisicao.json()
dados = list(dicionario_moeda.keys())


lista_moedas = list(dados)

def pegar_cotacao():
    moeda = combobox_selecionarmoeda.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[0:2]
    link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    label_textocotacao['text'] = f'A cotação da {moeda} no dia {data_cotacao} foi de R$ {valor_moeda}'

def selecionar_arquivo():
    caminhoarquivo = askopenfilename(title='Selecione o arquivo da moeda')
    var_caminhoarquivo.set(caminhoarquivo)
    if caminhoarquivo:
        label_arquivoselecionado['text'] = f'Arquivo selecionado: {caminhoarquivo}'

def atualizar_cotacoes():
    df = pd.read_excel(var_caminhoarquivo.get())
    moedas = df.iloc[:0]
    datainicial = calendario_datainicial.get()
    datafinal = calendario_datafinal.get()
    ano_inicial = datainicial[-4:]
    mes_inicial = datainicial[3:5]
    dia_inicial = datainicial[0:2]

    ano_final = datafinal[-4:]
    mes_final = datafinal[3:5]
    dia_final = datafinal[0:2]

    for moeda in moedas:
        link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?' \
               f'start_date={ano_inicial}{mes_inicial}{dia_inicial}&' \
               f'end_date={ano_final}{mes_final}{dia_final}'
        requisicao_moeda = requests.get(link)
        cotacoes = requisicao_moeda.json()
        for cotacao in cotacoes:
            timestamp = int(cotacao["timestamp"])
            bid = float(cotacao["bid"])
            data = datetime.fromtimestamp(timestamp)
            data = data.strftime('%d/%m/%Y')
            if data not in df:
                df[data] = np.nan

            df.loc[df.iloc[:,0] == moeda, data] = bid


    df.to_excel('teste.xlsx')
    label_atualizar_cotacoes['text'] = 'Arquivo atualizado com sucesso'




janela = tk.Tk()
janela.title('Ferramenta cotação de moedas')

label_cotacaomoeda = tk.Label(text='Cotação Moeda Única', borderwidth=2, relief='solid', background='blue', foreground='white')
label_cotacaomoeda.grid(row=0, column=0, padx=10, pady=10, sticky='nswe',columnspan=3)

label_selecionarmoeda = tk.Label(text='Selecionar Moeda', borderwidth=2, relief='solid', background='blue', foreground='white', anchor='e')
label_selecionarmoeda.grid(row=1, column=0, padx=10, pady=10, sticky='nswe',columnspan=2)

combobox_selecionarmoeda = ttk.Combobox(values=lista_moedas)
combobox_selecionarmoeda.grid(row=1, column=2,padx=10, pady=10, sticky='nswe')

label_selecionardata = tk.Label(text='Selecionae o dia que deseja saber a cotação', borderwidth=2, relief='solid', background='blue', foreground='white', anchor='e')
label_selecionardata.grid(row=2, column=0, padx=10, pady=10, sticky='nswe',columnspan=2)

calendario_moeda = DateEntry(year=2022, locale='pt_br')
calendario_moeda.grid(row=2, column=2, pady=10, padx=10, sticky='nsew')

label_textocotacao = tk.Label(text='')
label_textocotacao.grid(row=3, column=0, padx=10, pady=10, sticky='nswe')

botao_pegarcotacao = tk.Button(text='Pegar Cotação ', command=pegar_cotacao)
botao_pegarcotacao.grid(row=3, column=2, pady=10, padx=10, sticky='nswe')

#cotação varias moedas

label_cotacaomulti = tk.Label(text='Cotação de multiplas moedas', borderwidth=2, relief='solid', background='blue', foreground='white')
label_cotacaomulti.grid(row=4, column=0, padx=10, pady=10, sticky='nswe',columnspan=3)

label_selecionararquivo = tk.Label(text='Selecione uma coluna em Excel com as Moedas na coluna A',background='gray', foreground='white')
label_selecionararquivo.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='nswe')

var_caminhoarquivo = tk.StringVar()

botao_selecionararquivo = tk.Button(text='Clique aqui para selecionar um arquivo', relief='solid', background='blue', foreground='white', command=selecionar_arquivo)
botao_selecionararquivo.grid(row=5, column=2, pady=10, padx=10, sticky='nswe')

label_arquivoselecionado = tk.Label(text='Nenhum arquivo selecionado', anchor='e')
label_arquivoselecionado.grid(row=6, column=0, columnspan=3, sticky='nswe')

label_datainicial = tk.Label(text='Data inicial')
label_datainicial.grid(row=7, column=0, pady=10, padx=10, sticky='nswe')

calendario_datainicial = DateEntry(year=2022, locale='pt_br', anchor='w')
calendario_datainicial.grid(row=7, column=1, pady=10, padx=10, sticky='nswe')


label_datafinal = tk.Label(text='Data final')
label_datafinal.grid(row=8, column=0, padx=10, pady=10, sticky='nwse')


calendario_datafinal = DateEntry(year=2022, locale='pt_br',anchor='w')
calendario_datafinal.grid(row=8, column=1, padx=10, pady=10, sticky='nwse')


botao_atualizarcotacoes = tk.Button(text='Atualizar Cotações', command=atualizar_cotacoes)
botao_atualizarcotacoes.grid(row=9, column=0 , pady=10, padx=10, sticky='nwes')


label_atualizar_cotacoes = tk.Label(text='',anchor='e')
label_atualizar_cotacoes.grid(row=9, column=1, columnspan=2,padx=10, pady=10, sticky='nswe')


botao_fechar = tk.Button(text='Fechar', command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nswe')

janela.mainloop()