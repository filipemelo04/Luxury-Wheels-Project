from tkinter import ttk
from tkinter import *
import sqlite3
from tkinter.ttk import Combobox
from tkinter import messagebox
from tkcalendar import DateEntry
from datetime import datetime, date
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
from tkinterdnd2 import TkinterDnD, DND_FILES
import shutil
import numpy as np
import schedule
import time
from threading import Thread

class AppDesktop:
    db = 'database/LuxuryWheels_Database.db'
    tipos = ["Mota", "Carro"]
    categorias = ["Pequeno", "Médio", "Grande", "SUV", "Luxo"]
    transmissoes = ["Automático", "Manual"]
    qtd_pessoas = ["1-4", "5-6", "> 7"]
    metodos_pagamento = ["Numerário", "Cartão Multibanco", "MBWay"]

    def __init__(self, root):
        # Inicializar a janela principal
        self.janela = root
        self.janela.title("Aplicação Desktop Luxury Wheels")
        #self.janela.resizable(1, 1)
        self.janela.resizable(height = None, width = None)
        self.janela.wm_iconbitmap('recursos/car_icon.ico')
        self.caminho_destino_arquivo = ''  # Inicializar a variavel
        # Ajustando o tamanho e posição da janela
        largura = 1200
        altura = 540
        x = 30
        y = 30
        self.janela.geometry(f"{largura}x{altura}+{x}+{y}")
        #self.janela.configure(bg="lightblue")  # Alterar a cor de fundo

        # Criar a barra de menu
        menu_bar = Menu(self.janela)

        # Criar um menu "Menu"
        menu = Menu(menu_bar, tearoff=0)
        menu.add_command(label="Dashboard", command=self.pagina_dashboard)
        menu.add_command(label="Veiculos", command=self.pagina_veiculos)
        menu.add_command(label="Clientes", command=self.pagina_clientes)
        menu.add_command(label="Reservas", command=self.pagina_reservas)
        menu.add_command(label="Pagamentos", command=self.pagina_pagamentos)
        menu.add_separator()  # Adiciona uma linha separadora
        menu.add_command(label="Sair", command=self.sair)

        # Adicionar o Menu  à barra de menu
        menu_bar.add_cascade(label="Menu", menu=menu)

        # Configurar a barra de menu na janela
        self.janela.config(menu=menu_bar)

        self.pagina_dashboard()

    def limpar_pagina(self):
        '''Este metodo tem como objetivo sempre que altera a página limpa todos os widgets e inicializa de novo'''
        # Remove todos os widgets da janela
        for widget in self.janela.winfo_children():
            widget.destroy()

        # Recrie a barra de menu após limpar
        menu_bar = Menu(self.janela)

        # Criar um menu "Menu"
        menu = Menu(menu_bar, tearoff=0)
        menu.add_command(label="Dashboard", command=self.pagina_dashboard)
        menu.add_command(label="Veiculos", command=self.pagina_veiculos)
        menu.add_command(label="Clientes", command=self.pagina_clientes)
        menu.add_command(label="Reservas", command=self.pagina_reservas)
        menu.add_command(label="Pagamentos", command=self.pagina_pagamentos)
        menu.add_separator()  # Adiciona uma linha separadora
        menu.add_command(label="Sair", command=self.sair)

        # Adicionar o Menu à barra de menu
        menu_bar.add_cascade(label="Menu", menu=menu)

        # Configurar a barra de menu na janela
        self.janela.config(menu=menu_bar)

# ------------------ INICIO PAGINA DASHBOARD  --------------------------------

    def pagina_dashboard(self):
        # Limpar pagina
        self.limpar_pagina()

        # Titulo da tabela Reservas
        self.titulo_tabela_reservas = Label(text="Reservas Atuais", fg="blue", font=('Calibri', 14, 'bold'))
        self.titulo_tabela_reservas.grid(row=0, column=1, columnspan=2)

        # Tabela com Reservas
        style = ttk.Style()
        style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Calibri', 10))
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 12, 'bold'))
        style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])

        # Estrutura da tabela
        self.tabela_reservas = ttk.Treeview(self.janela, height=5, columns=('col1', 'col2', 'col3', 'col4', 'col5'),
                                            style="mystyle.Treeview")
        self.tabela_reservas.grid(row=1, column=0, columnspan=6)
        # Cabeçalhos
        self.tabela_reservas.heading('#0', text='Matricula do Veiculo', anchor=CENTER)
        self.tabela_reservas.heading('#1', text='NIF do Cliente', anchor=CENTER)
        self.tabela_reservas.heading('#2', text='Data de Inicio da Reserva', anchor=CENTER)
        self.tabela_reservas.heading('#3', text='Data de Término da Reserva', anchor=CENTER)
        self.tabela_reservas.heading('#4', text='Custo Total (€)', anchor=CENTER)
        self.tabela_reservas.heading('#5', text='Dias Restantes', anchor=CENTER)
        # Configuração da largura das colunas
        self.tabela_reservas.column('#0', width=80, anchor=CENTER)
        self.tabela_reservas.column('col1', width=110, anchor=CENTER)
        self.tabela_reservas.column('col2', width=120, anchor=CENTER)
        self.tabela_reservas.column('col3', width=120, anchor=CENTER)
        self.tabela_reservas.column('col4', width=120, anchor=CENTER)
        self.tabela_reservas.column('col5', width=110, anchor=CENTER)
        self.get_reservas_atuais()

        #Label Separatoria
        self.label_separatoria = Label(text='')
        self.label_separatoria.grid(row=0, column=6)

        # Titulo da tabela Clientes
        self.titulo_tabela_clientes = Label(text="Novos Clientes", fg="blue", font=('Calibri', 14, 'bold'))
        self.titulo_tabela_clientes.grid(row=0, column=8)

        # Tabela com Clientes
        style = ttk.Style()
        style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Calibri', 11))
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 13, 'bold'))
        style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])

        # Estrutura da tabela
        self.tabela_novos_clientes = ttk.Treeview(self.janela, height=5, columns=('col1', 'col2'),
                                            style="mystyle.Treeview")
        self.tabela_novos_clientes.grid(row=1, column=7, columnspan=8)
        # Cabeçalhos
        self.tabela_novos_clientes.heading('#0', text='NIF', anchor=CENTER)
        self.tabela_novos_clientes.heading('#1', text='Nome', anchor=CENTER)
        self.tabela_novos_clientes.heading('#2', text='Idade', anchor=CENTER)
        # Configuração da largura das colunas
        self.tabela_novos_clientes.column('#0', width=100, anchor=CENTER)
        self.tabela_novos_clientes.column('col1', width=120, anchor=CENTER)
        self.tabela_novos_clientes.column('col2', width=80, anchor=CENTER)
        self.get_novos_clientes()

        self.label_separador = Label(text='')
        self.label_separador.grid(row=2, column=0)

        # Titulo da tabela Veiculos Disponiveis
        self.titulo_tabela_veiculos = Label(text="Veiculos Disponiveis", width=16, fg="blue", font=('Calibri', 14, 'bold'))
        self.titulo_tabela_veiculos.grid(row=3, column=2, columnspan=2)

        # Criação da FRAME Informação Veiculo
        frame_filtrar_veiculos = LabelFrame(self.janela, text="Filtros Veiculo", font=('Calibri', 16, 'bold'))
        frame_filtrar_veiculos.grid(row=4, column=0, columnspan=2)

        #Label Tipos
        self.label_filtro_tipos = Label(frame_filtrar_veiculos, text="Tipos:", width=10)
        self.label_filtro_tipos.grid(row=0, column=0)
        self.input_filtro_tipos = Combobox(frame_filtrar_veiculos, values=self.tipos, width=10)
        self.input_filtro_tipos.set(self.tipos[1])
        self.input_filtro_tipos.grid(row=0, column=1, sticky='w')

        # Label Categoria
        self.label_filtro_categoria = Label(frame_filtrar_veiculos, text="Categorias:", width=10, anchor="e")
        self.label_filtro_categoria.grid(row=1, column=0)
        self.input_filtro_categoria = Combobox(frame_filtrar_veiculos, values=self.categorias, width=10)
        self.input_filtro_categoria.set(self.categorias[1])
        self.input_filtro_categoria.grid(row=1, column=1, sticky='w')

        # Botão Filtrar Veiculos
        self.botao_filtro_veiculo = ttk.Button(frame_filtrar_veiculos, text="FILTRAR", command=self.get_filtrar_veiculo)
        self.botao_filtro_veiculo.grid(row=2, column=0, columnspan=2, sticky=W + E)

        # Mensagem Informativa
        self.mensagem_dashboard = Label(frame_filtrar_veiculos, text='', fg='red')
        self.mensagem_dashboard.grid(row=3, column=0, columnspan=2, sticky=W + E)

        self.get_filtrar_veiculo()

        self.label_separador = Label(text='')
        self.label_separador.grid(row=5)

        # Titulo da tabela Clientes
        self.titulo_tabela_veiculos = Label(text="Veiculos com datas de inspeção e revisão a expirar", width=50, fg="blue", font=('Calibri', 14, 'bold'))
        self.titulo_tabela_veiculos.grid(row=6, column=0, columnspan=4)

        self.mtd_tabela_inspecao_revisao_expirada()

        # Titulo Gráficos Demonstrativos
        self.titulo_graficos_demostrativos = Label(text="Representações Gráficas", fg="blue", font=('Calibri', 14, 'bold'))
        self.titulo_graficos_demostrativos.grid(row=6, column=9)

        # Criação da FRAME Graficos Demonstrativos
        frame_graficos = LabelFrame(self.janela, font=('Calibri', 16, 'bold'))
        frame_graficos.grid(row=7, column=9)

        # Botão Grafico 1
        self.botao1 = ttk.Button(frame_graficos, text="Pagamento por Dia", command=self.grafico_pagamento_data)
        self.botao1.grid(row=0, column=0, columnspan=2, sticky=W + E)

        # Botão Grafico 2
        self.botao2 = ttk.Button(frame_graficos, text="Reservas por Veiculo", command=self.grafico_ocupacao_veiculo)
        self.botao2.grid(row=1, column=0, columnspan=2, sticky=W + E)

        # Botão Grafico 3
        self.botao3 = ttk.Button(frame_graficos, text="Reservas por Cliente", command=self.grafico_ocupacao_cliente)
        self.botao3.grid(row=2, column=0, columnspan=2, sticky=W + E)

        # Botão Grafico 4
        self.botao4 = ttk.Button(frame_graficos, text="Reservas de Veiculos por Tipo", command=self.grafico_veiculo_tipo)
        self.botao4.grid(row=3, column=0, columnspan=2, sticky=W + E)


    def get_novos_clientes(self):
        '''Coloca no gráfico os últimos 5 clientes criados'''
        # Limpar tabela
        registos_tabela = self.tabela_novos_clientes.get_children()
        for linha in registos_tabela:
            self.tabela_novos_clientes.delete(linha)

        # Consulta SQL
        query = 'SELECT * FROM clientes ORDER BY id DESC LIMIT 5'
        clientes = self.db_consulta(query)

        # Escrever dados na tabela
        for linha in clientes:
            self.tabela_novos_clientes.insert('', 0, text=linha[1], values=(linha[2], linha[3]))

    def get_filtrar_veiculo(self):
        '''Cria tabela com veiculos consoante os filtros: tipo e categoria'''
        self.mensagem_dashboard['text'] = ''

        # Verifica se os filtros não estão em branco
        if len(self.input_filtro_tipos.get()) == 0 or len(self.input_filtro_categoria.get()) == 0:
            self.mensagem_dashboard['text'] = 'Selecione os filtros !!'
            return

        # Verifica se nos filtros não tem Mota e SUV simultaneamente
        if self.input_filtro_tipos.get() == "Mota" and self.input_filtro_categoria.get() == "SUV":
            self.mensagem_dashboard['text'] = 'Uma mota não pode ser SUV !!'
            return

        # Tabela com Reservas
        style = ttk.Style()
        style.configure("mystyle.Treeview", highlightthickness=1, bd=1, font=('Calibri', 10))
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 12, 'bold'))
        style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])

        # Estrutura da tabela
        self.tabela_veiculo_filtro = ttk.Treeview(self.janela, height=5, columns=('col1', 'col2', 'col3', 'col4', 'col5',
                                                                            'col6', 'col7', 'col8', 'col9'), style="mystyle.Treeview")
        self.tabela_veiculo_filtro.grid(row=4, column=2, columnspan=10)
        # Cabeçalhos
        self.tabela_veiculo_filtro.heading('#0', text='Matricula', anchor=CENTER)
        self.tabela_veiculo_filtro.heading('#1', text='Marca', anchor=CENTER)
        self.tabela_veiculo_filtro.heading('#2', text='Modelo', anchor=CENTER)
        self.tabela_veiculo_filtro.heading('#3', text='Tipo', anchor=CENTER)  # Cabeçalho 0
        self.tabela_veiculo_filtro.heading('#4', text='Categoria', anchor=CENTER)  # Cabeçalho 1
        self.tabela_veiculo_filtro.heading('#5', text='Transmissão', anchor=CENTER)
        self.tabela_veiculo_filtro.heading('#6', text='Valor Diário', anchor=CENTER)
        self.tabela_veiculo_filtro.heading('#7', text='N pessoas', anchor=CENTER)
        self.tabela_veiculo_filtro.heading('#8', text='Data próxima revisão', anchor=CENTER)
        self.tabela_veiculo_filtro.heading('#9', text='Data ultima inspeção', anchor=CENTER)
        # Configuração da largura das colunas
        self.tabela_veiculo_filtro.column('#0', width=80, anchor=CENTER)
        self.tabela_veiculo_filtro.column('col1', width=80, anchor=CENTER)
        self.tabela_veiculo_filtro.column('col2', width=80, anchor=CENTER)
        self.tabela_veiculo_filtro.column('col3', width=80, anchor=CENTER)
        self.tabela_veiculo_filtro.column('col4', width=80, anchor=CENTER)
        self.tabela_veiculo_filtro.column('col5', width=80, anchor=CENTER)
        self.tabela_veiculo_filtro.column('col6', width=80, anchor=CENTER)
        self.tabela_veiculo_filtro.column('col7', width=80, anchor=CENTER)
        self.tabela_veiculo_filtro.column('col8', width=130, anchor=CENTER)
        self.tabela_veiculo_filtro.column('col9', width=130, anchor=CENTER)
        self.get_veiculos_filtro()

    def get_veiculos_filtro(self):
        '''Coloca na tabela os veiculos consoante os filtros: tipo e categoria'''
        # Limpar tabela
        registos_tabela = self.tabela_veiculo_filtro.get_children()
        for linha in registos_tabela:
            self.tabela_veiculo_filtro.delete(linha)

        tipo = self.input_filtro_tipos.get()
        categoria = self.input_filtro_categoria.get()

        # Consulta SQL
        #query = 'SELECT * FROM veiculos WHERE ocupado=0 ORDER BY ' + filtro + ' ASC'
        query = 'SELECT * FROM veiculos WHERE ocupado=0 AND tipo=? AND categoria=?'
        parametros = (tipo, categoria)
        veiculos = self.db_consulta(query, parametros)

        # Escrever dados na tabela
        for linha in veiculos:
            self.tabela_veiculo_filtro.insert('', 0, text=linha[1], values=(linha[2], linha[3], linha[4],
                                                                      linha[5], linha[6], linha[7], linha[8],
                                                                      linha[10], linha[11]))

    def mtd_tabela_inspecao_revisao_expirada(self):
        ''' Cria tabela com veiculos com inspeção ou revisão a expirar '''
        style = ttk.Style()
        style.configure("mystyle.Treeview", highlightthickness=1, bd=1, font=('Calibri', 10))
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 12, 'bold'))
        style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])

        # Estrutura da tabela
        self.tabela_inspecao_revisao_expirada = ttk.Treeview(self.janela, height=4, columns=('col1', 'col2', 'col3', 'col4', 'col5'),
                                                             style="mystyle.Treeview")
        self.tabela_inspecao_revisao_expirada.grid(row=7, column=0, columnspan=8)
        # Cabeçalhos
        self.tabela_inspecao_revisao_expirada.heading('#0', text='Matricula', anchor=CENTER)
        self.tabela_inspecao_revisao_expirada.heading('#1', text='Marca', anchor=CENTER)
        self.tabela_inspecao_revisao_expirada.heading('#2', text='Modelo', anchor=CENTER)
        self.tabela_inspecao_revisao_expirada.heading('#3', text='Data próxima revisão', anchor=CENTER)
        self.tabela_inspecao_revisao_expirada.heading('#4', text='Data ultima inspeção', anchor=CENTER)
        self.tabela_inspecao_revisao_expirada.heading('#5', text='Comentário ', anchor=CENTER)
        # Definindo a tag para a coluna vermelha
        self.tabela_inspecao_revisao_expirada.tag_configure('vermelho', background='red')
        # Configuração da largura das colunas
        self.tabela_inspecao_revisao_expirada.column('#0', width=80, anchor=CENTER)
        self.tabela_inspecao_revisao_expirada.column('col1', width=80, anchor=CENTER)
        self.tabela_inspecao_revisao_expirada.column('col2', width=80, anchor=CENTER)
        self.tabela_inspecao_revisao_expirada.column('col3', width=120, anchor=CENTER)
        self.tabela_inspecao_revisao_expirada.column('col4', width=120, anchor=CENTER)
        self.tabela_inspecao_revisao_expirada.column('col5', width=330, anchor=CENTER)
        self.get_veiculos_exprirar_datas()

    def get_veiculos_exprirar_datas(self):
        '''Coloca veiculos com datas de revisão ou inspação a expirar'''
        # Limpar tabela
        registos_tabela = self.tabela_inspecao_revisao_expirada.get_children()
        for linha in registos_tabela:
            self.tabela_inspecao_revisao_expirada.delete(linha)

        # Consulta SQL
        query = 'SELECT * FROM veiculos'
        veiculos = self.db_consulta(query)

        # Escrever dados na tabela
        for linha in veiculos:
            # Calcula a diferença o dia da proxima revisão e o dia atual
            dias_revisao = self.diferenca_atual_em_dias(linha[10])
            # Calcula a diferença entre o dia da ultima inpeção e o dia atual mais 1 ano
            dias_inpecao = self.diferenca_atual_em_dias(linha[11]) + 365

            if 0 < dias_revisao < 15:
                comentario = "Faltam " + str(dias_revisao) + " dias para a próxima revisão do veiculo."
                self.tabela_inspecao_revisao_expirada.insert('', 0, text=linha[1], values=(linha[2], linha[3],
                                                                          linha[10], linha[11], comentario))
            elif dias_revisao < 0:
                comentario = "Já passaram " + str(abs(dias_revisao)) + " dias para a próxima revisão do veiculo."
                self.tabela_inspecao_revisao_expirada.insert('', 0, text=linha[1], values=(linha[2], linha[3],
                                                                                           linha[10], linha[11], comentario))
            if 0 < dias_inpecao < 15:
                comentario = "Faltam " + str(dias_inpecao) + " dias para a próxima inspeção do veiculo."
                self.tabela_inspecao_revisao_expirada.insert('', 0, text=linha[1], values=(linha[2], linha[3],
                                                                          linha[10], linha[11], comentario))
            elif dias_inpecao < 0:
                comentario = "Já passaram " + str(abs(dias_inpecao)) + " dias para a próxima inspeção do veiculo."
                self.tabela_inspecao_revisao_expirada.insert('', 0, text=linha[1], values=(linha[2], linha[3],
                                                                          linha[10], linha[11], comentario))

    def get_reservas_atuais(self):
        '''Coloca na tabela as reservas que estão ativas no momento'''
        # Limpar tabela
        reservas_tabela = self.tabela_reservas.get_children()
        for linha in reservas_tabela:
            self.tabela_reservas.delete(linha)

        # Consulta SQL
        query = 'SELECT * FROM reservas WHERE estado=1 ORDER BY estado ASC, data_fim ASC'
        reservas = self.db_consulta(query)

        # Escrever dados na tabela
        for linha in reservas:
            dias_restantes = self.diferenca_atual_em_dias(linha[4]) if linha[6] else ''
            estado = "Em execução" if linha[6] else "Finalizada"
            self.tabela_reservas.insert('', 0, text=linha[1],
                                        values=(linha[2], linha[3], linha[4], linha[5], dias_restantes, estado))

    def grafico_pagamento_data(self):
        ''' Mostra gráfico de barras relativo aos pagamentos feitos nas respetivas datas'''
        # Consulta SQL das datas de pagamentos e somatorio dos custos por cada data
        query = 'SELECT data_pagamento, SUM(custo) FROM pagamentos GROUP BY data_pagamento'
        pagamentos = self.db_consulta(query)

        # Cria List com datas dos pagamentos
        datas = [pagamento[0] for pagamento in pagamentos]
        # Cria List com somatorios dos pagamentos por cada dia
        custos = [round(pagamento[1], 3) for pagamento in pagamentos]

        # ORDENAR DATAS E CUSTOS RESPETIVAMENTE
        # Combinar as listas em pares (data, custo)
        pares = list(zip(datas, custos))
        # Ordenar os pares pela data, convertendo para datetime
        pares_ordenados = sorted(pares, key=lambda x: datetime.strptime(x[0], '%d/%m/%Y'))
        # Separar novamente em listas de datas e custos
        datas_ordenadas = [par[0] for par in pares_ordenados]
        custos_ordenados = [par[1] for par in pares_ordenados]

        # Converter datas para índices numéricos (para regressão linear)
        x = list(range(len(datas_ordenadas)))  # Usando índices como valores de x
        y = custos_ordenados

        coef = np.polyfit(x, y, 1)  # Coeficientes da regressão linear (a, b)
        linear_fit = np.poly1d(coef)  # Função da linha de regressão

        # Criar a figura do Matplotlib
        figura = Figure(figsize=(5, 4), dpi=100)
        eixo = figura.add_subplot(111)

        # Gráfico de barras
        eixo.bar(x, y, color='blue', label='Pagamentos')
        # Adicionar linha de regressão
        eixo.plot(x, linear_fit(x), color='red', label='Regressão Linear', linewidth=2)

        # Ajustes visuais
        eixo.set_title("Pagamentos por Dia", fontsize=14)
        eixo.set_xlabel("Data", fontsize=12)
        eixo.set_ylabel("Custo", fontsize=12)
        eixo.set_xticks(x)
        eixo.set_xticklabels(datas_ordenadas, rotation=45)  # Datas originais no eixo X
        eixo.legend()
        # Ajustar layout para que a legenda do eixo do x não seja cortada
        figura.tight_layout()

        # Cria janela para mostrar gráfico
        self.janela_grafico_pagamento = Toplevel()
        self.janela_grafico_pagamento.title = "Grafico de Pagamentos por Data"
        self.janela_grafico_pagamento.resizable(1, 1)
        self.janela_grafico_pagamento.wm_iconbitmap('recursos/car_icon.ico')
        # Ajustando o tamanho e posição da janela
        largura = 500
        altura = 400
        x = 100
        y = 100
        self.janela_grafico_pagamento.geometry(f"{largura}x{altura}+{x}+{y}")

        canvas = FigureCanvasTkAgg(figura, master=self.janela_grafico_pagamento)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.grid(row=8, column=4, columnspan=3)
        canvas.draw()

    def grafico_ocupacao_veiculo(self):
        ''' Mostra gráfico relativo à utilização de cada veiculo'''
        # Consulta SQL das matriculas de cada pagamento
        query = 'SELECT matricula_veiculo FROM pagamentos ORDER BY matricula_veiculo'
        matriculas = self.db_consulta(query)
        matriculas = [matricula[0] for matricula in matriculas]
        # Dicionário para contar as ocorrências
        contagem = {}
        for matricula in matriculas:
            if matricula in contagem:
                contagem[matricula] += 1  # Incrementar o contador
            else:
                contagem[matricula] = 1  # Inicializar o contador
        # Ordenar os valores por ordem decrescente
        contagem = dict(sorted(contagem.items(), key=lambda item: item[1], reverse=True))
        # Converter itens em lista e fatiar os primeiros 10
        contagem = list(contagem.items())[:10]
        # Separar os resultados em listas
        matriculas_unicas = [i[0] for i in contagem]
        quantidade_repeticoes = [i[1] for i in contagem]
        cores = ['gold', 'lightblue', 'lightgreen', 'pink', 'blue', 'red', 'orange', 'purple', 'green', 'grey', 'lightgray',
                 'cyan', 'magenta', 'yellow', 'black']
        # Determinar qual fatia explodir (maior valor)
        maior_indice = quantidade_repeticoes.index(max(quantidade_repeticoes))  # Índice da maior fatia
        explode = [0.1 if i == maior_indice else 0 for i in range(len(quantidade_repeticoes))]  # Explodir apenas a maior

        # Criar figura do Matplotlib
        figura = Figure(figsize=(5, 4), dpi=100)
        eixo = figura.add_subplot(111)
        eixo.pie(
            quantidade_repeticoes,
            labels=matriculas_unicas,
            autopct='%1.1f%%',
            startangle=90,
            explode=explode,  # Explodir a maior fatia
            colors=cores
        )
        eixo.set_title("Distribuição por Utilização de Veiculos")

        # Cria janela para mostrar gráfico
        self.janela_grafico_ocupacao = Toplevel()
        self.janela_grafico_ocupacao.title = "Grafico"
        self.janela_grafico_ocupacao.resizable(1, 1)
        self.janela_grafico_ocupacao.wm_iconbitmap('recursos/car_icon.ico')
        # Ajustando o tamanho e posição da janela
        largura = 500
        altura = 400
        x = 100
        y = 100
        self.janela_grafico_ocupacao.geometry(f"{largura}x{altura}+{x}+{y}")

        canvas = FigureCanvasTkAgg(figura, master=self.janela_grafico_ocupacao)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.grid(row=8, column=4, columnspan=3)
        canvas.draw()

    def grafico_ocupacao_cliente(self):
        ''' Mostra gráfico relativo à utilização de cada cliente'''
        # Consulta SQL do nif de clientes por pagamento
        query = 'SELECT nif_cliente FROM pagamentos ORDER BY nif_cliente'
        nifs = self.db_consulta(query)
        nifs = [nif[0] for nif in nifs]
        # Dicionário para contar as ocorrências
        contagem = {}
        for nif in nifs:
            if nif in contagem:
                contagem[nif] += 1  # Incrementar o contador
            else:
                contagem[nif] = 1  # Inicializar o contador
        # Separar os resultados em listas
        nifs_unicos = list(contagem.keys())
        quantidade_repeticoes = list(contagem.values())
        cores = ['gold', 'lightblue', 'lightgreen', 'pink', 'blue', 'red', 'orange', 'purple', 'green', 'grey',
                 'lightgray',
                 'cyan', 'magenta', 'yellow', 'black']
        # Determinar qual fatia explodir (maior valor)
        maior_indice = quantidade_repeticoes.index(max(quantidade_repeticoes))  # Índice da maior fatia
        explode = [0.1 if i == maior_indice else 0 for i in range(len(quantidade_repeticoes))]  # Explodir apenas a maior

        # Criar figura do Matplotlib
        figura = Figure(figsize=(5, 4), dpi=100)
        eixo = figura.add_subplot(111)
        eixo.pie(
            quantidade_repeticoes,
            labels=nifs_unicos,
            autopct='%1.1f%%',
            startangle=90,
            explode=explode,  # Explodir a maior fatia
            colors=cores
        )
        eixo.set_title("Distribuição por Utilização de Clientes")

        # Cria janela para mostrar gráfico
        self.janela_grafico_clientes = Toplevel()
        self.janela_grafico_clientes.title = "Grafico"
        self.janela_grafico_clientes.resizable(1, 1)
        self.janela_grafico_clientes.wm_iconbitmap('recursos/car_icon.ico')
        # Ajustando o tamanho e posição da janela
        largura = 500
        altura = 400
        x = 100
        y = 100
        self.janela_grafico_clientes.geometry(f"{largura}x{altura}+{x}+{y}")

        canvas = FigureCanvasTkAgg(figura, master=self.janela_grafico_clientes)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.grid(row=8, column=4, columnspan=3)
        canvas.draw()

    def grafico_veiculo_tipo(self):
        ''' Cria gráfico de barras relativo à utilização de veiculos por tipo'''
        # Consulta SQL das datas de pagamentos e somatorio dos custos por cada data
        query = 'SELECT matricula_veiculo FROM pagamentos ORDER BY matricula_veiculo'
        matriculas = self.db_consulta(query)
        matriculas = [matricula[0] for matricula in matriculas]

        # Consulta SQL das matriculas de todas as motas
        query = 'SELECT matricula FROM veiculos WHERE tipo="Mota"'
        matriculas_motas = self.db_consulta(query)
        matriculas_motas = [matricula[0] for matricula in matriculas_motas]

        # Contar quantas motas e quantos carros têm reservas
        somatorios = {"Mota":0, "Carro":0}
        for matricula in matriculas:
            if matricula in matriculas_motas:
                somatorios["Mota"] += 1
            else: somatorios["Carro"] += 1

        # Criar a figura e o eixo
        figura = Figure(figsize=(6, 4), dpi=100)
        eixo = figura.add_subplot(111)
        eixo.bar(somatorios.keys(), somatorios.values(), color="blue")
        eixo.set_title("Gráfico de Barras")
        eixo.set_xlabel("Tipos")
        eixo.set_ylabel("Nº Reservas")

        # Cria janela para mostrar gráfico
        self.janela_grafico_pagamento = Toplevel()
        self.janela_grafico_pagamento.title = "Grafico de Pagamentos por Data"
        self.janela_grafico_pagamento.resizable(1, 1)
        self.janela_grafico_pagamento.wm_iconbitmap('recursos/car_icon.ico')
        # Ajustando o tamanho e posição da janela
        largura = 600
        altura = 400
        x = 100
        y = 100
        self.janela_grafico_pagamento.geometry(f"{largura}x{altura}+{x}+{y}")

        canvas = FigureCanvasTkAgg(figura, master=self.janela_grafico_pagamento)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.grid(row=8, column=4, columnspan=3)
        canvas.draw()

# ------------------ FIM PAGINA DASHBOARD  --------------------------------


# ------------------ INICIO PAGINA VEICULOS  --------------------------------

    def pagina_veiculos(self):
        ''' Criação da pagina Veiculos '''
        # Limpar pagina
        self.limpar_pagina()

        # Tabela com Veiculos
        style = ttk.Style()
        style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Calibri', 11))
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 13, 'bold'))
        style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])

        # Estrutura da tabela
        self.tabela_veiculos = ttk.Treeview(self.janela, height=15, columns=('col1', 'col2', 'col3', 'col4', 'col5',
                                                                            'col6', 'col7', 'col8', 'col9', 'col10',
                                                                            'col11', 'col12'), style="mystyle.Treeview")
        self.tabela_veiculos.grid(row=0, column=0, columnspan=8)
        # Cabeçalhos
        self.tabela_veiculos.heading('#0', text='Matricula', anchor=CENTER)
        self.tabela_veiculos.heading('#1', text='Marca', anchor=CENTER)
        self.tabela_veiculos.heading('#2', text='Modelo', anchor=CENTER)
        self.tabela_veiculos.heading('#3', text='Tipo', anchor=CENTER)
        self.tabela_veiculos.heading('#4', text='Categoria', anchor=CENTER)
        self.tabela_veiculos.heading('#5', text='Transmissão', anchor=CENTER)
        self.tabela_veiculos.heading('#6', text='Valor Diário', anchor=CENTER)
        self.tabela_veiculos.heading('#7', text='N pessoas', anchor=CENTER)
        self.tabela_veiculos.heading('#8', text='Ocupação', anchor=CENTER)
        self.tabela_veiculos.heading('#9', text='Data próxima revisão', anchor=CENTER)
        self.tabela_veiculos.heading('#10', text='Data ultima inspeção', anchor=CENTER)
        self.tabela_veiculos.heading('#11', text='Endereço Imagem', anchor=CENTER)
        # Configuração da largura das colunas
        self.tabela_veiculos.column('#0', width=90, anchor=CENTER)
        self.tabela_veiculos.column('col1', width=100, anchor=CENTER)
        self.tabela_veiculos.column('col2', width=100, anchor=CENTER)
        self.tabela_veiculos.column('col3', width=80, anchor=CENTER)
        self.tabela_veiculos.column('col4', width=80, anchor=CENTER)
        self.tabela_veiculos.column('col5', width=100, anchor=CENTER)
        self.tabela_veiculos.column('col6', width=100, anchor=CENTER)
        self.tabela_veiculos.column('col7', width=80, anchor=CENTER)
        self.tabela_veiculos.column('col8', width=80, anchor=CENTER)
        self.tabela_veiculos.column('col9', width=140, anchor=CENTER)
        self.tabela_veiculos.column('col10', width=140, anchor=CENTER)
        self.tabela_veiculos.column('col11', width=120, anchor=CENTER)

        self.get_veiculos()

        # Mensagem Informativa
        self.mensagem_veiculo = Label(self.janela, text='', fg='red')
        self.mensagem_veiculo.grid(row=1, column=0, columnspan=2, sticky=W + E)

        # Botão adicionar veiculo:
        self.botao_pagina_adicionar_veiculo = ttk.Button(self.janela, text="Adicionar Veiculo",
                                                         command=self.pagina_adicionar_veiculo)
        self.botao_pagina_adicionar_veiculo.grid(row=2, column=0, sticky = W + E)

        # Botao Editar Veiculo
        self.botao_pagina_editar_veiculo = ttk.Button(self.janela, text="Editar Veiculo",
                                                      command=self.pagina_editar_veiculo)
        self.botao_pagina_editar_veiculo.grid(row=2, column=1, sticky = W + E)

        # Botao Eliminar Veiculo
        self.botao_eliminar_veiculo = ttk.Button(self.janela, text="Eliminar Veiculo",
                                                      command=self.eliminar_veiculo)
        self.botao_eliminar_veiculo.grid(row=2, column=2, sticky = W + E)

        # Botao Colocar Veiculo em Manutenção
        self.botao_exportar_dados = ttk.Button(self.janela, text="Colocar Veiculo em Manutenção",
                                               command=self.manutencao_veiculo)
        self.botao_exportar_dados.grid(row=2, column=3, rowspan=2, sticky=W + E)

        # Botao Exportar Dados
        self.botao_exportar_dados = ttk.Button(self.janela, text="Exportar Dados",
                                               command=self.exportar_veiculos)
        self.botao_exportar_dados.grid(row=2, column=4, sticky=W + E)

    def pagina_adicionar_veiculo(self):
        ''' Criação de página para adicionar veiculo '''
        self.mensagem_veiculo['text'] = ''

        self.janela_adicionar_veiculo = Toplevel()
        self.janela_adicionar_veiculo.title = "Adicionar Veiculo"   # VERIFICAR RESULTADO
        self.janela_adicionar_veiculo.resizable(1, 1)
        self.janela_adicionar_veiculo.wm_iconbitmap('recursos/car_icon.ico')

        # Criação da frame Informação Veiculo
        frame_info_veiculo = LabelFrame(self.janela_adicionar_veiculo, text="Informações do veiculo",
                                        font=('Calibri', 16, 'bold'))
        frame_info_veiculo.grid(row=0, column=0, columnspan=3, pady=20)

        # Label matricula do Veiculo
        self.etiqueta_matricula = Label(frame_info_veiculo, text="Matricula do veiculo: ")
        self.etiqueta_matricula.grid(row=1, column=0)
        self.entry_matricula = Entry(frame_info_veiculo, width=23)
        self.entry_matricula.grid(row=1, column=1)

        # Label marca do Veiculo
        self.etiqueta_marca = Label(frame_info_veiculo, text="Marca do veiculo: ")
        self.etiqueta_marca.grid(row=2, column=0)
        self.entry_marca = Entry(frame_info_veiculo, width=23)
        self.entry_marca.grid(row=2, column=1)

        # Label modelo do Veiculo
        self.etiqueta_modelo = Label(frame_info_veiculo, text="Modelo do veiculo: ")
        self.etiqueta_modelo.grid(row=3, column=0)
        self.entry_modelo = Entry(frame_info_veiculo, width=23)
        self.entry_modelo.grid(row=3, column=1)

        # Label Tipo de Veiculo
        self.etiqueta_tipo = Label(frame_info_veiculo, text="Tipo de veiculo: ")  # Etiqueta de texto localizada no frame
        self.etiqueta_tipo.grid(row=4, column=0)  # Posicionamento através de grid
        self.select_tipo = Combobox(frame_info_veiculo, values=self.tipos)  # Caixa de texto (input de texto) localizada no frame
        self.select_tipo.grid(row=4, column=1)

        # Label Categoria
        self.etiqueta_categoria = Label(frame_info_veiculo, text="Categoria do veiculo: ")
        self.etiqueta_categoria.grid(row=5, column=0)
        self.select_categoria = Combobox(frame_info_veiculo, values=self.categorias)
        self.select_categoria.grid(row=5, column=1)

        # Label Transmissão
        self.etiqueta_transmissao = Label(frame_info_veiculo, text="Transmissão do veiculo: ")
        self.etiqueta_transmissao.grid(row=6, column=0)
        self.select_transmissao = Combobox(frame_info_veiculo, values=self.transmissoes)
        self.select_transmissao.grid(row=6, column=1)

        # Label Valor Diário
        self.etiqueta_valor_diario = Label(frame_info_veiculo, text="Valor diário do veiculo (€): ")
        self.etiqueta_valor_diario.grid(row=7, column=0)
        self.entry_valor_diario = Entry(frame_info_veiculo, width=23)
        self.entry_valor_diario.grid(row=7, column=1)

        # Label Quantidade Pessoas
        self.etiqueta_qtd_dpessoas = Label(frame_info_veiculo, text="Quantidade de pessoas do veiculo: ")
        self.etiqueta_qtd_dpessoas.grid(row=8, column=0)
        self.select_qtd_pessoas = Combobox(frame_info_veiculo, values=self.qtd_pessoas)
        self.select_qtd_pessoas.grid(row=8, column=1)

        # Label Data da Ultima Revisão
        self.etiqueta_data_proxima_revisao = Label(frame_info_veiculo, text="Data da proxima revisão do veiculo: ")
        self.etiqueta_data_proxima_revisao.grid(row=9, column=0)
        self.input_data_proxima_revisao = DateEntry(frame_info_veiculo, width=20, background='blue', foreground='white',
                                                borderwidth=2, date_pattern="dd/mm/yyyy")
        self.input_data_proxima_revisao.grid(row=9, column=1)
        self.input_data_proxima_revisao.delete(0, END)

        # Label Data da Ultima Inspeção
        self.etiqueta_data_ult_inspecao = Label(frame_info_veiculo, text="Data da ultima inspeção do veiculo: ")
        self.etiqueta_data_ult_inspecao.grid(row=10, column=0)
        self.input_data_ult_inspecao = DateEntry(frame_info_veiculo, width=20, background='blue', foreground='white',
                                                borderwidth=2, date_pattern="dd/mm/yyyy")
        self.input_data_ult_inspecao.grid(row=10, column=1)
        self.input_data_ult_inspecao.delete(0, END)

        # Criar um rótulo para a área de arrastar e soltar
        self.area_arraste = Label(frame_info_veiculo, text="Arraste uma foto do veiculo para aqui !!", relief="solid", width=50, height=10)
        self.area_arraste.grid(row=11, column=0, columnspan=2)

        # Habilitar a funcionalidade de arrastar e soltar na área
        self.area_arraste.drop_target_register(DND_FILES)
        self.area_arraste.dnd_bind('<<Drop>>', self.processar_arquivo_arrastado)

        # Mensagem Informativa
        self.mensagem_adicionar_veiculo = Label(frame_info_veiculo, text='', fg='red')
        self.mensagem_adicionar_veiculo.grid(row=12, column=0, columnspan=2, sticky=W+E)

        # Botao adicionar veiculo
        self.botao_adicionar_veiculo = ttk.Button(frame_info_veiculo, text="Adicionar Veiculo",
                                                  command=self.adicionar_veiculo)
        self.botao_adicionar_veiculo.grid(row=13, columnspan = 2, sticky = W + E)

    def pagina_editar_veiculo(self):
        ''' Criação de página para editar veiculos '''
        self.mensagem_veiculo['text'] = ''

        # Verifica se algum veiculo foi selecionado
        try:
            self.tabela_veiculos.item(self.tabela_veiculos.selection())['text'][0]
        except IndexError as e:
            self.mensagem_veiculo['text'] = 'Por favor, selecione um veiculo'
            return

        matricula = self.tabela_veiculos.item(self.tabela_veiculos.selection())['text']
        marca = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][0]
        modelo = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][1]
        tipo = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][2]
        categoria = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][3]
        transmissao = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][4]
        valor_diario = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][5]
        qtd_pessoas = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][6]
        ocupado = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][7]
        data_revisao = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][8]
        data_inspecao = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][9]
        endereco_imagem = self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][10]

        # Criar página Editar
        self.janela_editar_veiculo = Toplevel()
        self.janela_editar_veiculo.title = "Editar Veiculo"  # VERIFICAR RESULTADO
        self.janela_editar_veiculo.resizable(1, 1)
        self.janela_editar_veiculo.wm_iconbitmap('recursos/car_icon.ico')

        titulo = Label(self.janela_editar_veiculo, text="Edição de Veiculos", font=('Calibri', 30, 'bold'), fg='blue')
        titulo.grid(row=0, column=0)

        # CRIAR FRAME EDITAR VEICULOS
        frame_editar = LabelFrame(self.janela_editar_veiculo, text="Editar o seguinte veiculo",
                              font=('Calibri', 20, 'bold'))
        frame_editar.grid(row=1, column=0, columnspan=20, pady=20)
        self.cabecalho_frame = Label(frame_editar, text="Dados antigos")
        self.cabecalho_frame.grid(row=1, column=0)
        self.cabecalho_frame = Label(frame_editar, text="Dados novos")
        self.cabecalho_frame.grid(row=1, column=2)
        # Entry Matricula Antiga
        self.input_matricula_antiga = Entry(frame_editar, textvariable=StringVar(self.janela_editar_veiculo, value=matricula),
                                            state='readonly')
        self.input_matricula_antiga.grid(row=2, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=2, column=1)
        # Entry Matricula Nova
        self.input_matricula_nova = Entry(frame_editar, width=23)
        self.input_matricula_nova.grid(row=2, column=2)

        # Entry Marca Antiga
        self.input_marca_antiga = Entry(frame_editar, textvariable=StringVar(self.janela_editar_veiculo, value=marca),
                                            state='readonly')
        self.input_marca_antiga.grid(row=3, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=3, column=1)
        # Entry Marca Nova
        self.input_marca_nova = Entry(frame_editar, width=23)
        self.input_marca_nova.grid(row=3, column=2)

        # Entry Modelo Antigo
        self.input_modelo_antigo = Entry(frame_editar, textvariable=StringVar(self.janela_editar_veiculo, value=modelo),
                                            state='readonly')
        self.input_modelo_antigo.grid(row=4, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=4, column=1)
        # Entry Modelo Novo
        self.input_modelo_novo = Entry(frame_editar, width=23)
        self.input_modelo_novo.grid(row=4, column=2)

        # Entry Tipo Antigo
        self.input_tipo_antigo = Entry(frame_editar, textvariable=StringVar(self.janela_editar_veiculo, value=tipo),
                                         state='readonly')
        self.input_tipo_antigo.grid(row=5, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=5, column=1)
        # Select Tipo Novo
        self.input_tipo_novo = Combobox(frame_editar, values=self.tipos)
        self.input_tipo_novo.grid(row=5, column=2)

        # Entry Categoria Antiga
        self.input_categoria_antiga = Entry(frame_editar, textvariable=StringVar(self.janela_editar_veiculo, value=categoria),
                                       state='readonly')
        self.input_categoria_antiga.grid(row=6, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=6, column=1)
        # Select Categoria  Nova
        self.input_categoria_novo = Combobox(frame_editar, values=self.categorias)
        self.input_categoria_novo.grid(row=6, column=2)

        # Entry Transmissão Antiga
        self.input_transmissao_antiga = Entry(frame_editar, textvariable=StringVar(self.janela_editar_veiculo, value=transmissao),
                                       state='readonly')
        self.input_transmissao_antiga.grid(row=7, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=7, column=1)
        # Select Transmissao Nova
        self.input_transmissao_nova = Combobox(frame_editar, values=self.transmissoes)
        self.input_transmissao_nova.grid(row=7, column=2)

        # Entry Valor Diário Antigo
        self.input_valorDiario_antigo = Entry(frame_editar, textvariable=StringVar(self.janela_editar_veiculo, value=valor_diario),
                                              state='readonly')
        self.input_valorDiario_antigo.grid(row=8, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=8, column=1)
        # Entry Valor Diário Novo
        self.input_valorDiario_novo = Entry(frame_editar, width=23)
        self.input_valorDiario_novo.grid(row=8, column=2)

        # Entry Quantidade Pessoas Antigo
        self.input_qtdPessoas_antigo = Entry(frame_editar, textvariable=StringVar(self.janela_editar_veiculo, value=qtd_pessoas),
                                       state='readonly')
        self.input_qtdPessoas_antigo.grid(row=9, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=9, column=1)
        # Select Tipo Quantidade Pessoas Novo
        self.input_qtdPessoas_novo = Combobox(frame_editar, values=self.qtd_pessoas)
        self.input_qtdPessoas_novo.grid(row=9, column=2)

        # Entry Data Proxima Revisão Antiga
        self.input_data_revisao_antiga = Entry(frame_editar, textvariable=StringVar(self.janela_editar_veiculo, value=data_revisao),
                                             state='readonly')
        self.input_data_revisao_antiga.grid(row=10, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=10, column=1)
        # Input Data Proxima Revisão Novo
        self.input_data_revisao_nova = DateEntry(frame_editar, width=20, background='blue', foreground='white',
                                                borderwidth=2, date_pattern="dd/mm/yyyy")
        self.input_data_revisao_nova.grid(row=10, column=2)
        self.input_data_revisao_nova.delete(0, END)

        # Entry Data Ultima Inspeção Antiga
        self.input_data_inspecao_antiga = Entry(frame_editar, textvariable=StringVar(self.janela_editar_veiculo, value=data_inspecao),
                                               state='readonly')
        self.input_data_inspecao_antiga.grid(row=11, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=11, column=1)
        # Input Data Ultima Inspeção Novo
        self.input_data_inspecao_nova = DateEntry(frame_editar, width=20, background='blue', foreground='white',
                                                 borderwidth=2, date_pattern="dd/mm/yyyy")
        self.input_data_inspecao_nova.grid(row=11, column=2)
        self.input_data_inspecao_nova.delete(0, END)

        # Entry Endereço da Imagem Antigo
        self.input_endereco_imagem_antigo = Entry(frame_editar, textvariable=StringVar(self.janela_editar_veiculo, value=endereco_imagem),
                                              state='readonly')
        self.input_endereco_imagem_antigo.grid(row=12, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=12, column=1)
        # Entry Valor Diário Novo
        self.input_endereco_imagem_novo = Entry(frame_editar, width=23)
        self.input_endereco_imagem_novo.grid(row=12, column=2)

        # Mensagem Informativa
        self.mensagem_editar_veiculo = Label(frame_editar, text='', fg='red')
        self.mensagem_editar_veiculo.grid(row=13, column=0, columnspan=2, sticky=W+E)

        # Botao adicionar veiculo
        self.botao_editar_veiculo = ttk.Button(frame_editar, text="Editar Veiculo", command=self.editar_veiculo)
        self.botao_editar_veiculo.grid(row=14, column=2, sticky = W + E)

    def editar_veiculo(self):
        ''' Método que altera os valores pretendidos do veiculo em questão '''
        matricula_antiga = self.input_matricula_antiga.get()

        lista_antigos_elementos = [self.input_matricula_antiga.get(), self.input_marca_antiga.get(),
                                   self.input_modelo_antigo.get(), self.input_tipo_antigo.get(),
                                   self.input_categoria_antiga.get(), self.input_transmissao_antiga.get(),
                                   self.input_valorDiario_antigo.get(), self.input_qtdPessoas_antigo.get(),
                                   self.input_data_revisao_antiga.get(), self.input_data_inspecao_antiga.get(),
                                   self.input_endereco_imagem_antigo.get()]
        lista_novos_elementos = [self.input_matricula_nova.get(), self.input_marca_nova.get(),
                                 self.input_modelo_novo.get(), self.input_tipo_novo.get(),
                                 self.input_categoria_novo.get(), self.input_transmissao_nova.get(),
                                 self.input_valorDiario_novo.get(), self.input_qtdPessoas_novo.get(),
                                   self.input_data_revisao_nova.get(), self.input_data_inspecao_nova.get(),
                                   self.input_endereco_imagem_novo.get()]
        self.mensagem_editar_veiculo['text'] = ''

        # Verifica se os campos estão vazios:
        if self.verifica_lista_vazia(lista_novos_elementos):
            self.mensagem_editar_veiculo['text'] = 'Preencha pelo menos um campo caso queira alterar o veiculo !!'
            return

        # Atualizar nova lista
        lista_atualizada = ()
        for i, elemento_novo in enumerate(lista_novos_elementos):
            if len(elemento_novo) == 0: lista_atualizada += (lista_antigos_elementos[i],)
            else: lista_atualizada += (elemento_novo,)

        # Verifica se é moto e SUV simultaneamentw
        if not self.validar_dados_tipo(lista_atualizada[3], lista_atualizada[4], lista_atualizada[7]):
            self.mensagem_editar_veiculo['text'] = 'Um Moto não pode ser SUV ou ter mais de 2 lugares !!'
            return

        # Executar a consulta para obter o ID baseado na matricula)
        with sqlite3.connect(self.db) as con:
            cursor = con.cursor()
            cursor.execute("SELECT id FROM veiculos WHERE matricula = ?", (lista_atualizada[0],))
            resultado = cursor.fetchone()
            id_desejado = resultado[0]

        # Executa a alteração dos dados pretendidos
        lista_atualizada += (id_desejado,)
        query = 'UPDATE veiculos SET matricula=?, marca=?, modelo=?, tipo=?, categoria=?, transmissao=?, valor_diario=?, qtd_pessoas=?, data_proxima_revisao=?, data_ult_inspecao=?, endereco_imagem=? WHERE id = ?'
        self.db_consulta(query, lista_atualizada)  # Executar a consulta
        self.janela_editar_veiculo.destroy()  # Fechar a janela de edição de veiculos
        self.mensagem_veiculo['text'] = 'O veiculo com matricula {} foi atualizado com êxito'.format(matricula_antiga)  # Mostrar mensagem para o utilizador
        self.get_veiculos()  # Atualizar a tabela de veiculos

    def eliminar_veiculo(self):
        '''Elimina veiculo através da matricula'''
        self.mensagem_veiculo['text'] = ''
        # Verifica se foi selecionado um veiculo
        try:
            self.tabela_veiculos.item(self.tabela_veiculos.selection())['text'][0]
        except IndexError as e:
            self.mensagem_veiculo['text'] = "Por favor, selecione um veiculo"
            return

        # Busca da matricula na linha da tabela selecionada
        matricula = self.tabela_veiculos.item(self.tabela_veiculos.selection())['text']
        # Consulta SQL para Eliminar veiculo atrvés da matricula
        query = 'DELETE FROM veiculos WHERE matricula = ?'
        self.db_consulta(query, (matricula,))  # Executar a consulta
        self.mensagem_veiculo['text'] = 'O veiculo com a matricula: {} foi eliminado com êxito'.format(matricula)
        self.get_veiculos()  # Atualizar a tabela de produtos

    def get_veiculos(self):
        ''' Insere na tabela todos os veiculos criados'''
        # Limpar tabela
        registos_tabela = self.tabela_veiculos.get_children()
        for linha in registos_tabela:
            self.tabela_veiculos.delete(linha)

        # Consulta SQL
        query = 'SELECT * FROM veiculos'
        veiculos = self.db_consulta(query)

        # Escrever dados na tabela
        for linha in veiculos:
            ocupacao = "Livre" if linha[9] == 0 else "Ocupado"
            self.tabela_veiculos.insert('', 0, text= linha[1], values = (linha[2], linha[3], linha[4],
                                                                         linha[5], linha[6], linha[7], linha[8], ocupacao,
                                                                         linha[10], linha[11], linha[12]))
        # Associar evento de clique.
        self.tabela_veiculos.bind("<Button-1>", self.identificar_imagem)

    def identificar_imagem(self, event):
        '''Identifica a celula pressionada, identifica o caminho e abre a imagem'''
        # Identificar a célula clicada
        item_id = self.tabela_veiculos.identify_row(event.y)
        coluna_id = self.tabela_veiculos.identify_column(event.x)

        if item_id and coluna_id == "#11":  # Verificar se clica na última coluna
            valores = self.tabela_veiculos.item(item_id, "values")
            caminho = valores[10]  # Matricula da pessoa na linha
            os.startfile(caminho)  # Abre o arquivo com o aplicativo padrão do sistema


    def adicionar_veiculo(self):
        '''Verifica as devidas condições e adiciona o veiculo'''
        self.mensagem_adicionar_veiculo['text'] = ""
        lista_elementos = [self.entry_matricula.get(), self.entry_marca.get(), self.entry_modelo.get(),
                           self.select_categoria.get(), self.select_transmissao.get(), self.select_tipo.get(),
                           self.entry_valor_diario.get(), self.select_qtd_pessoas.get(), self.input_data_proxima_revisao.get(),
                           self.input_data_ult_inspecao.get()]

        # Verifica se todos os campos foram preenchidos
        if not self.dados_nao_vazios(lista_elementos):
            self.mensagem_adicionar_veiculo['text'] = "Preencha todo o formulário !!"
            return
        # Verifica se matricula já existe na DB
        elif self.verifica_matricula(self.entry_matricula.get()):
            self.mensagem_adicionar_veiculo['text'] = "A matricula inserida já existe no sistema"
            return
        # Verifica se um moto é um SUV ou tem 5 ou mais lugares2
        elif not self.validar_dados_tipo(self.select_tipo.get(),self.select_categoria.get(), self.select_qtd_pessoas.get()):
            self.mensagem_adicionar_veiculo['text'] = "Um Moto não pode ser SUV ou ter mais de 4 lugares !!"
            return
        # Verific se o Valor Diario inserido é um número
        try:
            float(self.entry_valor_diario.get())
        except ValueError as e:
            self.mensagem_adicionar_veiculo['text'] = "O Valor Diário inserido não é válido ..."
            return
        # Verifica se foi submetido foto do veiculo
        if len(self.caminho_destino_arquivo) == 0:
            self.mensagem_adicionar_veiculo['text'] = "É necessário submeter uma foto do veiculo"
            return

        # Copia o arquivo para o diretório de destino
        shutil.copy(self.arquivo_arrastado, self.caminho_destino_arquivo)
        print(f"Arquivo copiado para {self.caminho_destino_arquivo}")
        print("Nome Arquivo: " + self.nome_arquivo)

        # Consulta ou insere o novo veiculo na DB
        query = 'INSERT INTO veiculos VALUES(NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
        parametros = (self.entry_matricula.get(), self.entry_marca.get(), self.entry_modelo.get(), self.select_tipo.get(), self.select_categoria.get(), self.select_transmissao.get(),
                    self.entry_valor_diario.get(), self.select_qtd_pessoas.get(), False, self.input_data_proxima_revisao.get(),
                    self.input_data_ult_inspecao.get(), self.caminho_destino_arquivo)

        self.db_consulta(query, parametros)
        self.mensagem_adicionar_veiculo['text'] = "Dados guardados com sucesso !!"

        # Limpar dados do formulario quando o veiculo inserido com sucesso
        self.entry_matricula.delete(0, END)
        self.entry_marca.delete(0, END)
        self.entry_modelo.delete(0, END)
        self.select_tipo.delete(0, END)
        self.select_categoria.delete(0, END)
        self.select_transmissao.delete(0, END)
        self.select_qtd_pessoas.delete(0, END)
        self.entry_valor_diario.delete(0, END)
        self.input_data_proxima_revisao.delete(0, END)
        self.input_data_ult_inspecao.delete(0, END)
        self.caminho_destino_arquivo = ''
        self.area_arraste.config(text="Arraste uma foto do veiculo para aqui !!")

        self.get_veiculos()

    def manutencao_veiculo(self):
        ''' Faz com que veiculo fique Ocupado por tempo indeterminado'''
        self.mensagem_veiculo['text'] = ''

        # Verifica se algum veiculo foi selecionado
        try:
            self.tabela_veiculos.item(self.tabela_veiculos.selection())['text'][0]
        except IndexError as e:
            self.mensagem_veiculo['text'] = 'Por favor, selecione um veiculo'
            return

        # Verifica se o veiculo selecionado está Livre ou Ocupado
        if self.tabela_veiculos.item(self.tabela_veiculos.selection())['values'][7] == "Ocupado":
            self.mensagem_veiculo['text'] = 'O veiculo selecionado está neste momento ocupado. Aguarde até que este fique livre.'
            return

        # Query que coloca veiculo como Ocupado
        query = 'UPDATE veiculos SET ocupado=1 WHERE matricula = ?'
        parametros = (self.tabela_veiculos.item(self.tabela_veiculos.selection())['text'],)
        self.db_consulta(query, parametros)  # Executar a consulta

        self.get_veiculos()

        self.mensagem_veiculo['text'] = 'O veiculo com a matricula {} está em manutenção.'.format(
                                        self.tabela_veiculos.item(self.tabela_veiculos.selection())['text'])

    def exportar_veiculos(self):
        """Criar um ficheiro Excel com os Veiculos"""
        # Consulta SQL de todos os Veiculos
        query = 'SELECT * FROM veiculos'
        veiculos = self.db_consulta(query)
        dados = []
        for veiculo in veiculos:
            dados.append(list(veiculo)[1:])

        columns = ["Matricula", "Marca", "Modelo", "Tipo", "Categoria", "Transmissão", "Valor Diário", "Número Pessoas",
                   "Ocupação", "Data Proxima Revisão", "Data Última Inspeção", "Endereço da imagem"]
        df = pd.DataFrame(dados, columns=columns)

        # Obtendo a data de hoje no formato desejado
        data_hoje = datetime.now().strftime('%d-%m-%Y')  # Formato: DD-MM-YYYY+
        # Definindo o nome do arquivo com a data de hoje
        nome_arquivo = f'relatorio_veiculos_{data_hoje}.xlsx'
        # Exportando o DataFrame para um arquivo Excel
        df.to_excel(nome_arquivo, index=False)
        print(f'Arquivo "{nome_arquivo}" exportado com sucesso!')
        self.mensagem_veiculo['text'] = "Os dados foram exportados com sucesso !!"

    def processar_arquivo_arrastado(self, event):
        matricula = self.entry_matricula.get()

        # Diretório de destino onde o arquivo será armazenado
        caminho_destino = r"C:\Users\SBUtilizador\PycharmProjects\ProjetoFinal\ProjetoFinal\recursos\fotos_veiculos"  # Substitua pelo seu caminho de destino

        # Verifica se foi inserida matricula. É necessária para guardar a imgem com o nome da matricula
        if len(matricula) == 0:
            self.mensagem_adicionar_veiculo['text'] = "Preencha todos os dados !!"
            return

        # Obtém o caminho do arquivo arrastado
        self.arquivo_arrastado = event.data

        if os.path.isfile(self.arquivo_arrastado):
            self.nome_arquivo = os.path.basename(self.arquivo_arrastado)
            formato_arquivo = self.nome_arquivo.split(".")
            # Verifica se tem um formato de imagem
            if formato_arquivo[1].lower() in ['jpeg', 'png', 'bmp', 'gif', 'tiff', 'jpg']:
                nome_arquivo = os.path.basename(self.arquivo_arrastado)
                # Cria o caminho a guardar as imagens: caminho destino + matricula + formato do arquivo
                self.caminho_destino_arquivo = os.path.join(caminho_destino, matricula+"."+formato_arquivo[1])
                self.area_arraste.config(text=f"Arquivo {nome_arquivo} copiado com sucesso!")
            else:
                self.area_arraste.config(text=f"O formato do arquivo não é válido.\n Verifique o formato e tente novamente.")
        else:
            self.area_arraste.config(text=f"Não foi possível copiar o arquivo.\n Verifique se é um arquivo válido e tente novamente.")


    # ---------------------------------------------------
    # VALIDAÇÃO DE DADOS FORMULÁRIO

    def dados_nao_vazios(self, lista_elementos):
        '''Retorna True caso algum elemento esteja vazio'''
        for elemento in lista_elementos:
            if len(elemento) == 0: return False
        return True

    def verifica_matricula(self, matricula):
        '''Retorna True se já existir a matricula na Base de Dados'''
        # Consulta SQL
        query = 'SELECT matricula FROM veiculos'
        matriculas = self.db_consulta(query)
        lista_matriculas = []
        for matricula_ in matriculas:
            lista_matriculas.append(matricula_[0])
        if matricula in lista_matriculas: return True
        else: return False

    def verifica_lista_vazia(self, lista_elementos):
        '''Retorna True se lista estiver vazia'''
        for elemento in lista_elementos:
            if len(elemento) != 0: return False
        return True

    def validar_dados_tipo(self, tipo, categoria, qtd_pessoas):
        '''Retorna True se for validado'''
        if tipo == "Mota":
            if categoria == "SUV" or qtd_pessoas != "1-4":
                return False
            return True
        return True
    # --------------------------------------------------------------
# ------------------------------  FIM PAGINA VEICULOS  --------------------------------


# ------------------------------- INCIO PAGINA CLIENTES ---------------------------------

    def pagina_clientes(self):
        # Limpar pagina
        self.limpar_pagina()

        # Titulo da página
        self.titulo_pagina_clientes = Label(text="Página Clientes", font=('Calibri', 13, 'bold'))
        self.titulo_pagina_clientes.grid(row=0, column=3)

        # Tabela com Clientes
        style = ttk.Style()
        style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Calibri', 11))
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 13, 'bold'))
        style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])

        # Estrutura da tabela
        self.tabela_clientes = ttk.Treeview(self.janela, height=12, columns=('col1', 'col2', 'col3'), style="mystyle.Treeview")
        self.tabela_clientes.grid(row=0, column=0, columnspan=4)
        # Cabeçalhos
        self.tabela_clientes.heading('#0', text='NIF', anchor=CENTER)
        self.tabela_clientes.heading('#1', text='Nome', anchor=CENTER)
        self.tabela_clientes.heading('#2', text='Idade', anchor=CENTER)
        self.tabela_clientes.heading('#3', text='Username', anchor=CENTER)
        # Configuração da largura das colunas
        self.tabela_clientes.column('#0', width=100, anchor=CENTER)
        self.tabela_clientes.column('col1', width=120, anchor=CENTER)
        self.tabela_clientes.column('col2', width=80, anchor=CENTER)
        self.tabela_clientes.column('col3', width=80, anchor=CENTER)
        self.get_clientes()

        # Mensagem Informativa
        self.mensagem_cliente = Label(self.janela, text='', fg='red')
        self.mensagem_cliente.grid(row=2, column=0, columnspan=2, sticky=W + E)

        # Botão adicionar cliente:
        self.botao_pagina_adicionar_cliente= ttk.Button(self.janela, text="Adicionar Cliente",
                                                         command=self.pagina_adicionar_cliente)
        self.botao_pagina_adicionar_cliente.grid(row=3, column=0, sticky=W + E)

        # Botao Editar cliente
        self.botao_pagina_editar_cliente = ttk.Button(self.janela, text="Editar Cliente", command=self.pagina_editar_cliente)
        self.botao_pagina_editar_cliente.grid(row=3, column=1, sticky=W + E)

        # Botao Eliminar cliente
        self.botao_eliminar_cliente = ttk.Button(self.janela, text="Eliminar Cliente", command=self.eliminar_cliente)
        self.botao_eliminar_cliente.grid(row=3, column=2, sticky=W + E)

        # Botao Exportar Clientes
        self.botao_exportar_clientes = ttk.Button(self.janela, text="Exportar Clientes", command=self.exportar_clientes)
        self.botao_exportar_clientes.grid(row=3, column=3, sticky=W + E)

    def pagina_adicionar_cliente(self):
        '''Pagina para adicionar clientes'''
        self.mensagem_cliente['text'] = ''

        # Criar nova janela
        self.janela_adicionar_cliente = Toplevel()
        self.janela_adicionar_cliente.title = "Adicionar Cliente"
        self.janela_adicionar_cliente.resizable(1, 1)
        self.janela_adicionar_cliente.wm_iconbitmap('recursos/car_icon.ico')

        # Criação da frame Informação Cliente
        frame_info_cliente = LabelFrame(self.janela_adicionar_cliente, text="Informações do cliente",
                                        font=('Calibri', 16, 'bold'))
        frame_info_cliente.grid(row=0, column=0, columnspan=3, pady=20)

        # Label NIF do cliente
        self.etiqueta_nif_cliente = Label(frame_info_cliente, text="NIF: ")
        self.etiqueta_nif_cliente.grid(row=1, column=0)
        self.input_nif_cliente = Entry(frame_info_cliente)
        self.input_nif_cliente.grid(row=1, column=1)

        # Label nome do cliente
        self.etiqueta_nome_cliente = Label(frame_info_cliente, text="Nome: ")
        self.etiqueta_nome_cliente.grid(row=2, column=0)
        self.input_nome_cliente = Entry(frame_info_cliente)
        self.input_nome_cliente.grid(row=2, column=1)

        # Label idade do cliente
        self.etiqueta_idade_cliente = Label(frame_info_cliente, text="Idade: ")
        self.etiqueta_idade_cliente.grid(row=3, column=0)
        self.input_idade_cliente = Entry(frame_info_cliente)
        self.input_idade_cliente.grid(row=3, column=1)

        # Label idade do cliente
        self.etiqueta_username_cliente = Label(frame_info_cliente, text="Username: ")
        self.etiqueta_username_cliente.grid(row=4, column=0)
        self.input_username_cliente = Entry(frame_info_cliente)
        self.input_username_cliente.grid(row=4, column=1)

        # Label idade do cliente
        self.etiqueta_password_cliente = Label(frame_info_cliente, text="Password: ")
        self.etiqueta_password_cliente.grid(row=5, column=0)
        self.input_password_cliente = Entry(frame_info_cliente)
        self.input_password_cliente.grid(row=5, column=1)

        # Mensagem Informativa
        self.mensagem_adicionar_cliente = Label(frame_info_cliente, text='', fg='red')
        self.mensagem_adicionar_cliente.grid(row=6, column=0, columnspan=2, sticky=W + E)

        # Botao adicionar cliente
        self.botao_adicionar_cliente = ttk.Button(frame_info_cliente, text="Adicionar Cliente",
                                                  command=self.adicionar_cliente)
        self.botao_adicionar_cliente.grid(row=7, columnspan=2, sticky=W + E)

    def pagina_editar_cliente(self):
        '''Pagina para editar clientes'''
        self.mensagem_cliente['text'] = ''

        # Verifica se tem algum cliente selecionado
        try:
            self.tabela_clientes.item(self.tabela_clientes.selection())['values'][0]
        except IndexError as e:
            self.mensagem_cliente['text'] = 'Por favor, selecione um cliente'
            return

        nif = self.tabela_clientes.item(self.tabela_clientes.selection())['text']
        nome = self.tabela_clientes.item(self.tabela_clientes.selection())['values'][0]
        idade = self.tabela_clientes.item(self.tabela_clientes.selection())['values'][1]
        username = self.tabela_clientes.item(self.tabela_clientes.selection())['values'][2]

        # Consulta da password através do nif
        query = 'SELECT password FROM clientes WHERE nif=?'
        parametros = (nif,)
        password = self.db_consulta(query, parametros)

        # Criar página Editar
        self.janela_editar_cliente = Toplevel()
        self.janela_editar_cliente.title = "Editar Cliente"
        self.janela_editar_cliente.resizable(1, 1)
        self.janela_editar_cliente.wm_iconbitmap('recursos/car_icon.ico')

        titulo = Label(self.janela_editar_cliente, text="Edição de Clientes", font=('Calibri', 30, 'bold'), fg='blue')
        titulo.grid(row=0, column=0)

        # CRIAR FRAME EDITAR CLIENTES
        frame_editar = LabelFrame(self.janela_editar_cliente, text="Editar o seguinte cliente",
                                  font=('Calibri', 20, 'bold'))
        frame_editar.grid(row=1, column=0, columnspan=20, pady=20)
        self.cabecalho_frame = Label(frame_editar, text="Dados antigos")
        self.cabecalho_frame.grid(row=1, column=0)
        self.cabecalho_frame = Label(frame_editar, text="Dados novos")
        self.cabecalho_frame.grid(row=1, column=2)

        # Entry NIF Antigo
        self.input_nif_antigo = Entry(frame_editar, textvariable=StringVar(self.janela_editar_cliente, value=nif),
                                            state='readonly')
        self.input_nif_antigo.grid(row=2, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=2, column=1)
        # Entry NIF Novo
        self.input_nif_novo = Entry(frame_editar)
        self.input_nif_novo.grid(row=2, column=2)

        # Entry Nome Antigo
        self.input_nome_antigo = Entry(frame_editar, textvariable=StringVar(self.janela_editar_cliente, value=nome),
                                        state='readonly')
        self.input_nome_antigo.grid(row=3, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=3, column=1)
        # Entry Nome Novo
        self.input_nome_novo = Entry(frame_editar)
        self.input_nome_novo.grid(row=3, column=2)

        # Entry Idade Antiga
        self.input_idade_antiga = Entry(frame_editar, textvariable=StringVar(self.janela_editar_cliente, value=idade),
                                         state='readonly')
        self.input_idade_antiga.grid(row=4, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=4, column=1)
        # Entry Idade Nova
        self.input_idade_nova = Entry(frame_editar)
        self.input_idade_nova.grid(row=4, column=2)

        # Entry Username Antigo
        self.input_username_antigo = Entry(frame_editar, textvariable=StringVar(self.janela_editar_cliente, value=username),
                                        state='readonly')
        self.input_username_antigo.grid(row=5, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=5, column=1)
        # Entry Username Novo
        self.input_username_novo = Entry(frame_editar)
        self.input_username_novo.grid(row=5, column=2)

        # Entry Password Antiga
        self.input_password_antiga = Entry(frame_editar, textvariable=StringVar(self.janela_editar_cliente, value=password),
                                        state='readonly')
        self.input_password_antiga.grid(row=6, column=0)
        # Seta
        self.seta = Label(frame_editar, text="  ->")
        self.seta.grid(row=6, column=1)
        # Entry Password Nova
        self.input_password_nova = Entry(frame_editar)
        self.input_password_nova.grid(row=6, column=2)

        # Mensagem Informativa
        self.mensagem_editar_cliente = Label(frame_editar, text='', fg='red')
        self.mensagem_editar_cliente.grid(row=7, column=0, columnspan=5, sticky=W + E)

        # Botao adicionar veiculo
        self.botao_editar_cliente = ttk.Button(frame_editar, text="Editar Cliente", command=self.editar_cliente)
        self.botao_editar_cliente.grid(row=8, column=2, sticky=W + E)

    def adicionar_cliente(self):
        '''Depois de devidas verificações insere cliente na DB'''
        lista_elementos = [self.input_nif_cliente.get(), self.input_nome_cliente.get(), self.input_idade_cliente.get(),
                           self.input_username_cliente.get(), self.input_password_cliente.get()]

        # Verifica se todos os elementos foram introduzidos
        if not self.dados_nao_vazios(lista_elementos):
            self.mensagem_adicionar_cliente['text'] = "Preencha todo o formolário !!"
            return

        # Verifica se os dados introduzidos são números (inteiros)
        try:
            int(self.input_nif_cliente.get())
        except ValueError as e:
            self.mensagem_adicionar_cliente['text'] = "O NIF inserido não é válido ..."
            return

        # Verifica se o NIF inserido já existe na DB
        if self.verifica_nif_existente(self.input_nif_cliente.get()):
            self.mensagem_adicionar_cliente['text'] = "O NIF inserido já existe na Base de Dados !!"
            return

        # Verifica se os dados introduzidos são números (inteiros)
        try:
            int(self.input_idade_cliente.get())
        except ValueError as e:
            self.mensagem_adicionar_cliente['text'] = "A idade inserida não é válida ..."
            return

        # Verifica se o cliente é maior de 18 anos
        if int(self.input_idade_cliente.get()) < 18:
            self.mensagem_adicionar_cliente['text'] = "Só maiores de 18 anos se podem inscrever"
            return

        # Verificar se Username já existe na DB
        if self.verifica_username_existente(self.input_username_cliente.get()):
            self.mensagem_adicionar_cliente['text'] = "O username inserido já existe no sistema."
            return

        # Query que insere cliente na DB
        query = 'INSERT INTO clientes VALUES(NULL, ?, ?, ?, ?, ?)'
        parametros = (self.input_nif_cliente.get(), self.input_nome_cliente.get(), self.input_idade_cliente.get(),
                      self.input_username_cliente.get(), self.input_password_cliente.get())
        self.db_consulta(query, parametros)
        self.mensagem_adicionar_cliente['text'] = "Cliente adicionado com sucesso !!"

        # Limpar dados do formulario quando o veiculo inserido com sucesso
        self.input_nif_cliente.delete(0, END)
        self.input_nome_cliente.delete(0, END)
        self.input_idade_cliente.delete(0, END)
        self.input_username_cliente.delete(0, END)
        self.input_password_cliente.delete(0, END)

        self.get_clientes()

    def editar_cliente(self):
        '''Depois das devidas verificações edita os dados do cliente na DB'''
        nif_antigo = self.input_nif_antigo.get()
        lista_antigos_elementos = [self.input_nif_antigo.get(), self.input_nome_antigo.get(), self.input_idade_antiga.get(),
                                   self.input_username_antigo.get(), self.input_password_antiga.get()]
        lista_novos_elementos = [self.input_nif_novo.get(), self.input_nome_novo.get(), self.input_idade_nova.get(),
                                 self.input_username_novo.get(), self.input_password_nova.get()]
        self.mensagem_editar_cliente['text'] = ''

        # Verifica se há alguma alteração:
        if self.verifica_lista_vazia(lista_novos_elementos):
            self.mensagem_editar_cliente['text'] = 'Preencha pelo menos um campo caso queira fazer a alteração !!'
            return

        #Verifica se o NIF é um numero e se já existe na DB
        if len(self.input_nif_novo.get()) != 0:
            try:
                if self.verifica_nif_existente(int(self.input_nif_novo.get())):
                    self.mensagem_editar_cliente['text'] = "O NIF inserido já existe na Base de Dados !!"
                    return
            except ValueError as e:
                self.mensagem_editar_cliente['text'] = "O NIF inserido não é válido."
                return

        # Verifica se idade é um numero e se é maior do que 18 anos
        if len(self.input_idade_nova.get()) != 0:
            try:
                if int(self.input_idade_nova.get()) < 18:
                    self.mensagem_editar_cliente['text'] = "Só maiores de 18 anos se podem inscrever"
                    return
            except ValueError as e:
                self.mensagem_editar_cliente['text'] = "A idade inserida não é válida."
                return

        # Verifica se o Username já existe na DB
        if  len(self.input_username_novo.get()) != 0:
            if self.verifica_username_existente(self.input_username_novo.get()):
                self.mensagem_editar_cliente['text'] = 'O Username inserido já existe no sistema.'
                return

        # Atualizar nova lista
        lista_atualizada = ()
        for i, elemento_novo in enumerate(lista_novos_elementos):
            if len(elemento_novo) == 0:
                lista_atualizada += (lista_antigos_elementos[i],)
            else:
                lista_atualizada += (elemento_novo,)

        # Executar a consulta para obter o ID através do NIF)
        with sqlite3.connect(self.db) as con:
            cursor = con.cursor()
            cursor.execute("SELECT id FROM clientes WHERE nif = ?", (nif_antigo,))
            resultado = cursor.fetchone()
            id_desejado = resultado[0]

        lista_atualizada += (id_desejado,)
        # Query que edita os dados do cliente
        query = 'UPDATE clientes SET nif=?, nome=?, idade=?, username=?, password=? WHERE id = ?'
        self.db_consulta(query, lista_atualizada)  # Executar a consulta
        self.janela_editar_cliente.destroy()  # Fechar a janela de edição de clientes
        self.mensagem_cliente['text'] = 'O cliente com o NIF {} foi atualizado com êxito'.format(nif_antigo)  # Mostrar mensagem para o utilizador
        self.get_clientes()  # Atualizar a tabela de clientes

    def eliminar_cliente(self):
        """Elimina cliente através do NIF"""
        self.mensagem_cliente['text'] = ''
        # Verifica se algum cliente foi selecionado
        try:
            self.tabela_clientes.item(self.tabela_clientes.selection())['values'][0]
        except IndexError as e:
            self.mensagem_veiculo['text'] = "Por favor, selecione um cliente"
            return

        nif = self.tabela_clientes.item(self.tabela_clientes.selection())['text']
        # Query para apagar cliente através do NIF
        query = 'DELETE FROM clientes WHERE nif = ?'  # Consulta SQL
        self.db_consulta(query, (nif,))  # Executar a consulta
        self.mensagem_cliente['text'] = 'O cliente com o NIF: {} foi eliminado com êxito'.format(nif)
        self.get_clientes()  # Atualizar a tabela de clientes

    def exportar_clientes(self):
        """Criar um ficheiro Excel com os Clientes"""
        # Consulta SQL de todos os clientes
        query = 'SELECT * FROM clientes'
        clientes = self.db_consulta(query)
        dados = []
        for cliente in clientes:
            dados.append(list(cliente)[1:-1])

        columns = ["NIF", "Nome", "Idade", "Username"]
        df = pd.DataFrame(dados, columns=columns)

        # Obtendo a data de hoje no formato desejado
        data_hoje = datetime.now().strftime('%d-%m-%Y')  # Formato: DD-MM-YYYY+
        # Definindo o nome do arquivo com a data de hoje
        nome_arquivo = f'relatorio_clientes_{data_hoje}.xlsx'
        # Exportando o DataFrame para um arquivo Excel
        df.to_excel(nome_arquivo, index=False)
        self.mensagem_cliente['text'] = "Os dados foram exportados com sucesso !!"

    def get_clientes(self):
        '''Insere os dados de todos os clientes na tabela'''
        # Limpar tabela
        registos_tabela = self.tabela_clientes.get_children()
        for linha in registos_tabela:
            self.tabela_clientes.delete(linha)

        # Consulta SQL de todos os clientes
        query = 'SELECT * FROM clientes'
        clientes = self.db_consulta(query)

        # Escrever dados na tabela
        for linha in clientes:
            self.tabela_clientes.insert('', 0, text= linha[1], values = (linha[2], linha[3], linha[4]))

    def verifica_nif_existente(self, nif):
        '''Retorna True se já existir o NIF na Base de Dados'''
        # Consulta SQL dos NIFs de todos os clientes
        query = 'SELECT nif FROM clientes'
        nifs = self.db_consulta(query)
        lista_nifs = []
        lista_nifs = [nif_[0] for nif_ in nifs]

        if int(nif) in lista_nifs: return True
        else:  return False

    def verifica_username_existente(self, username):
        '''Retorna True se já existir o USERNAME na Base de Dados'''
        # Consulta SQL dos usernames de todos os clientes
        query = 'SELECT username FROM clientes'
        usernames = self.db_consulta(query)
        lista_users = []
        lista_users = [user[0] for user in usernames]

        if username in lista_users: return True
        else: return False  # CONFIRMAR FUNÇAO !!!

# --------------------------- FIM PAGINA CLIENTES ---------------------------------

# -------------------------- INICIO PAGINA RESERVAS ----------------------------

    def pagina_reservas(self):
        # Limpar pagina
        self.limpar_pagina()

        # Titulo da página
        self.titulo_pagina_reservas = Label(text="Página Reservas", fg="blue", font=('Calibri', 16, 'bold'))
        self.titulo_pagina_reservas.grid(row=0, column=1)

        # Tabela com Reservas
        style = ttk.Style()
        style.configure("mystyle.Treeview", highlightthickness=1, bd=1, font=('Calibri', 11))
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 13, 'bold'))
        style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])

        # Estrutura da tabela
        self.tabela_reservas = ttk.Treeview(self.janela, height=10, columns=('col1', 'col2', 'col3', 'col4', 'col5', 'col6'),
                                            style="mystyle.Treeview")
        self.tabela_reservas.grid(row=1, column=0, columnspan=5)
        # Cabeçalhos
        self.tabela_reservas.heading('#0', text='Matricula do Veiculo', anchor=CENTER)
        self.tabela_reservas.heading('#1', text='NIF do Cliente', anchor=CENTER)
        self.tabela_reservas.heading('#2', text='Data de Inicio da Reserva', anchor=CENTER)
        self.tabela_reservas.heading('#3', text='Data de Término da Reserva', anchor=CENTER)
        self.tabela_reservas.heading('#4', text='Custo Total (€)', anchor=CENTER)
        self.tabela_reservas.heading('#5', text='Dias Restantes', anchor=CENTER)
        self.tabela_reservas.heading('#6', text='Estado', anchor=CENTER)

        # Configuração da largura das colunas
        self.tabela_reservas.column('#0', width=80, anchor=CENTER)
        self.tabela_reservas.column('col1', width=110, anchor=CENTER)
        self.tabela_reservas.column('col2', width=120, anchor=CENTER)
        self.tabela_reservas.column('col3', width=120, anchor=CENTER)
        self.tabela_reservas.column('col4', width=120, anchor=CENTER)
        self.tabela_reservas.column('col5', width=110, anchor=CENTER)
        self.tabela_reservas.column('col6', width=110, anchor=CENTER)
        self.get_reservas()

        # Mensagem Informativa
        self.mensagem_reservas = Label(self.janela, text='', fg='red')
        self.mensagem_reservas.grid(row=2, column=0, columnspan=2, sticky=W + E)

        # Botão adicionar reserva:
        self.botao_pagina_adicionar_reserva = ttk.Button(self.janela, text="Adicionar Reserva",
                                                         command=self.pagina_adicionar_reserva)
        self.botao_pagina_adicionar_reserva.grid(row=3, column=0, sticky=W + E)

        # Botão cancelar reserva:
        self.botao_pagina_cancelar_reserva = ttk.Button(self.janela, text="Cancelar Reserva",
                                                        command=self.cancelar_reserva)
        self.botao_pagina_cancelar_reserva.grid(row=3, column=1, sticky=W + E)

        # Botao Exportar Reservas
        self.botao_exportar_reservas = ttk.Button(self.janela, text="Exportar Reservas", command=self.exportar_reservas)
        self.botao_exportar_reservas.grid(row=3, column=2, sticky=W + E)

    def pagina_adicionar_reserva(self):
        self.mensagem_reservas['text'] = ''

        self.janela_adicionar_reserva = Toplevel()
        self.janela_adicionar_reserva.title = "Reservar Veiculo"
        self.janela_adicionar_reserva.resizable(1, 1)
        self.janela_adicionar_reserva.wm_iconbitmap('recursos/car_icon.ico')
        self.janela_adicionar_reserva.geometry('820x330')

        # Tabela de Veiculos por ocupação
        # Estrutura da tabela
        self.tabela_reservar_veiculo = ttk.Treeview(self.janela_adicionar_reserva, height=5,
                                                    columns=('col1', 'col2', 'col3', 'col4', 'col5'), style="mystyle.Treeview")
        self.tabela_reservar_veiculo.grid(row=0, column=0, columnspan=5)
        # Cabeçalhos
        self.tabela_reservar_veiculo.heading('#0', text='Matricula', anchor=CENTER)
        self.tabela_reservar_veiculo.heading('#1', text='Marca', anchor=CENTER)
        self.tabela_reservar_veiculo.heading('#2', text='Modelo', anchor=CENTER)
        self.tabela_reservar_veiculo.heading('#3', text='Ocupação', anchor=CENTER)
        self.tabela_reservar_veiculo.heading('#4', text='Data próxima revisão', anchor=CENTER)
        self.tabela_reservar_veiculo.heading('#5', text='Data última inspeção', anchor=CENTER)

        # Configuração da largura das colunas
        self.tabela_reservar_veiculo.column('#0', width=100, anchor=CENTER)
        self.tabela_reservar_veiculo.column('col1', width=100, anchor=CENTER)
        self.tabela_reservar_veiculo.column('col2', width=100, anchor=CENTER)
        self.tabela_reservar_veiculo.column('col3', width=100, anchor=CENTER)
        self.tabela_reservar_veiculo.column('col4', width=150, anchor=CENTER)
        self.tabela_reservar_veiculo.column('col5', width=150, anchor=CENTER)
        self.get_veiculos_por_ocupacao()

        # Criação da frame Reservar Veiculo
        frame_datas_reserva_veiculo = LabelFrame(self.janela_adicionar_reserva, text="Datas, Cliente e Metodo de Pagamento da reserva de veiculo",
                                        font=('Calibri', 16, 'bold'))
        frame_datas_reserva_veiculo.grid(row=1, column=0, columnspan=3, pady=20)

        # Label Data do Inicio da Reserva
        self.etiqueta_data_inicio_reserva = Label(frame_datas_reserva_veiculo, text="Data de início da reserva: ")
        self.etiqueta_data_inicio_reserva.grid(row=0, column=0)
        self.input_data_inicio_reserva = DateEntry(frame_datas_reserva_veiculo, width=20, background='blue',
                                                   foreground='white', borderwidth=2, date_pattern="dd/mm/yyyy")
        self.input_data_inicio_reserva.grid(row=0, column=1)

        # Label Data do Fim da Reserva
        self.etiqueta_data_fim_reserva = Label(frame_datas_reserva_veiculo, text="Data de fim da reserva: ")
        self.etiqueta_data_fim_reserva.grid(row=1, column=0)
        self.input_data_fim_reserva = DateEntry(frame_datas_reserva_veiculo, width=20, background='blue',
                                                   foreground='white', borderwidth=2, date_pattern="dd/mm/yyyy")
        self.input_data_fim_reserva.grid(row=1, column=1)
        self.input_data_fim_reserva.delete(0, END)

        # Label Clientes
        self.etiqueta_reservar_cliente = Label(frame_datas_reserva_veiculo, anchor='w', text='Cliente: ')
        self.etiqueta_reservar_cliente.grid(row=0, column=2)
        self.input_reservar_cliente = Combobox(frame_datas_reserva_veiculo, values=self.get_nome_clientes())  # Caixa de texto (input de texto) localizada no frame
        self.input_reservar_cliente.grid(row=0, column=3)

        # Label Metodo de Pagamento
        self.teiqueta_metodo_pagamento = Label(frame_datas_reserva_veiculo, text="Metodo de Pagamento: ")
        self.teiqueta_metodo_pagamento.grid(row=1, column=2)
        self.input_metodo_pagamento = Combobox(frame_datas_reserva_veiculo, values=self.metodos_pagamento)
        self.input_metodo_pagamento.grid(row=1, column=3)

        # Mensagem Informativa
        self.mensagem_reservar_veiculo = Label(self.janela_adicionar_reserva, text='', fg='red')
        self.mensagem_reservar_veiculo.grid(row=2, column=0, columnspan=2, sticky=W + E)

        # Botão adicionar reserva:
        self.botao_adicionar_reserva = ttk.Button(self.janela_adicionar_reserva, text="Confirmar Reserva",
                                                         command=self.confirmar_reserva)
        self.botao_adicionar_reserva.grid(row=3, column=1, sticky=W + E)


    def confirmar_reserva(self):
        self.mensagem_reservar_veiculo['text'] = ''

        try:
            self.tabela_reservar_veiculo.item(self.tabela_reservar_veiculo.selection())['text'][0]
        except IndexError as e:
            self.mensagem_reservar_veiculo['text'] = 'Por favor, selecione o veiculo a reservar'
            return

        matricula = self.tabela_reservar_veiculo.item(self.tabela_reservar_veiculo.selection())['text']
        marca = self.tabela_reservar_veiculo.item(self.tabela_reservar_veiculo.selection())['values'][0]
        modelo = self.tabela_reservar_veiculo.item(self.tabela_reservar_veiculo.selection())['values'][1]
        ocupacao = self.tabela_reservar_veiculo.item(self.tabela_reservar_veiculo.selection())['values'][2]
        data_proxima_revisao = self.tabela_reservar_veiculo.item(self.tabela_reservar_veiculo.selection())['values'][3]
        data_ult_inspecao = self.tabela_reservar_veiculo.item(self.tabela_reservar_veiculo.selection())['values'][4]
        data_inicio = self.input_data_inicio_reserva.get()
        data_fim = self.input_data_fim_reserva.get()
        nome_cliente = self.input_reservar_cliente.get()
        metodo_pagamento = self.input_metodo_pagamento.get()

        if len(data_inicio) == 0 or len(data_fim) == 0:
            self.mensagem_reservar_veiculo['text'] = 'Por favor, selecione a data de inicio e de fim da reserva'
            return

        if len(metodo_pagamento) == 0:
            self.mensagem_reservar_veiculo['text'] = 'Por favor, selecione o método de pagamento.'
            return

        if ocupacao != "Livre":
            self.mensagem_reservar_veiculo['text'] = 'O veiculo selecionado encontra-se ocupado'
            return

        if self.diferenca_datas_em_dias(data_inicio, data_fim) < 0:
            self.mensagem_reservar_veiculo['text'] = 'DATAS INVÁLIDAS: A data final tem que ser posterior à data inicial !!'
            return

        # Veiculo fica indisponivel caso a data da proxima revisao seja inferior à data final da reserva
        if self.diferenca_datas_em_dias(data_fim, data_proxima_revisao) < 0:
            self.mensagem_reservar_veiculo['text'] = 'O veiculo selecionado não tem a revisão em dia nas datas da reserva selecionadas'
            return

        # Veiculo fica indisponivel caso a data da proxima inspecao seja maior do que 1 ano relativamente à data final da reserva
        if self.diferenca_datas_em_dias(data_ult_inspecao, data_fim) > 365:
            self.mensagem_reservar_veiculo['text'] = 'O veiculo selecionado não tem a inspeção em dia nas datas da reserva selecionadas'
            return

        custo = round((self.diferenca_datas_em_dias(data_inicio, data_fim) + 1) *
                      self.get_custoHora_byMatricula(matricula), 2)

        # Armazenar Reserva na Base de Dados
        query = 'INSERT INTO reservas VALUES(NULL, ?, ?, ?, ?, ?, ?)'
        parametros = ( matricula, self.get_nif_byNome(nome_cliente), data_inicio, data_fim, custo, 1)
        self.db_consulta(query, parametros)

        # Colocar carro como ocupado
        query = 'UPDATE veiculos SET ocupado=? WHERE matricula = ?'
        parametros = (True, matricula, )
        self.db_consulta(query, parametros)  # Executar a consulta

        # Armazenar Custos na Tabela "pagamentos" na Base de Dados
        query = 'INSERT INTO pagamentos VALUES(NULL, ?, ?, ?, ?, ?, ?, ?)'
        parametros = (datetime.now().strftime('%d/%m/%Y'), custo, metodo_pagamento, self.get_nif_byNome(nome_cliente), matricula, data_inicio, data_fim)
        self.db_consulta(query, parametros)

        self.janela_adicionar_reserva.destroy()  # Fechar a janela de edição de veiculos
        messagebox.showinfo("", 'Nome do cliente: {}\nMatricula do carro: {}\nQuantidade de dias da reserva: {}\nCusto a pagar: {} €\n Método de pagamento: {}'
                            .format(nome_cliente, matricula, self.diferenca_datas_em_dias(data_inicio, data_fim) + 1, custo, metodo_pagamento))

        self.mensagem_reservas['text'] = 'O veiculo com matricula {} foi reservado com sucesso.'.format(matricula, nome_cliente)  # Mostrar mensagem para o utilizador

        self.get_reservas()

    def cancelar_reserva(self):
        """Elimina cliente através da matricula"""
        self.mensagem_reservas['text'] = ''
        try:
            self.tabela_reservas.item(self.tabela_reservas.selection())['values'][0]
        except IndexError as e:
            self.mensagem_veiculo['text'] = "Por favor, selecione uma reserva"
            return

        matricula = self.tabela_reservas.item(self.tabela_reservas.selection())['text']
        nif = self.tabela_reservas.item(self.tabela_reservas.selection())['values'][0]

        '''# ELiminar Reserva
        query = 'DELETE FROM reservas WHERE matricula_veiculo = ?'  # Consulta SQL
        self.db_consulta(query, (matricula,))  # Executar a consulta'''

        # Finalizar Reserva
        query = 'UPDATE reservas SET estado=? WHERE matricula_veiculo=?'
        parametros = (False, matricula, )
        self.db_consulta(query, parametros)

        # Passar o veiculo a Livre
        query = 'UPDATE veiculos SET ocupado=? WHERE matricula = ?'
        parametros = (False, matricula,)
        self.db_consulta(query, parametros)  # Executar a consulta

        self.mensagem_reservas['text'] = 'A reserva do carro com matricula: {} e do cliente com NIF: {} foi cancelada com êxito'.format(matricula, nif)
        self.get_reservas()  # Atualizar a tabela de clientes

    def get_reservas(self):
        # Limpar tabela
        reservas_tabela = self.tabela_reservas.get_children()
        for linha in reservas_tabela:
            self.tabela_reservas.delete(linha)

        # Consulta SQL
        query = 'SELECT * FROM reservas ORDER BY estado ASC, data_fim ASC'
        reservas = self.db_consulta(query)

        # Escrever dados na tabela
        for linha in reservas:
            dias_restantes = self.diferenca_atual_em_dias(linha[4]) if linha[6] else ''
            estado = "Em execução" if linha[6] else "Finalizada"
            self.tabela_reservas.insert('', 0, text=linha[1], values=(linha[2], linha[3], linha[4], linha[5], dias_restantes, estado))

    def exportar_reservas(self):
        """Criar um ficheiro Excel com as Reservas"""
        # Consulta SQL
        query = 'SELECT * FROM reservas'
        reservas = self.db_consulta(query)
        dados = []
        for reserva in reservas:
            dados.append(list(reserva)[1:-1])

        #dados = [i[1:] for i in dados]

        columns = ["Matricula Veiculo", "NIF Cliente", "Data Inicio", "Data Fim", "Custo (€)"]
        df = pd.DataFrame(dados, columns=columns)

        # Obtendo a data de hoje no formato desejado
        data_hoje = datetime.now().strftime('%d-%m-%Y')  # Formato: DD-MM-YYYY+
        # Definindo o nome do arquivo com a data de hoje
        nome_arquivo = f'relatorio_reservas_{data_hoje}.xlsx'
        # Exportando o DataFrame para um arquivo Excel
        df.to_excel(nome_arquivo, index=False)
        self.mensagem_reservas['text'] = "Os dados foram exportados com sucesso !!"


    def get_veiculos_por_ocupacao(self):
        """Retorna todos os veiculos da Bese de Dados ordendo por veiculos Livres até os Ocupados"""
        # Limpar tabela
        registos_tabela = self.tabela_reservar_veiculo.get_children()
        for linha in registos_tabela:
            self.tabela_reservar_veiculo.delete(linha)

        # Consulta SQL
        query = 'SELECT matricula, marca, modelo, ocupado, data_proxima_revisao, data_ult_inspecao FROM veiculos ORDER BY ocupado DESC'
        veiculos_por_ocupacao = self.db_consulta(query)

        # Escrever dados na tabela
        for linha in veiculos_por_ocupacao:
            ocupacao = "Livre" if linha[3] == 0 else "Ocupado"
            self.tabela_reservar_veiculo.insert('', 0, text=linha[0], values=(linha[1], linha[2], ocupacao, linha[4], linha[5]))

    def get_nome_clientes(self):
        # Consulta SQL
        query = 'SELECT nome FROM clientes'
        nome_clientes = self.db_consulta(query)
        lista = []
        for nome in nome_clientes:
            lista.append(nome[0])
        return lista

    def get_custoHora_byMatricula(self, matricula):
        # Consulta SQL
        query = 'SELECT valor_diario FROM veiculos WHERE matricula=?'
        parametros = (matricula, )
        nome_clientes = self.db_consulta(query, parametros)
        dado = []
        for i in nome_clientes:
            dado.append(i)
        return dado[0][0]

    def get_nif_byNome(self, nome):
        # Consulta SQL
        query = 'SELECT nif FROM clientes WHERE nome=?'
        parametros = (nome,)
        nif_cliente = self.db_consulta(query, parametros)
        dado = []
        for i in nif_cliente:
            dado.append(i)
        return dado[0][0]

# ------------------------------ FIM PAGINA RESERVAS ----------------------------

# ---------------------------- INICIO PAGINA PAGAMENTOS -------------------------

    def pagina_pagamentos(self):
        # Limpar pagina
        self.limpar_pagina()

        # Titulo da página
        self.titulo_pagina_pagamentos = Label(text="Página Pagamentos", fg='blue', font=('Calibri', 16, 'bold'))
        self.titulo_pagina_pagamentos.grid(row=0, column=1)

        # Tabela com Reservas
        style = ttk.Style()
        style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Calibri', 11))
        style.configure("mystyle.Treeview.Heading", font=('Calibri', 13, 'bold'))
        style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])

        # Estrutura da tabela
        self.tabela_pagamentos = ttk.Treeview(self.janela, height=5, columns=('col1', 'col2', 'col3', 'col4', 'col5', 'col6', 'col7'),
                                            style="mystyle.Treeview")
        self.tabela_pagamentos.grid(row=1, column=0, columnspan=6)
        # Cabeçalhos
        self.tabela_pagamentos.heading('#0', text='Data do Pagamento', anchor=CENTER)
        self.tabela_pagamentos.heading('#1', text='Custo da Reserva (€)', anchor=CENTER)
        self.tabela_pagamentos.heading('#2', text='Metodo de Pagamento', anchor=CENTER)
        self.tabela_pagamentos.heading('#3', text='NIF do Cliente', anchor=CENTER)
        self.tabela_pagamentos.heading('#4', text='Matricula do Veiculo', anchor=CENTER)
        self.tabela_pagamentos.heading('#5', text='Data de Inicio da Reserva', anchor=CENTER)
        self.tabela_pagamentos.heading('#6', text='Data de Fim da Reserva', anchor=CENTER)
        # Configuração da largura das colunas
        self.tabela_pagamentos.column('#0', width=130, anchor=CENTER)
        self.tabela_pagamentos.column('col1', width=130, anchor=CENTER)
        self.tabela_pagamentos.column('col2', width=130, anchor=CENTER)
        self.tabela_pagamentos.column('col3', width=130, anchor=CENTER)
        self.tabela_pagamentos.column('col4', width=130, anchor=CENTER)
        self.tabela_pagamentos.column('col5', width=130, anchor=CENTER)
        self.tabela_pagamentos.column('col6', width=130, anchor=CENTER)
        self.get_pagamentos()

        # Mensagem Informativa
        self.mensagem_pagamentos = Label(self.janela, text='', fg='red')
        self.mensagem_pagamentos.grid(row=2, column=0, columnspan=2, sticky=W + E)

        # Botao Exportar Reservas
        self.botao_exportar_pagamentos = ttk.Button(self.janela, text="Exportar Pagamentos", command=self.exportar_pagamentos)
        self.botao_exportar_pagamentos.grid(row=3, column=0, sticky=W + E)


    def get_pagamentos(self):
        # Limpar tabela
        tabela_pagamentos = self.tabela_pagamentos.get_children()
        for linha in tabela_pagamentos:
            self.tabela_pagamentos.delete(linha)

        # Consulta SQL
        query = 'SELECT * FROM pagamentos'
        pagamentos = self.db_consulta(query)

        custo_total = 0
        contador = 0
        # Escrever dados na tabela
        for linha in pagamentos:
            self.tabela_pagamentos.insert('', 0, text=linha[1], values=(linha[2], linha[3], linha[4], linha[5], linha[6], linha[7]))
            custo_total += linha[2]
            contador+=1
        self.tabela_pagamentos.insert('', contador+1, text="TOTAL (€):", values=(round(custo_total, 2)))

    def exportar_pagamentos(self):
        """Criar um ficheiro Excel com os Pagamentos"""
        # Consulta SQL
        query = 'SELECT * FROM pagamentos'
        pagamentos = self.db_consulta(query)
        dados = []
        for pagamento in pagamentos:
            dados.append(list(pagamento)[1:])

        columns = ["Data do Pagamento", "Custo da Reserva (€)", "Método de Pagamento", "NIF Cliente", "Matricula Veiculo", "Data Inicio da Reserva", "Data Fim da Reserva"]
        df = pd.DataFrame(dados, columns=columns)

        # Obtendo a data de hoje no formato desejado
        data_hoje = datetime.now().strftime('%d-%m-%Y')  # Formato: DD-MM-YYYY+
        # Definindo o nome do arquivo com a data de hoje
        nome_arquivo = f'relatorio_pagamentos_{data_hoje}.xlsx'
        # Exportando o DataFrame para um arquivo Excel
        df.to_excel(nome_arquivo, index=False)
        self.mensagem_pagamentos['text'] = "Os dados foram exportados com sucesso !!"



# ------------------------------ FIM PAGINA PAGAMENTOS -------------------------

    def db_consulta(self, consulta, parametros = ()):
        with sqlite3.connect(self.db) as con:
            cursor = con.cursor()
            resultado = cursor.execute(consulta, parametros)
            con.commit()
            return resultado.fetchall()

    def diferenca_atual_em_dias(self, data_str):
        """
        Calcula a diferença em dias entre a data fornecida e a data de hoje
        Retorna numero negativo se a data inserida for anterior à atual
        """
        # Converter a string para objeto datetime
        try:
            data = datetime.strptime(data_str, "%d/%m/%Y").date()
        except ValueError:
            return None

        data_atual = date.today() # Obter a data de hoje
        diferenca = data - data_atual  # Calcular a diferença

        return diferenca.days  # Retornar a diferença em dias

    def diferenca_datas_em_dias(self, data1, data2):
        """
        Calcula a diferença em dias entre a data1 e a data2
        Retorna numero negativo se a data2 for anterior à data1
        """
        # Convertendo as strings para objetos datetime
        data_inicio_dt = datetime.strptime(data1, "%d/%m/%Y")
        data_fim_dt = datetime.strptime(data2, "%d/%m/%Y")
        diferenca = data_fim_dt - data_inicio_dt
        return diferenca.days

    def iniciar_agendamento(self):
        # Agendar a tarefa para meia-noite
        schedule.every().day.at("12:15").do(self.tarefa_diaria)

        def executar_agendamentos():
            while True:
                schedule.run_pending()
                time.sleep(1)  # Aguardar 1 segundo antes de verificar novamente

        # Rodar o agendador em uma thread separada
        agendador_thread = Thread(target=executar_agendamentos, daemon=True)
        agendador_thread.start()

    def tarefa_diaria(self):
        '''
        Caso a data final da reseva chegar ao fim:
            - Desativa essa reserva
            - Coloca o veiculo dessa reserva como Livre
        '''
        print("A fazer tarefa diaria .....")
        # Consulta SQL
        query = 'SELECT * FROM reservas'
        reservas = self.db_consulta(query)

        for reserva in reservas:
            dias_restantes = self.diferenca_atual_em_dias(reserva[4])
            if dias_restantes < 0:
                query = 'UPDATE reservas SET estado=? WHERE rowid=?'  # Finalizar Reserva
                parametros = (False, reserva[0],)
                self.db_consulta(query, parametros)

                matricula = reserva[1]                          # Coloca o veiculo como Livre
                query = 'UPDATE veiculos SET ocupado=? WHERE matricula = ?'
                parametros = (False, matricula,)
                self.db_consulta(query, parametros)

    def sair(self):
        root.quit()


if __name__ == '__main__':
    root = TkinterDnD.Tk()  # Instancia da janela principal
    app = AppDesktop(root)  # Envia-se para a classe Veiculos o controlo sobre a janela root
    app.iniciar_agendamento()
    root.mainloop() # Começamos o cilco de aplicação
