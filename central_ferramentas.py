import sqlite3
import sys
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import crud_ferramentas as crud
import xlsxwriter
import pandas as pd
from tkcalendar.dateentry import DateEntry
import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from cores import *

################# cores ###############

co0 = "#CDBE70"  # bege escuro
co1 = "#feffff"  # branca
co2 = "#4fa882"  # verde
co3 = "#38576b"  # valor
co4 = "#403d3d"   # letra
co5 = "#e06636"   # - profit
co6 = "#8470FF"   # azul
co7 = "#ef5350"   # vermelha
co8 = "#263238"   # + verde
co9 = "#00C957"   # + verde
co10 = "#D9D9D9"  # cinza fundo
co11 = "#000080"  # azul login titulo
co12 = "#848484"  # cinza login corpo

#Criando e conectando banco de dados para Cadastro Ferramentas

conexao = sqlite3.connect('ferramentas_clientes.db')
c = conexao.cursor()
c.execute("""CREATE TABLE IF NOT EXISTS cadastro_ferramentas (
        id INTEGER PRIMARY KEY,
        descrição TEXT(60),
        fabricante TEXT(30),
        voltagem TEXT(15),
        part_number TEXT(25),
        tamanho INTEGER(20),
        unidade_medida TEXT(15),
        tipo TEXT(15),
        material TEXT(15),
        tempo_maximo INTEGER
        )
""")
conexao.commit()
c.close()
conexao.close()

#criando BD para cadastro de técnicos

conexao = sqlite3.connect('ferramentas_clientes.db')
cursor = conexao.cursor()
cursor.execute('''CREATE TABLE IF NOT EXISTS cadastro_tecnico (
        cpf TEXT,
        nome TEXT(40),
        telefone TEXT,
        turno TEXT,
        equipe TEXT(30)
)''')
conexao.commit()
cursor.close()
conexao.close()


# criando tabela no BD para Solicitação de Ferramentas

conexao = sqlite3.connect('ferramentas_clientes.db')
cursor = conexao.cursor()
cursor.execute('''CREATE TABLE IF NOT EXISTS cadastro_solicitação (
        id_ferramenta INTEGER,
        descri_solic TEXT(60),
        data_ret TEXT(20),
        hora_ret TEXT(10),
        data_dev TEXT(20),
        hora_dev TEXT(10),
        nome_tec TEXT(50)
        )
''')
conexao.commit()
cursor.close()
conexao.close()


#Criando classe PlaceHolder

class PlaceholderEntry(ttk.Entry):
    #placeholder texto nos campos
    def __init__(self, parent, placeholder='', color='#888', *args, **kwargs) -> None:
        super().__init__(parent, *args, **kwargs)
        self.placeholder = placeholder
        self._ph_color = color
        self._default_fg = self._get_fg_string()
        self.bind('<FocusIn>', self.clear_placeholder)
        self.bind('<FocusOut>', self.set_placeholder)
        self.set_placeholder()

    def clear_placeholder(self, *args) -> None:
        if self._get_fg_string() == self._ph_color:
            self.delete(0, tk.END)
            self.configure(foreground=self._default_fg)

    def set_placeholder(self, *args) -> None:
        if not self.get():
            self.insert(0, self.placeholder)
            self.configure(foreground=self._ph_color)

    def _get_fg_string(self) -> str:
        return str(self.cget('foreground'))

#Criando classe janela_ferramentas

class janela_ferramentas:
    def __init__(self, janela):

        self.ferramentasBD = crud.crud_ferramentas()
        #criando abas(notebooks)

        self.notebook = ttk.Notebook(janela)
        self.notebook.place(x=0, y=0, width=1100, height=600)

        self.aba1 = ttk.Frame(self.notebook)
        self.notebook.add(self.aba1, text='Ferramentas')

        self.aba2 = ttk.Frame(self.notebook)
        self.notebook.add(self.aba2, text='Funcionários')

        self.aba3 = ttk.Frame(self.notebook)
        self.notebook.add(self.aba3, text='Solicitação de reserva')

        #componentes:
        #labels

        self.id_lbl = tk.Label(self.aba1, text='ID:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.descricao_lbl = tk.Label(self.aba1, text='Descrição:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.fabricante_lbl = tk.Label(self.aba1, text='Fabricante:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.voltagem_lbl = tk.Label(self.aba1, text='Voltagem:',font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.part_number_lbl = tk.Label(self.aba1, text='Número da partição:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.tamanho_lbl = tk.Label(self.aba1, text='Tamanho:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.unidade_medida_lbl = tk.Label(self.aba1, text='Unidade de medida:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.tipo_lbl = tk.Label(self.aba1, text='Tipo de ferramenta:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.material_lbl = tk.Label(self.aba1, text='Material:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.tempo_maximo_lbl = tk.Label(self.aba1, text='Tempo máximo:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.aba_ferramentas_text = tk.Label(self.aba1, text='GERENCIAMENTO DE FERRAMENTAS', font=('Trebuchet 14 bold'), bg=co2, fg=co1, relief='flat')

        #widgets de entrada

        self.id_entry = PlaceholderEntry(self.aba1, width=10)
        self.id_entry.bind("<KeyRelease>", self.avisoID)
        self.descricao_entry = PlaceholderEntry(self.aba1, placeholder='Entre com a descrição da ferramenta', width=120)
        self.fabricante_entry = PlaceholderEntry(self.aba1, placeholder='Entre com a fabricante da ferramenta', width=40)
        selecionar_voltagem = tk.StringVar()
        self.voltagem_combobox = ttk.Combobox(self.aba1, textvariable=selecionar_voltagem)
        self.voltagem_combobox['values'] = ('110', '220', 'N/A')
        self.voltagem_combobox.bind('<<ComboboxSelected>>')
        self.part_number_entry = PlaceholderEntry(self.aba1, placeholder='Número no fabricante', width=25)
        self.tamanho_entry = PlaceholderEntry(self.aba1, placeholder='Número', width=10)
        selecionar_unidade_medida = tk.StringVar()
        self.unidade_medida_combobox = ttk.Combobox(self.aba1, textvariable=selecionar_unidade_medida)
        self.unidade_medida_combobox['values'] = ('Centímetros', 'Metros', 'Polegadas')
        self.unidade_medida_combobox.bind('<<ComboboxSelected>>')
        self.tipo_entry = PlaceholderEntry(self.aba1, placeholder='Elétrica, mecânica, segurança...', width=50)
        self.material_entry = PlaceholderEntry(self.aba1, placeholder='Metal, plástico, borracha...', width=40)
        self.tempo_maximo_entry = PlaceholderEntry(self.aba1, placeholder='Reserva em horas', width=20)

        #componente treeview

        colunas_treeview = ('ID', 'Descrição', 'Fabricante', 'Voltagem', 'Número da partição', 'Tamanho',
                            'Unidade de medida', 'Tipo de ferramenta', 'Material', 'Tempo máximo')
        self.ferramentas_treeview = ttk.Treeview(self.aba1, columns=colunas_treeview, show='headings', selectmode='browse')
        self.ferramentas_treeview.heading('ID', text='ID')
        self.ferramentas_treeview.heading('Descrição', text='Descrição')
        self.ferramentas_treeview.heading('Fabricante', text='Fabricante')
        self.ferramentas_treeview.heading('Voltagem', text='Voltagem')
        self.ferramentas_treeview.heading('Número da partição', text='Número da partição')
        self.ferramentas_treeview.heading('Tamanho', text='Tamanho')
        self.ferramentas_treeview.heading('Unidade de medida', text='Unidade de medida')
        self.ferramentas_treeview.heading('Tipo de ferramenta', text='Tipo de ferramenta')
        self.ferramentas_treeview.heading('Material', text='Material')
        self.ferramentas_treeview.heading('Tempo máximo', text='Tempo máximo')

        self.ferramentas_treeview.column('ID', minwidth=1, width=25)
        self.ferramentas_treeview.column('Descrição', minwidth=1, width=210)
        self.ferramentas_treeview.column('Fabricante', minwidth=1, width=100)
        self.ferramentas_treeview.column('Voltagem', minwidth=1, width=60)
        self.ferramentas_treeview.column('Número da partição', minwidth=1, width=80)
        self.ferramentas_treeview.column('Tamanho', minwidth=1, width=70)
        self.ferramentas_treeview.column('Unidade de medida', minwidth=1, width=150)
        self.ferramentas_treeview.column('Tipo de ferramenta', minwidth=1, width=150)
        self.ferramentas_treeview.column('Material', minwidth=1, width=150)
        self.ferramentas_treeview.column('Tempo máximo', minwidth=1, width=60)

        self.treeview_scrollbar_x = ttk.Scrollbar(self.aba1, orient='horizontal', command=self.ferramentas_treeview.xview)
        self.treeview_scrollbar_x.pack(side='bottom', fill='x')
        self.treeview_scrollbar = ttk.Scrollbar(self.aba1, orient='vertical', command=self.ferramentas_treeview.yview)
        self.treeview_scrollbar.pack(side='right', fill='x')
        self.ferramentas_treeview.configure(yscrollcommand=self.treeview_scrollbar.set,
                                            xscrollcommand=self.treeview_scrollbar_x.set)
        self.ferramentas_treeview.bind('<<TreeviewSelect>>', self.mostraFerramentaSelect)
        self.carregaFerramentas()

        #criação de botões

        self.btnCadastrarFerramenta = tk.Button(self.aba1, text='Cadastrar', width=14, font=("Trebuchet 10 bold"), fg=co1,
                                             bg=co9, relief='raised', overrelief='ridge', command=self.fCadastrarProduto)
        self.btnAtualizarFerramenta = tk.Button(self.aba1, text='Atualizar', width=14, font=("Trebuchet 10 bold"), fg=co1,
                                             bg=co6, relief='raised', overrelief='ridge', command=self.fAtualizaFerramenta)
        self.btnDeletarFerramenta = tk.Button(self.aba1, text='Deletar', width=14, font=("Trebuchet 10 bold"), fg=co1,
                                             bg=co7, relief='raised', overrelief='ridge', command=self.fDeletaProduto)
        self.btnLimparFerramenta = tk.Button(self.aba1, text='Limpar', width=14, font=("Trebuchet 10 bold"), fg=co1,
                                             bg=co0, relief='raised', overrelief='ridge', command=self.fLimpaCampos)
        self.btnExportarExcelFerramenta = tk.Button(self.aba1, text='Exportar para Excel', width=14, font=("Trebuchet 10 bold"), fg=co1,
                                             bg=co8, relief='raised', overrelief='ridge', command=self.fExportarExcel)


        #posicionando componentes
        #posicionando labels

        self.id_lbl.place(x=20, y=50)
        self.descricao_lbl.place(x=140, y=50)
        self.fabricante_lbl.place(x=20, y=100)
        self.voltagem_lbl.place(x=390, y=100)
        self.part_number_lbl.place(x=650, y=100)
        self.tamanho_lbl.place(x=20, y=150)
        self.unidade_medida_lbl.place(x=200, y=150)
        self.tipo_lbl.place(x=513, y=150)
        self.material_lbl.place(x=20, y=200)
        self.tempo_maximo_lbl.place(x=400, y=200)
        self.aba_ferramentas_text.place(x=370, y=10)

        #posicionando widgets de entrada

        self.id_entry.place(x=50, y=53)
        self.descricao_entry.place(x=235, y=53)
        self.fabricante_entry.place(x=120, y=103)
        self.voltagem_combobox.place(x=481, y=103)
        self.part_number_entry.place(x=822, y=103)
        self.tamanho_entry.place(x=110, y=153)
        self.unidade_medida_combobox.place(x=365, y=153)
        self.tipo_entry.place(x=680, y=153)
        self.material_entry.place(x=100, y=203)
        self.tempo_maximo_entry.place(x=540, y=203)

        self.ferramentas_treeview.place(x=20, y=250)
        self.treeview_scrollbar_x.place(x=20, y=476, width=1059)
        self.treeview_scrollbar.place(x=1079, y=250, height=225)

        #posicionando botões
        self.btnCadastrarFerramenta.place(x=40, y=510)
        self.btnAtualizarFerramenta.place(x=180, y=510)
        self.btnDeletarFerramenta.place(x=320, y=510)
        self.btnLimparFerramenta.place(x=460, y=510)
        self.btnExportarExcelFerramenta.place(x=600, y=510, width=200)

########################################################################################################################
#                        CRIAÇÃO DA ABA 2 CADASTRO TÉCNICOS
########################################################################################################################

#criando labels

        self.cpf_label = tk.Label(self.aba2, text='CPF:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.nome_tec_label = tk.Label(self.aba2, text='Nome:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.telefone_label = tk.Label(self.aba2, text='Telefone ou Rádio:',font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.turno_label = tk.Label(self.aba2, text='Turno:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.equipe_label = tk.Label(self.aba2, text='Equipe:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.cadastro_tec_label = tk.Label(self.aba2, text='CADASTRO DE FUNCIONÁRIOS',
                                           font=('Trebuchet 14 bold'), bg=co2, fg=co1, relief='flat')

#criando entrys

        tcpf = (janela.register(self.validaCPF))
        self.cpf_entry = PlaceholderEntry(self.aba2, placeholder='ex: 123.456.789-00', width=45, validate='all',
                                          validatecommand=(tcpf, '%d', '%s', '%S', '%v', '%V'))
        self.cpf_entry.bind("<KeyRelease>", self.preencheCPF)
        self.nome_tec_entry = PlaceholderEntry(self.aba2, placeholder='ex: João da Silva Costa', width=45)
        self.telefone_entry = PlaceholderEntry(self.aba2, placeholder='ex: 988556677', width=30)
        turno_combo = tk.StringVar()
        self.turno_combobox = ttk.Combobox(self.aba2, width=30, textvariable=turno_combo)
        self.turno_combobox['values'] = ('Manhã', 'Tarde', 'Noite')
        self.turno_combobox.bind('<<ComboboxSelected>>')
        self.equipe_entry = PlaceholderEntry(self.aba2, placeholder='ex: Manutenção externa', width=45)

#posicionando labels

        self.cpf_label.place(x=20, y=75)
        self.nome_tec_label.place(x=20, y=150)
        self.telefone_label.place(x=20, y=225)
        self.turno_label.place(x=20, y=300)
        self.equipe_label.place(x=20, y=375)
        self.cadastro_tec_label.place(x=12, y=20)

#posicionado entrys

        self.cpf_entry.place(x=20, y=100)
        self.nome_tec_entry.place(x=20, y=175)
        self.telefone_entry.place(x=20, y=250)
        self.turno_combobox.place(x=20, y=325)
        self.equipe_entry.place(x=20, y=400)

#criando e posicionando botões

        self.btnCadastrarTecnico = tk.Button(self.aba2, text='Cadastrar', width=14, font=("Trebuchet 10 bold"), fg=co1,
                                             bg=co9, relief='raised', overrelief='ridge', command=self.fCadastrarTecnico)
        self.btnDeletarTecnico = tk.Button(self.aba2, text='Deletar', width=14, font=("Trebuchet 10 bold"), fg=co1,
                                           bg=co7, relief='raised', overrelief='ridge', command=self.fDeletarTecnico)
        self.btnAtualizarTecnico = tk.Button(self.aba2, text='Atualizar', width=14, font=("Trebuchet 10 bold"), fg=co1,
                                             bg=co6, relief='raised', overrelief='ridge', command=self.fAtualizaTecnico)
        self.btnLimparTecnico = tk.Button(self.aba2, text='Limpar', width=14, font=("Trebuchet 10 bold"), fg=co1,
                                          bg=co0, relief='raised', overrelief='ridge', command=self.limpaTecnico)
        self.btnExportarExcelTecnico = tk.Button(self.aba2, text='Exportar para Excel', width=18,
                                                 font=("Trebuchet 10 bold"), fg=co1, bg=co8, relief='raised',
                                                 overrelief='ridge', command=self.tecnicoExportarExcel)

        self.btnCadastrarTecnico.place(x=20, y=440)
        self.btnDeletarTecnico.place(x=160, y=440)
        self.btnAtualizarTecnico.place(x=20, y=480)
        self.btnLimparTecnico.place(x=160, y=480)
        self.btnExportarExcelTecnico.place(x=70, y=520)

#componente treeview

        colunas_tec_treeview = ('CPF', 'Nome','Telefone ou Rádio', 'Turno', 'Equipe')
        self.tecnico_treeview = ttk.Treeview(self.aba2, columns=colunas_tec_treeview, show='headings', selectmode='browse')
        self.tecnico_treeview.heading('CPF', text='CPF')
        self.tecnico_treeview.heading('Nome', text='Nome')
        self.tecnico_treeview.heading('Telefone ou Rádio', text='Telefone ou Rádio')
        self.tecnico_treeview.heading('Turno', text='Turno')
        self.tecnico_treeview.heading('Equipe', text='Equipe')

        self.tecnico_treeview.column('CPF', minwidth=1, width=150)
        self.tecnico_treeview.column('Nome', minwidth=1, width=170)
        self.tecnico_treeview.column('Telefone ou Rádio', minwidth=1, width=150)
        self.tecnico_treeview.column('Turno', minwidth=1, width=120)
        self.tecnico_treeview.column('Equipe', minwidth=1, width=150)

        self.scrollbar_tecnico = ttk.Scrollbar(self.aba2, orient='vertical',
                                                  command=self.tecnico_treeview.yview)
        self.scrollbar_tecnico.pack(side='right', fill='x')
        self.tecnico_treeview.configure(yscrollcommand=self.scrollbar_tecnico.set)
        self.tecnico_treeview.bind('<<TreeviewSelect>>', self.mostraTecnicoSelect)
        self.carregaTecnicos()

        self.tecnico_treeview.place(x=320, y=20, height=500)
        self.scrollbar_tecnico.place(x=1064, y=20, height=500)


########################################################################################################################
#                        CRIAÇÃO DA ABA 3 SOLICITAÇÃO DE RESERVA DE FERRAMENTAS
########################################################################################################################

        # criando labels

        self.id_ferramenta_label = tk.Label(self.aba3, text='Código da ferramenta:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.descri_solic_label = tk.Label(self.aba3, text='Descreva a solicitação:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.data_ret_label = tk.Label(self.aba3, text='Data e hora da retirada:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.data_dev_label = tk.Label(self.aba3, text='Devolução:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.tec_resp_label = tk.Label(self.aba3, text='Nome completo do técnico:', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.aba_reserva_text = tk.Label(self.aba3, text='SOLITICAÇÃO DE RESERVA DE FERRAMENTA', font=('Trebuchet 14 bold'),
                                         bg=co2, fg=co1, relief='flat')
        self.hora_as_label = tk.Label(self.aba3, text='as', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.hora_as1_label = tk.Label(self.aba3, text='as', font=('Trebuchet 13 bold'), bg=co10, fg=co4, relief='flat')
        self.aviso_email_text = tk.Label(self.aba3, text='**A solicitação só será confirmada após cadastro do pedido e envio por email.**',
                                         font='bold', bg='brown', fg='white', height=1)

        # criando entrys

        self.id_ferramenta_entry = PlaceholderEntry(self.aba3, placeholder='ex:001', width=10)
        self.descri_solic_entry = PlaceholderEntry(self.aba3, placeholder='Descreva sua solicitação', width=120)
        self.datahora_ret_entry = DateEntry(self.aba3, width=12, background='darkblue',
                                            foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy')
        dre = self.datahora_ret_entry.get_date()
        self.data_ret_entry = dre.strftime('%Y-%m-%d')
        self.datahora_dev_entry = DateEntry(self.aba3, width=12, background='darkblue',
                                            foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy')
        dde = self.datahora_dev_entry.get_date()
        self.data_dev_entry = dde.strftime('%Y-%m-%d')
        vcmd = (janela.register(self.ValidarEntry))
        self.hora_ret_entry = PlaceholderEntry(self.aba3, placeholder='ex: 13:00', width=15,
                                               validate='all', validatecommand=(vcmd, '%d', '%s', '%S'))
        self.hora_ret_entry.bind("<KeyRelease>", self.hora_24)
        self.hora_dev_entry = PlaceholderEntry(self.aba3, placeholder='ex: 13:00', width=15,
                                               validate='all', validatecommand=(vcmd, '%d', '%s', '%S'))
        self.hora_dev_entry.bind("<KeyRelease>", self.hora_24)
        self.tec_resp_entry = PlaceholderEntry(self.aba3, placeholder='ex: João da Silva Costa', width=60)

        #criando treeview

        colunas_treeview1 = ('ID da ferramenta', 'Descrição da solicitação', 'Data de retirada', 'Hora retirada',
                             'Data devolução', 'Hora devolução', 'Nome do técnico')
        self.solicitacao_treeview = ttk.Treeview(self.aba3, columns=colunas_treeview1, selectmode='browse', show='headings')
        self.solicitacao_treeview.heading('ID da ferramenta', text='ID da ferramenta')
        self.solicitacao_treeview.heading('Descrição da solicitação', text='Descrição da solicitação')
        self.solicitacao_treeview.heading('Data de retirada', text='Data de retirada')
        self.solicitacao_treeview.heading('Hora retirada', text='Hora retirada')
        self.solicitacao_treeview.heading('Data devolução', text='Data devolução')
        self.solicitacao_treeview.heading('Hora devolução', text='Hora devolução')
        self.solicitacao_treeview.heading('Nome do técnico', text='Nome do técnico')

        self.solicitacao_treeview.column('ID da ferramenta', minwidth=1, width=100)
        self.solicitacao_treeview.column('Descrição da solicitação', minwidth=1, width=250)
        self.solicitacao_treeview.column('Data de retirada', minwidth=1, width=110)
        self.solicitacao_treeview.column('Hora retirada', minwidth=1, width=110)
        self.solicitacao_treeview.column('Data devolução', minwidth=1, width=110)
        self.solicitacao_treeview.column('Hora devolução', minwidth=1, width=110)
        self.solicitacao_treeview.column('Nome do técnico', minwidth=1, width=250)

        self.treeview1_scrollbar = ttk.Scrollbar(self.aba3, orient='vertical', command=self.solicitacao_treeview.yview)
        self.treeview1_scrollbar.pack(side='right', fill='x')
        self.solicitacao_treeview.configure(yscrollcommand=self.treeview1_scrollbar.set)
        self.solicitacao_treeview.bind('<<TreeviewSelect>>', self.mostraSolicitacaoSelect)
        self.carregaSolicitacao()

        # criando botões

        self.btnLimparSolicitacao = tk.Button(self.aba3, text='Limpar', width=14, font=("Trebuchet 10 bold"), fg=co1,
                                             bg=co0, relief='raised', overrelief='ridge', command=self.limpaSolicitacao)
        self.btnCadastrarSolicitacao = tk.Button(self.aba3, text='Solicitar', width=14, font=("Trebuchet 10 bold"), fg=co1,
                                             bg=co9, relief='raised', overrelief='ridge', command=self.fCadastrarSolicitacao)
        self.btnDeletarSolicitacao = tk.Button(self.aba3, text='Deletar solicitação', width=18, font=("Trebuchet 10 bold"), fg=co1,
                                             bg=co7, relief='raised', overrelief='ridge', command=self.fDeletarSolicitacao)
        self.btnExportarExcelSolicitacao = tk.Button(self.aba3, text='Exportar para Excel', width=22, font=("Trebuchet 10 bold"), fg=co1,
                                             bg=co8, relief='raised', overrelief='ridge', command=self.solicitacaoExportarExcel)
        self.btnEnviarEmail = tk.Button(self.aba3, text='Enviar solicitação por email', width=14, font=("Trebuchet 10 bold"), fg=co1,
                                             bg=co6, relief='raised', overrelief='ridge', command=self.enviarEmail)

        # posicionando labels

        self.id_ferramenta_label.place(x=20, y=50)
        self.descri_solic_label.place(x=20, y=90)
        self.data_ret_label.place(x=20, y=130)
        self.data_dev_label.place(x=460, y=130)
        self.tec_resp_label.place(x=20, y=170)
        self.aba_reserva_text.place(x=335, y=10)
        self.hora_as_label.place(x=320, y=130)
        self.hora_as1_label.place(x=660, y=130)
        self.aviso_email_text.place(x=250, y=465)

        # posicionando widgets

        self.id_ferramenta_entry.place(x=220, y=53)
        self.descri_solic_entry.place(x=220, y=93)
        self.datahora_ret_entry.place(x=220, y=133)
        self.datahora_dev_entry.place(x=560, y=133)
        self.tec_resp_entry.place(x=250, y=173)
        self.hora_ret_entry.place(x=353, y=133)
        self.hora_dev_entry.place(x=693, y=133)
        self.solicitacao_treeview.place(x=25, y=230)
        self.treeview1_scrollbar.place(x=1066, y=230, height=230)

        # posicionando botões

        self.btnLimparSolicitacao.place(x=415, y=500)
        self.btnCadastrarSolicitacao.place(x=100, y=500)
        self.btnDeletarSolicitacao.place(x=240, y=500)
        self.btnExportarExcelSolicitacao.place(x=560, y=500)
        self.btnEnviarEmail.place(x=770, y=500, width=200)

#validando entradas horas retirada e devolução

    def ValidarEntry(self, d, s, S):
        if d == "0":
            return True
        if ((S == ":" and len(s) != 2) or (not S.isdigit() and
                                           S != ":") or (len(s) == 3 and int(S) > 5) or len(s) > 4):
            return False
        return True

    def hora_24(self, event):
        s = event.widget
        if len(s.get()) == 2 and event.keysym == "BackSpace":
            s.delete(len(s.get()) - 1, tk.END)
        if event.keysym == "BackSpace":
            return
        if len(s.get()) == 1 and int(s.get()) > 2:
            s.insert(0, "0")
            s.insert("end", ":")
        elif len(s.get()) == 2 and int(s.get()) < 24:
            s.insert(2, ":")
        elif len(s.get()) >= 2 and s.get()[2:3] != ":":
            s.delete(1, tk.END)

#validando cpf correto

    def validaCPF(self, d, s, v, S, V):
        if d == '0':
            return True
        if (len(s) > 13) or (v.isalpha()):
            return False
        return True

    def preencheCPF(self, event):
        s = event.widget
        x = s.get()
        if (len(s.get()) == 4 or len(s.get()) == 8 or len(s.get()) == 12) and event.keysym == "BackSpace":
            s.delete(len(s.get()) - 2, tk.END)
        elif (len(x) == 3 or len(x) == 7 or len(x) == 11) and event.keysym == 'BackSpace':
            s.delete(len(x) - 1, tk.END)
        elif event.keysym == "BackSpace":
            return
        if len(s.get()) == 3:
            s.insert("end", ".")
        elif len(s.get()) == 7:
            s.insert("end", ".")
        elif len(s.get()) == 11:
            s.insert("end", "-")


#bloqueando entrada de dados no ID

    def avisoID(self, event):
        s = event.widget
        if s.get():
            s.delete(0, 'end')
            messagebox.showinfo('AUTOMÁTICO', message='Não é necessário inserir o ID, campo gerado automaticamente')

#validando telefone e rádio

    def validaTel(self):
        if len(self.telefone_entry.get()) == 8 or len(self.telefone_entry.get()) == 9:
            return True
        else:
            self.cpf_entry.delete(0, 'end')
            self.nome_tec_entry.delete(0, 'end')
            self.telefone_entry.delete(0, 'end')
            self.turno_combobox.delete(0, 'end')
            self.equipe_entry.delete(0, 'end')
            messagebox.showwarning('DADO INVÁLIDO', message='O telefone deve conter 9 dígitos ou o rádio 8 dígitos.')


#carregar dados das ferramentas no BD

    def carregaFerramentas(self):
        ferramentas = self.ferramentasBD.selecionarFerramenta()

        for ferramenta in ferramentas:
            id = ferramenta[0]
            descriçao = ferramenta[1]
            fabricante = ferramenta[2]
            voltagem = ferramenta[3]
            part_number = ferramenta[4]
            tamanho = ferramenta[5]
            unidade_medida = ferramenta[6]
            tipo = ferramenta[7]
            material = ferramenta[8]
            tempo_maximo = ferramenta[9]

            self.ferramentas_treeview.insert('', 'end', values=(id, descriçao, fabricante, voltagem,
                                         part_number, tamanho, unidade_medida, tipo, material, tempo_maximo))


#carregar solicitações

    def carregaSolicitacao(self):
        solicitacao = self.ferramentasBD.selecionaSolicitacao()

        for item in solicitacao:
            id_ferramenta = item[0]
            descri_solic = item[1]
            data_ret = item[2]
            hora_ret = item[3]
            data_dev = item[4]
            hora_dev = item[5]
            nome_tec = item[6]

            self.solicitacao_treeview.insert('', 'end', values=(id_ferramenta, descri_solic, data_ret, hora_ret,
                                                                data_dev, hora_dev, nome_tec))


#carrega tecnico
    def carregaTecnicos(self):
        tecnico = self.ferramentasBD.selecionaTecnico()

        for item in tecnico:
            cpf = item[0]
            nome = item[1]
            telefone = item[2]
            turno = item[3]
            equipe = item[4]

            self.tecnico_treeview.insert('', 'end', values=(cpf, nome, telefone, turno, equipe))

#apresentar ferramenta selecionada

    def mostraFerramentaSelect(self, event):
        self.fLimpaCampos()
        for ferramenta in self.ferramentas_treeview.selection():
            item = self.ferramentas_treeview.item(ferramenta)
            id, descriçao, fabricante, voltagem, part_number, tamanho, unidade_medida, tipo, material, tempo_maximo = item['values'][0:10]
            self.id_entry.insert(0, id)
            self.descricao_entry.insert(0, descriçao)
            self.fabricante_entry.insert(0, fabricante)
            self.voltagem_combobox.insert(0, voltagem)
            self.part_number_entry.insert(0, part_number)
            self.tamanho_entry.insert(0, tamanho)
            self.unidade_medida_combobox.insert(0, unidade_medida)
            self.tipo_entry.insert(0, tipo)
            self.material_entry.insert(0, material)
            self.tempo_maximo_entry.insert(0, tempo_maximo)

    def mostraTecnicoSelect(self, event):
        self.limpaTecnico()
        for tecnico in self.tecnico_treeview.selection():
            item = self.tecnico_treeview.item(tecnico)
            cpf, nome, telefone, turno, equipe = item['values'][0:5]
            self.cpf_entry.insert(0, cpf)
            self.nome_tec_entry.insert(0, nome)
            self.telefone_entry.insert(0, telefone)
            self.turno_combobox.insert(0, turno)
            self.equipe_entry.insert(0, equipe)

    def mostraSolicitacaoSelect(self, event):
        self.limpaSolicitacao()
        for solicitacao in self.solicitacao_treeview.selection():
            item = self.solicitacao_treeview.item(solicitacao)
            id_ferramenta = item['values'][0]
            self.id_ferramenta_entry.insert(0, id_ferramenta)

#função para ler os campospreenchidos ferramentas

    def fLerCampos(self):
        id = self.id_entry.get()
        descriçao = self.descricao_entry.get()
        fabricante = self.fabricante_entry.get()
        voltagem = self.voltagem_combobox.get()
        part_number = self.part_number_entry.get()
        tamanho = self.tamanho_entry.get()
        unidade_medida = self.unidade_medida_combobox.get()
        tipo = self.tipo_entry.get()
        material = self.material_entry.get()
        tempo_maximo = self.tempo_maximo_entry.get()
        return id, descriçao, fabricante, voltagem, part_number, tamanho, unidade_medida, tipo, material, tempo_maximo

#ler campos preenchidos de tecnicos

    def lerTecnico(self):
        cpf = self.cpf_entry.get()
        nome = self.nome_tec_entry.get()
        telefone = self.telefone_entry.get()
        turno = self.turno_combobox.get()
        equipe = self.equipe_entry.get()
        return cpf, nome, telefone, turno, equipe

#função ler campos preenchidos solicitação

    def lerSolicitacao(self):
        id_ferramenta = self.id_ferramenta_entry.get()
        descri_solic = self.descri_solic_entry.get()
        data_ret = self.datahora_ret_entry.get()
        hora_ret = self.hora_ret_entry.get()
        data_dev = self.datahora_dev_entry.get()
        hora_dev = self.hora_dev_entry.get()
        nome_tec = self.tec_resp_entry.get()
        return id_ferramenta, descri_solic, data_ret, hora_ret, data_dev, hora_dev, nome_tec

#função para cadastrar produto
    def fCadastrarProduto(self):
        id, descriçao, fabricante, voltagem, part_number, tamanho, unidade_medida, tipo, material, tempo_maximo = self.fLerCampos()
        self.ferramentasBD.cadastrarFerramenta(descriçao, fabricante, voltagem, part_number, tamanho,
                                               unidade_medida, tipo, material, tempo_maximo)
        self.ferramentas_treeview.insert('', 'end', values=(id, descriçao, fabricante, voltagem,
                                         part_number, tamanho, unidade_medida, tipo, material, tempo_maximo))
        self.ferramentas_treeview.delete(*self.ferramentas_treeview.get_children())
        self.carregaFerramentas()
        self.fLimpaCampos()
        messagebox.showinfo('Cadastro', 'Ferramenta cadastrada com sucesso!')

#cadastrar tecnico

    def fCadastrarTecnico(self):
        try:
            if self.validaTel() != True:
                return
            else:
                cpf, nome, telefone, turno, equipe = self.lerTecnico()
                self.ferramentasBD.cadastrarTecnico(cpf, nome, telefone, turno, equipe)
                self.tecnico_treeview.insert('', 'end', values=(cpf, nome, telefone, turno, equipe))
                self.tecnico_treeview.delete(*self.tecnico_treeview.get_children())
                self.carregaTecnicos()
                self.limpaTecnico()
                messagebox.showinfo('Cadastro', 'Técnico cadastrado com sucesso!')
        except:
            messagebox.showerror('ERRO', message='Não foi possível realizar o cadastro.')

#cadastrar solicitação

    def fCadastrarSolicitacao(self):
        id_ferramenta, descri_solic, data_ret, hora_ret, data_dev, hora_dev, nome_tec = self.lerSolicitacao()
        self.ferramentasBD.cadastraSolicitacao(id_ferramenta, descri_solic, data_ret, hora_ret, data_dev,
                                               hora_dev, nome_tec)
        self.solicitacao_treeview.insert('', 'end', values=(id_ferramenta, descri_solic, data_ret, hora_ret, data_dev,
                                                            hora_dev, nome_tec))
        self.limpaSolicitacao()
        messagebox.showinfo('Cadastro', 'Solicitação cadastrada com sucesso!')

#função para atualizar o cadastro de produtros

    def fAtualizaFerramenta(self):
        id, descriçao, fabricante, voltagem, part_number, tamanho, unidade_medida, tipo, material, tempo_maximo = self.fLerCampos()
        self.ferramentasBD.atualizaFerramenta(descriçao, fabricante, voltagem, part_number, tamanho,
                                               unidade_medida, tipo, material, tempo_maximo, id)
        #recarregar tela
        self.ferramentas_treeview.delete(*self.ferramentas_treeview.get_children())
        self.carregaFerramentas()
        self.fLimpaCampos()
        messagebox.showinfo('Atualização', 'Ferramenta atualizada com sucesso!')

    def fAtualizaTecnico(self):
        try:
            if self.validaTel() != True:
                return
            else:
                self.validaTel()
                cpf, nome, telefone, turno, equipe = self.lerTecnico()
                self.ferramentasBD.atualizaTecnico(nome, telefone, turno, equipe, cpf)
                self.tecnico_treeview.delete(*self.tecnico_treeview.get_children())
                self.carregaTecnicos()
                self.limpaTecnico()
                messagebox.showinfo('Atualização', 'Cadastro atualizado com sucesso!')
        except:
            messagebox.showerror('ERRO', message='Não foi possível realizar a atualização.')

#função para deletar o produto

    def fDeletaProduto(self):
        id, descriçao, fabricante, voltagem, part_number, tamanho, unidade_medida, tipo, material, tempo_maximo = self.fLerCampos()
        self.ferramentasBD.deletarFerramenta(id)

        #recarregar tela
        self.ferramentas_treeview.delete(*self.ferramentas_treeview.get_children())
        self.carregaFerramentas()
        self.fLimpaCampos()
        messagebox.showinfo('Deletar', 'Ferramenta deletada com sucesso!')

#deletar técnico

    def fDeletarTecnico(self):
        cpf, nome, telefone, turno, equipe = self.lerTecnico()
        self.ferramentasBD.deletaTecnico(cpf)
        self.tecnico_treeview.delete(*self.tecnico_treeview.get_children())
        self.carregaTecnicos()
        self.limpaTecnico()
        messagebox.showinfo('Deletar', 'Técnico deletado com sucesso!')

#deletar solicitação

    def fDeletarSolicitacao(self):
        id_ferramenta, descri_solic, data_ret, hora_ret, data_dev, hora_dev, nome_tec = self.lerSolicitacao()
        self.ferramentasBD.deletarSolicitacao(id_ferramenta)
        self.solicitacao_treeview.delete(*self.solicitacao_treeview.get_children())
        self.carregaSolicitacao()
        self.limpaSolicitacao()
        messagebox.showinfo('Deletar', 'Solicitação deletada com sucesso!')

#função para limpar dados

    def fLimpaCampos(self):
        self.id_entry.delete(0, 'end')
        self.descricao_entry.delete(0, 'end')
        self.fabricante_entry.delete(0, 'end')
        self.voltagem_combobox.delete(0, 'end')
        self.part_number_entry.delete(0, 'end')
        self.tamanho_entry.delete(0, 'end')
        self.unidade_medida_combobox.delete(0, 'end')
        self.tipo_entry.delete(0, 'end')
        self.material_entry.delete(0, 'end')
        self.tempo_maximo_entry.delete(0, 'end')

#função limpar tecnicos

    def limpaTecnico(self):
        self.cpf_entry.delete(0, 'end')
        self.nome_tec_entry.delete(0, 'end')
        self.telefone_entry.delete(0, 'end')
        self.turno_combobox.delete(0, 'end')
        self.equipe_entry.delete(0, 'end')

#limpar campos solicitação

    def limpaSolicitacao(self):
        self.id_ferramenta_entry.delete(0, 'end')
        self.descri_solic_entry.delete(0, 'end')
        self.datahora_ret_entry.delete(0, 'end')
        self.hora_ret_entry.delete(0, 'end')
        self.datahora_dev_entry.delete(0, 'end')
        self.hora_dev_entry.delete(0, 'end')
        self.tec_resp_entry.delete(0, 'end')

#função exportar para excel

    def fExportarExcel(self):
        conexao = self.conexaoFerramentas = sqlite3.connect('ferramentas_clientes.db')
        with pd.ExcelWriter("banco_ferramentas.xlsx", engine="xlsxwriter",
                            engine_kwargs={'options':{'strings_to_numbers': True, 'strings_to_formulas': False}}) as writer:
            try:
                df = pd.read_sql('''SELECT * FROM cadastro_ferramentas''', conexao)
                df.to_excel(writer, sheet_name='Ferramentas', header=True, index=False)
                print("Arquivo salvo com sucesso!")
                messagebox.showinfo('Exportação', 'Dados exportados para o Excel com sucesso!')
            except:
                print("Aconteceu um erro")

    def tecnicoExportarExcel(self):
        conexao = self.conexaoFerramentas = sqlite3.connect('ferramentas_clientes.db')
        with pd.ExcelWriter("banco_técnico.xlsx", engine="xlsxwriter",
                            engine_kwargs={
                                'options': {'strings_to_numbers': True, 'strings_to_formulas': False}}) as writer:
            try:
                df = pd.read_sql('''SELECT * FROM cadastro_tecnico''', conexao)
                df.to_excel(writer, sheet_name='Técnicos', header=True, index=False)
                print("Arquivo salvo com sucesso!")
                messagebox.showinfo('Exportação', 'Dados exportados para o Excel com sucesso!')
            except:
                print("Aconteceu um erro")

    def solicitacaoExportarExcel(self):
        conexao = self.conexaoFerramentas = sqlite3.connect('ferramentas_clientes.db')
        with pd.ExcelWriter("banco_solicitação.xlsx", engine="xlsxwriter",
                            engine_kwargs={'options':{'strings_to_numbers': True, 'strings_to_formulas': False}}) as writer:
            try:
                df = pd.read_sql('''SELECT * FROM cadastro_solicitação''', conexao)
                df.to_excel(writer, sheet_name='Solicitação', header=True, index=False)
                print("Arquivo salvo com sucesso!")
                messagebox.showinfo('Exportação', 'Dados exportados para o Excel com sucesso!')
            except:
                print("Aconteceu um erro")

#criando conexão com gmail

    def enviarEmail(self):
        assunto = 'Nova solicitação de reserva de ferramenta!!!'
        corpo = 'A tabela com as solicitação está em anexo.'
        password = 'zcludfnvpsdkoqzs'
        email_de = 'projetofinalm1@gmail.com'
        email_para = 'projetofinalm1@gmail.com'

        mensagem = MIMEMultipart()
        mensagem["From"] = email_de
        mensagem["To"] = email_para
        mensagem["Subject"] = assunto

        mensagem.attach(MIMEText(corpo))

        part = MIMEBase("application", "octet-stream")
        part.set_payload(open('banco_solicitação.xlsx', 'rb').read())
        encoders.encode_base64(part)

        part.add_header('Content-Disposition', 'attachment', filename='banco_solicitação.xlsx')
        mensagem.attach(part)


        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
            server.login(email_de, password)
            server.sendmail(email_de, email_para, mensagem.as_string())
        messagebox.showinfo('ENVIADO', message='Solicitação confirmada e enviada por email com sucesso!!!')


#janela login

login = tk.Tk()
login.title('Tela de Login')
login.geometry('800x600')


# Janela principal TKINTER
def login_janela():
    usuario = ['admin', 'admin']
    #condição para login
    if senha_entry.get() == usuario[1]:
        #login.withdraw()
        janela = tk.Toplevel()
        principal = janela_ferramentas(janela)
        janela.title('Cadastro de Ferramentas')
        style = ttk.Style(janela)
        style.theme_use("clam")
        janela.geometry("1100x600")
        janela.mainloop()
        login.destroy()
    else:
        messagebox.showwarning('Login', message='Usuário ou senha incorretos, realize o login novamente.')
        usuario_entry.delete(0,'end')
        senha_entry.delete(0, 'end')

#criando interface login

frame_titulo = tk.Frame(login, width=800, height=100, bg=co11, relief='raised')
frame_titulo.grid(row=0, column=0)

frame_principal = tk.Frame(login, width=800, height=500, bg=co12, relief='flat')
frame_principal.grid(row=1, column=0)

titulo_login = tk.Label(frame_titulo, text='CENTRAL DE FERRAMENTAS', height=1, anchor='nw', font=('Ivy 34 bold'),
                        bg=co11, fg='#FFFF00')
titulo_login.place(x=67, y=30)

corpo_label = tk.Label(frame_principal, text='ENTRE COM USUÁRIO E SENHA', font=('Ivy 20 bold'), fg='#292929',
                       bg=co12)
corpo_label.place(x=180, y=80)

usuario_label = tk.Label(frame_principal, text='USUÁRIO:', font=('Ivy 16 bold'), fg='#292929', bg=co12)
usuario_label.place(x=290, y=130)

usuario_entry = tk.Entry(frame_principal, font=('Ivy 10 bold'), width=30, borderwidth=2, border=2)
usuario_entry.place(x=290, y=160)

senha_label = tk.Label(frame_principal, text='SENHA:', font=('Ivy 16 bold'), fg='#292929', bg=co12)
senha_label.place(x=290, y=200)

senha_entry = tk.Entry(frame_principal, font=('Ivy 10 bold'), width=30, show='*', borderwidth=2, border=2)
senha_entry.place(x=290, y=230)

#janela botão

botao_login = tk.Button(login, text='Entrar', width=18, font=("Ivy 10 bold"), fg=co1,
                        bg=co11, relief='raised', overrelief='ridge', command=login_janela)
botao_login.place(x=290, y=400)

botao_sair = tk.Button(login, text='Sair', width=18, font=("Ivy 10 bold"), fg=co1,
                       bg=co11, relief='raised', overrelief='ridge', command=login.destroy)
botao_sair.place(x=290, y=460)

login.mainloop()






