import time
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from awesometkinter import *
import awesometkinter as atk
import tkinter as tk
from tkinter import filedialog
import sqlite3
import pandas as pd
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
import re
import os
import shutil
import win32com.client as client
from datetime import date, timedelta, datetime
import pywhatkit as kt
import webbrowser


lista = pd.read_excel('CERTIFICADOS.xlsx', engine='openpyxl')
listaxls = lista['EMPRESAS']
lista_tipos = list()
for i in listaxls:
    lista_tipos.append(i)
list_status = [' ', 'A RECEBER', 'RECEBIDO']
#primeiro dia do mês passado
mesrf = '11'
mes = date.today()
mes = mes.strftime('%m')
if mes >= mesrf:
    udmp = date.today().replace(day=1) - timedelta(days=1)
    udmpf = '{}/{}/{}'.format(udmp.day, udmp.month, udmp.year)
    #ultimo dia do mês passado
    pdmp = date.today().replace(day=1) - timedelta(days=udmp.day)
    pdmpf = '0{}/{}/{}'.format(pdmp.day, pdmp.month, pdmp.year)
else:
    udmp = date.today().replace(day=1) - timedelta(days=1)
    udmpf = '{}/0{}/{}'.format(udmp.day, udmp.month, udmp.year)
    #ultimo dia do mês passado
    pdmp = date.today().replace(day=1) - timedelta(days=udmp.day)
    pdmpf = '0{}/0{}/{}'.format(pdmp.day, pdmp.month, pdmp.year)

dia = 25
hoje = datetime.now()
mes1 = hoje.month
ano = hoje.year
venc = datetime(day=dia, month=mes1, year=ano)
venc = venc.strftime("%d/%m/%Y")
if mes1 == 1:
    ref = 12
    ano = ano - 1
else:
    ref = mes1 - 1

ref = datetime(day=dia,month=ref, year=ano).strftime("%m/%Y")
# decide qual caminho serar usado para salvaer as nfs de fora
caml = os.path.dirname(os.path.realpath(__file__))
cam = r'F:/meus programas/MENESES CONT 2'
caml = caml.replace('\\', '/')

if cam == caml:
    caminho = cam
else:
    caminho = r'\\Desktop-1v4f59g\meneses cont 2'



janela = Tk()

class Validadores:
    def validadores_entry(self, text):
        if text == "": return True
        try:
            value = int(text)
        except ValueError:
            return False
        return 0 <= value <= 100

class Funcs():
    def limpa_cliente(self):
        self.codigo_entry.delete(0, END)
        self.cidade_entry.delete(0, END)
        self.fone_entry.delete(0, END)
        self.nome_entry.delete(0, END)
        self.cel_entry.delete(0, END)
    def conecta_bd(self):
        self.conn = sqlite3.connect("clientes.db")
        self.cursor = self.conn.cursor()
        print("Conectando ao banco de dados")
    def desconecta_bd(self):
        self.conn.close()
        print("Desconectando ao banco de dados")
    def montaTabelas(self):
        self.conecta_bd()
        ### Criar tabela
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS clientes (
                cod INTEGER PRIMARY KEY,
                nome_cliente CHAR(40) NOT NULL,
                celular INTEGER(20),
                telefone INTEGER(20),
                email CHAR(40)               
            );
        """)
        self.conn.commit()
        print("Banco de dados criado")
        self.desconecta_bd()
    def variaveis(self):
        self.codigo = self.codigo_entry.get()
        self.nome = self.nome_entry.get()
        self.cel = self.cel_entry.get()
        self.fone = self.fone_entry.get()
        self.cidade = self.cidade_entry.get()
    def OnDoubleClick(self, event):
        self.limpa_cliente()

        for n in self.listaCli.selection():
            col1, col2, col3, col4, col5 = self.listaCli.item(n, 'values')
            self.codigo_entry.insert(END, col1)
            self.nome_entry.insert(END, col2)
            self.cel_entry.insert(END, col3)
            self.fone_entry.insert(END, col4)
            self.cidade_entry.insert(END, col5)
    def add_cliente(self):
        self.variaveis()
        self.conecta_bd()

        self.cursor.execute(""" INSERT INTO clientes (nome_cliente, celular, telefone, email)
            VALUES (?, ?, ?, ?)""", (self.nome, self.cel, self.fone, self.cidade))
        if self.nome == self.fone == self.cidade:
            tk.messagebox.showinfo(title='Informação', message='Todos os campos devem ser preenchidos !!')
        else:
            self.conn.commit()
            self.desconecta_bd()
            self.select_lista()
            self.limpa_cliente()
    def alterar_cliente(self):
        self.variaveis()
        self.conecta_bd()
        self.cursor.execute(""" UPDATE clientes SET nome_cliente = ?, celular = ?, telefone = ?, email = ?
                 WHERE cod = ? """, (self.nome, self.cel, self.fone, self.cidade, self.codigo))
        self.conn.commit()
        self.desconecta_bd()
        self.select_lista()
        self.limpa_cliente()

    def deleta_cliente(self):
        self.variaveis()
        self.conecta_bd()
        self.cursor.execute("""DELETE FROM clientes WHERE cod = ? """, (self.codigo))
        self.conn.commit()
        self.desconecta_bd()
        self.limpa_cliente()
        self.select_lista()
    def select_lista(self):
        self.listaCli.delete(*self.listaCli.get_children())
        self.conecta_bd()
        lista = self.cursor.execute(""" SELECT cod, nome_cliente, celular, telefone, email FROM clientes
            ORDER BY nome_cliente ASC; """)
        for i in lista:
            self.listaCli.insert("", END, values=i)
        self.desconecta_bd()
    def busca_cliente(self):
        self.conecta_bd()
        self.listaCli.delete(*self.listaCli.get_children())

        self.nome_entry.insert(END, '%')
        nome = self.nome_entry.get()
        self.cursor.execute(
            """ SELECT cod, nome_cliente, celular,  telefone, email FROM clientes
            WHERE nome_cliente LIKE '%s' ORDER BY nome_cliente ASC""" % nome)
        buscanomeCli = self.cursor.fetchall()
        for i in buscanomeCli:
            self.listaCli.insert("", END, values=i)
        self.limpa_cliente()
        self.desconecta_bd()
class Funcs2():
    def limpa_empresa(self):
        self.codigo1_entry.delete(0, END)
        self.cpf_entry.delete(0, END)
        self.cnpj_entry.delete(0, END)
        self.empresa_entry.delete(0, END)
        self.senha_entry.delete(0, END)
    def conecta1_bd(self):
        self.conn = sqlite3.connect("Empresas_SN.db")
        self.cursor1 = self.conn.cursor()
        print("Conectando ao banco de dados")
    def desconecta1_bd(self):
        self.conn.close()
        print("Desconectando ao banco de dados")
    def montaTabelas1(self):
        self.conecta1_bd()
        ### Criar tabela
        self.cursor1.execute("""
            CREATE TABLE IF NOT EXISTS Empresas_SN (
                cod INTEGER PRIMARY KEY,
                nome_Empresa CHAR(40) NOT NULL,
                CNPJ BLOB(20),
                CPF BLOB(20),
                Senha BLOB(20)              
            );
        """)
        self.conn.commit()
        print("Banco de dados criado")
        self.desconecta1_bd()
    def variaveis1(self):
        self.codigo1 = self.codigo1_entry.get()
        self.empresa = self.empresa_entry.get()
        self.cnpj = self.cnpj_entry.get()
        self.cpf = self.cpf_entry.get()
        self.senha = self.senha_entry.get()
    def OnDoubleClick1(self, event):
        self.limpa_empresa()
        self.listaEMP.selection()

        for n in self.listaEMP.selection():
            col1, col2, col3, col4, col5 = self.listaEMP.item(n, 'values')
            self.codigo1_entry.insert(END, col1)
            self.empresa_entry.insert(END, col2)
            self.cnpj_entry.insert(END, col3)
            self.cpf_entry.insert(END, col4)
            self.senha_entry.insert(END, col5)
    def add_empresa(self):
        self.variaveis1()
        self.conecta1_bd()

        self.cursor1.execute(""" INSERT INTO Empresas_SN (nome_Empresa, CNPJ, CPF, Senha)
            VALUES (?, ?, ?, ?)""", (self.empresa, self.cnpj, self.cpf, self.senha))
        if self.empresa == self.cnpj == self.cpf == self.senha:
            tk.messagebox.showinfo(title='Informação', message='Todos os campos devem ser preenchidos !!')
        else:
            self.conn.commit()
            self.desconecta1_bd()
            self.select_lista1()
            self.limpa_empresa()
    def altera_empresa(self):
        self.variaveis1()
        self.conecta1_bd()
        self.cursor1.execute(""" UPDATE Empresas_SN SET nome_Empresa = ?, CNPJ = ?, CPF = ?, Senha = ?
            WHERE cod = ? """,
                             (self.empresa, self.cnpj, self.cpf, self.senha, self.codigo1))
        self.conn.commit()
        self.desconecta1_bd()
        self.select_lista1()
        self.limpa_empresa()
    def deleta_empresa(self):
        self.variaveis1()
        self.conecta1_bd()
        self.cursor1.execute("""DELETE FROM Empresas_SN WHERE cod = ? """, (self.codigo1))
        self.conn.commit()
        self.desconecta1_bd()
        self.limpa_empresa()
        self.select_lista1()
    def select_lista1(self):
        self.listaEMP.delete(*self.listaEMP.get_children())
        self.conecta1_bd()
        lista = self.cursor1.execute(""" SELECT cod, nome_Empresa, CNPJ, CPF, Senha FROM Empresas_SN
            ORDER BY nome_Empresa ASC; """)
        for i in lista:
            self.listaEMP.insert("", END, values=i)
        self.desconecta1_bd()
    def busca_empresa(self):
        self.conecta1_bd()
        self.listaEMP.delete(*self.listaEMP.get_children())

        self.empresa_entry.insert(END, '%')
        nome = self.empresa_entry.get()
        self.cursor1.execute(
            """ SELECT cod, nome_Empresa, CNPJ, CPF, Senha FROM Empresas_SN
            WHERE nome_Empresa LIKE '%s' ORDER BY nome_Empresa ASC""" % nome)
        buscanomeCli = self.cursor1.fetchall()
        for i in buscanomeCli:
            self.listaEMP.insert("", END, values=i)
        self.limpa_empresa()
        self.desconecta1_bd()
class Funcs3():
    def limpa_empresa_sf(self):
        self.codigo2_entry.delete(0, END)
        self.usuario_entry.delete(0, END)
        self.empresa2_entry.delete(0, END)
        self.senha2_entry.delete(0, END)
    def conecta2_bd(self):
        self.conn = sqlite3.connect("Empresas_SF.db")
        self.cursor2 = self.conn.cursor()
        print("Conectando ao banco de dados")
    def desconecta2_bd(self):
        self.conn.close()
        print("Desconectando ao banco de dados")
    def montaTabelas2(self):
        self.conecta2_bd()
        ### Criar tabela
        self.cursor2.execute("""
            CREATE TABLE IF NOT EXISTS Empresas_SF (
                cod INTEGER PRIMARY KEY,
                nome_Empresa CHAR(40) NOT NULL,
                Usuario BLOB(20),               
                Senha BLOB(20)              
            );
        """)
        self.conn.commit()
        print("Banco de dados criado")
        self.desconecta2_bd()
    def variaveis2(self):
        self.codigo2 = self.codigo2_entry.get()
        self.empresa2 = self.empresa2_entry.get()
        self.usuario = self.usuario_entry.get()
        self.senha2 = self.senha2_entry.get()
    def OnDoubleClick2(self, event):
        self.limpa_empresa_sf()
        self.listaEMP.selection()

        for n in self.listaEMP.selection():
            col1, col2, col3, col4 = self.listaEMP.item(n, 'values')
            self.codigo2_entry.insert(END, col1)
            self.empresa2_entry.insert(END, col2)
            self.usuario_entry.insert(END, col3)
            self.senha2_entry.insert(END, col4)
    def add_empresa_sf(self):
        self.variaveis2()
        self.conecta2_bd()

        self.cursor2.execute(""" INSERT INTO Empresas_SF (nome_Empresa, Usuario, Senha)
            VALUES (?, ?, ?)""", (self.empresa2, self.usuario, self.senha2))
        if self.empresa2 == self.usuario == self.senha2:
            tk.messagebox.showinfo(title='Informação', message='Todos os campos devem ser preenchidos !!')
        else:
            self.conn.commit()
            self.desconecta2_bd()
            self.select_lista2()
            self.limpa_empresa_sf()
    def altera_empresa_sf(self):
        self.variaveis2()
        self.conecta2_bd()
        self.cursor2.execute(""" UPDATE Empresas_SF SET nome_Empresa = ?, Usuario = ?, Senha = ?
            WHERE cod = ? """,
                             (self.empresa2, self.usuario, self.senha2, self.codigo2))
        self.conn.commit()
        self.desconecta2_bd()
        self.select_lista2()
        self.limpa_empresa_sf()
    def deleta_empresa_sf(self):
        self.variaveis2()
        self.conecta2_bd()
        self.cursor2.execute("""DELETE FROM Empresas_SF WHERE cod = ? """, (self.codigo2))
        self.conn.commit()
        self.desconecta2_bd()
        self.limpa_empresa_sf()
        self.select_lista2()
    def select_lista2(self):
        self.listaEMP.delete(*self.listaEMP.get_children())
        self.conecta2_bd()
        lista = self.cursor2.execute(""" SELECT cod, nome_Empresa, Usuario, Senha FROM Empresas_SF
            ORDER BY nome_Empresa ASC; """)
        for i in lista:
            self.listaEMP.insert("", END, values=i)
        self.desconecta2_bd()
    def busca_empresa_sf(self):
        self.conecta2_bd()
        self.listaEMP.delete(*self.listaEMP.get_children())

        self.empresa2_entry.insert(END, '%')
        nome = self.empresa2_entry.get()
        self.cursor2.execute(
            """ SELECT cod, nome_Empresa, Usuario, Senha FROM Empresas_SF
            WHERE nome_Empresa LIKE '%s' ORDER BY nome_Empresa ASC""" % nome)
        buscanomeCli = self.cursor2.fetchall()
        for i in buscanomeCli:
            self.listaEMP.insert("", END, values=i)
        self.limpa_empresa_sf()
        self.desconecta2_bd()
class Funcs4():
    def limpa_empresa_ISS(self):
        self.codigo3_entry.delete(0, END)
        self.login_entry.delete(0, END)
        self.empresa3_entry.delete(0, END)
        self.senha3_entry.delete(0, END)
    def conecta3_bd(self):
        self.conn = sqlite3.connect("Empresas_ISS.db")
        self.cursor3 = self.conn.cursor()
        print("Conectando ao banco de dados")
    def desconecta3_bd(self):
        self.conn.close()
        print("Desconectando ao banco de dados")
    def montaTabelas3(self):
        self.conecta3_bd()
        ### Criar tabela
        self.cursor3.execute("""
            CREATE TABLE IF NOT EXISTS Empresas_ISS (
                cod INTEGER PRIMARY KEY,
                nome_Empresa CHAR(40) NOT NULL,
                Login BLOB(20),               
                Senha BLOB(20)              
            );
        """)
        self.conn.commit()
        print("Banco de dados criado")
        self.desconecta3_bd()
    def variaveis3(self):
        self.codigo3 = self.codigo3_entry.get()
        self.empresa3 = self.empresa3_entry.get()
        self.login = self.login_entry.get()
        self.senha3 = self.senha3_entry.get()
    def OnDoubleClick3(self, event):
        self.limpa_empresa_ISS()
        self.listaEMP.selection()

        for n in self.listaEMP.selection():
            col1, col2, col3, col4 = self.listaEMP.item(n, 'values')
            self.codigo3_entry.insert(END, col1)
            self.empresa3_entry.insert(END, col2)
            self.login_entry.insert(END, col3)
            self.senha3_entry.insert(END, col4)
    def add_empresa_ISS(self):
        self.variaveis3()
        self.conecta3_bd()

        self.cursor3.execute(""" INSERT INTO Empresas_ISS (nome_Empresa, Login, Senha)
            VALUES (?, ?, ?)""",
                             (self.empresa3, self.login, self.senha3))
        if self.empresa3 == self.login == self.senha3:
            tk.messagebox.showinfo(title='Informação', message='Todos os campos devem ser preenchidos !!')
        else:
            self.conn.commit()
            self.desconecta3_bd()
            self.select_lista3()
            self.limpa_empresa_ISS()
    def altera_empresa_ISS(self):
        self.variaveis3()
        self.conecta3_bd()
        self.cursor3.execute(""" UPDATE Empresas_ISS SET nome_Empresa = ?, Login = ?, Senha = ?
            WHERE cod = ? """, (self.empresa3, self.login, self.senha3, self.codigo3))
        self.conn.commit()
        self.desconecta3_bd()
        self.select_lista3()
        self.limpa_empresa_ISS()
    def deleta_empresa_ISS(self):
        self.variaveis3()
        self.conecta3_bd()
        self.cursor3.execute("""DELETE FROM Empresas_ISS WHERE cod = ? """, (self.codigo3))
        self.conn.commit()
        self.desconecta3_bd()
        self.limpa_empresa_ISS()
        self.select_lista3()
    def select_lista3(self):
        self.listaEMP.delete(*self.listaEMP.get_children())
        self.conecta3_bd()
        lista = self.cursor3.execute(""" SELECT cod, nome_Empresa, Login, Senha FROM Empresas_ISS
            ORDER BY nome_Empresa ASC; """)
        for i in lista:
            self.listaEMP.insert("", END, values=i)
        self.desconecta3_bd()
    def busca_empresa_ISS(self):
        self.conecta3_bd()
        self.listaEMP.delete(*self.listaEMP.get_children())

        self.empresa3_entry.insert(END, '%')
        nome = self.empresa3_entry.get()
        self.cursor3.execute(
            """ SELECT cod, nome_Empresa, Login, Senha FROM Empresas_ISS
            WHERE nome_Empresa LIKE '%s' ORDER BY nome_Empresa ASC""" % nome)
        buscanomeCli = self.cursor3.fetchall()
        for i in buscanomeCli:
            self.listaEMP.insert("", END, values=i)
        self.limpa_empresa_ISS()
        self.desconecta3_bd()
class Funcs5():
    def limpa_solicitar_arquivo(self):
        self.empresa4_entry.delete(0, END)
        self.email_entry.delete(0, END)
        self.obs_entry.delete(0, END)
        self.data_entry.delete(0, END)
        self.tipo_arquivo_entry.delete(0, END)
        self.status_entry.delete(0, END)
        self.codigo4_entry.delete(0, END)

    def conecta4_bd(self):
        self.conn = sqlite3.connect("Solicitar_arquivo.db")
        self.cursor4 = self.conn.cursor()
        print("Conectando ao banco de dados")

    def desconecta4_bd(self):
        self.conn.close()
        print("Desconectando ao banco de dados")

    def montaTabelas4(self):
        self.conecta4_bd()
        ### Criar tabela
        self.cursor4.execute("""
            CREATE TABLE IF NOT EXISTS Solicitar_arquivo (
                cod INTEGER PRIMARY KEY,
                EMPRESA CHAR(40) NOT NULL,
                TIPO INTEGER(20),               
                DATA INTEGER(20),
                STATUS INTEGER(20),               
                EMAIL INTEGER(30),
                OBSERVAÇÃO INTEGER(20)                 
            );
        """)
        self.conn.commit()
        print("Banco de dados criado")
        self.desconecta4_bd()

    def variaveis4(self):
        self.codigo4 = self.codigo4_entry.get()
        self.empresa4 = self.empresa4_entry.get()
        self.email = self.email_entry.get()
        self.obs = self.obs_entry.get()
        self.data = self.data_entry.get()
        self.tipo_arquivo = self.tipo_arquivo_entry.get()
        self.status = self.status_entry.get()

    def OnDoubleClick4(self, event):
        self.limpa_solicitar_arquivo()
        self.listaEMP.selection()

        for n in self.listaEMP.selection():
            col1, col2, col3, col4, col5, col6, col7 = self.listaEMP.item(n, 'values')
            self.codigo4_entry.insert(END, col1)
            self.empresa4_entry.insert(END, col2)
            self.tipo_arquivo_entry.insert(END, col3)
            self.data_entry.insert(END, col4)
            self.status_entry.insert(END, col5)
            self.email_entry.insert(END, col6)
            self.obs_entry.insert(END, col7)

    def add_solicitar_arquivo(self):
        self.variaveis4()
        self.conecta4_bd()

        self.cursor4.execute(""" INSERT INTO Solicitar_arquivo (EMPRESA, TIPO, DATA, STATUS, EMAIL, OBSERVAÇÃO)
            VALUES (?, ?, ?, ?, ?, ?)""",
                             (self.empresa4, self.tipo_arquivo, self.data, self.status, self.email, self.obs))
        if self.empresa4 == self.tipo_arquivo == self.data == self.status == self.email == self.obs:
            tk.messagebox.showinfo(title='Informação', message='Todos os campos devem ser preenchidos !!')
        else:
            self.cursor4.execute("SELECT *, oid FROM Solicitar_arquivo")
            self.banco_cliente = self.cursor4.fetchall()
            self.banco_cliente = pd.DataFrame(self.banco_cliente, columns=('CODIGO','EMPRESA','TIPO','DATA','STATUS', 'EMAIL', 'OBSERVAÇÃO', 'ID'))
            self.banco_cliente.to_excel('Cobranca_Cliante.xlsx')
            self.conn.commit()
            self.desconecta4_bd()
            self.select_lista4()
            self.limpa_solicitar_arquivo()

    def altera_solicitar_arquivo(self):
        self.variaveis4()
        self.conecta4_bd()
        self.cursor4.execute(""" UPDATE Solicitar_arquivo SET EMPRESA = ?, TIPO = ?, DATA = ?, STATUS = ?, EMAIL  = ?, OBSERVAÇÃO = ?
            WHERE cod = ? """, (
        self.empresa4, self.tipo_arquivo, self.data, self.status, self.email, self.obs, self.codigo4))
        self.cursor4.execute("SELECT *, oid FROM Solicitar_arquivo")
        self.banco_cliente = self.cursor4.fetchall()
        self.banco_cliente = pd.DataFrame(self.banco_cliente, columns=(
        'CODIGO', 'EMPRESA', 'TIPO', 'DATA', 'STATUS', 'EMAIL', 'OBSERVAÇÃO', 'ID'))
        self.banco_cliente.to_excel('Cobranca_Cliante.xlsx')
        self.conn.commit()
        self.conn.commit()
        self.desconecta4_bd()
        self.select_lista4()
        self.limpa_solicitar_arquivo()

    def deleta_solicitar_arquivo(self):
        self.variaveis4()
        self.conecta4_bd()
        self.cursor4.execute("""DELETE FROM Solicitar_arquivo WHERE cod = ? """, (self.codigo4))
        self.conn.commit()
        self.desconecta4_bd()
        self.limpa_solicitar_arquivo()
        self.select_lista4()

    def select_lista4(self):
        self.listaEMP.delete(*self.listaEMP.get_children())
        self.conecta4_bd()
        lista = self.cursor4.execute(""" SELECT cod, EMPRESA, TIPO, DATA, STATUS, EMAIL, OBSERVAÇÃO FROM Solicitar_arquivo
            ORDER BY EMPRESA ASC; """)
        for i in lista:
            self.listaEMP.insert("", END, values=i)
        self.desconecta4_bd()

    def busca_solicitar_arquivo(self):
        self.conecta4_bd()
        self.listaEMP.delete(*self.listaEMP.get_children())

        self.empresa4_entry.insert(END, '%')
        nome = self.empresa4_entry.get()
        self.cursor4.execute(
            """ SELECT cod, EMPRESA, TIPO, DATA, STATUS, EMAIL, OBSERVAÇÃO FROM Solicitar_arquivo
            WHERE EMPRESA LIKE '%s' ORDER BY EMPRESA ASC""" % nome)
        buscanomeCli = self.cursor4.fetchall()
        for i in buscanomeCli:
            self.listaEMP.insert("", END, values=i)
        self.limpa_solicitar_arquivo()
        self.desconecta4_bd()

    def solicita_arquivo_email(self):
        self.tabela = pd.read_excel('Cobranca_Cliante.xlsx')
        self.hoje = date.today()
        self.hoje.strftime('%d')
        self.hoje = '0{}'.format(self.hoje.day)
        print(self.hoje)
        self.tabela_devedores = self.tabela.loc[self.tabela['STATUS'] == 'A RECEBER']
        print(self.tabela_devedores)
        self.tabela_devedores.to_excel('tabela_devedores.xlsx')
        self.tabela2 = (self.tabela_devedores)
        self.tabela_areceber = (self.tabela2.loc[self.tabela2['DATA'] <= int(self.hoje)])
        print(self.tabela_areceber)
        self.outlook = client.Dispatch('Outlook.Application')
        self.emissor = self.outlook.session.Accounts['menesescontabil@outlook.com']
        self.dados = self.tabela_areceber[['EMPRESA', 'TIPO', 'DATA', 'EMAIL', 'OBSERVAÇÃO']].values.tolist()
        print(self.dados)
        for dado in self.dados:
            self.destinatario = dado[3]
            self.obs = dado[4]
            self.prazo = dado[2]
            self.cl = dado[0]
            # prazo = prazo.strftime("%d")
            self.tipo = dado[1]
            self.assunto = 'Relatorio de Vendas e Despesas'
            self.mensagem = self.outlook.CreateItem(0)
            self.mensagem.display
            self.mensagem.To = self.destinatario
            self.mensagem.Subject = self.destinatario
            self.corpo_mensagem = f'''
            Prezado Cliente,

            Verificamos que o Arquivo {self.tipo} da empresa {self.cl} Ainda não foi enviado para a contabilidade 
            Gostaríamos de verificar se há algum problema que necessite de auxílio.

            Em caso de dúvidas, é só entrar em contato 


            Att,
           Meneses contabilidade
            '''
            self.mensagem.Body = self.corpo_mensagem
            self.mensagem._oleobj_.Invoke(*(64209, 0, 8, 0, self.emissor))
            self.mensagem.Save()
            self.mensagem.Send()
class web_sn():
    def login_simples(self):
        cnpj = self.cnpj_entry.get()
        cpf = self.cpf_entry.get()
        cod_acesso = self.senha_entry.get()

        # self.options = webdriver.ChromeOptions()
        # self.options.add_argument("--start-maximized")
        # self.nav = webdriver.Chrome(options=self.options)
        # self.nav.implicitly_wait(10)
        # self.nav.get('https://www8.receita.fazenda.gov.br/SimplesNacional/controleAcesso/Autentica.aspx?id=60')
        # self.nav.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder_txtCNPJ"]').send_keys(self.cnpj_entry.get())
        # self.nav.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder_txtCPFResponsavel"]').send_keys(self.cpf_entry.get())
        # self.nav.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder_txtCodigoAcesso"]').send_keys(self.senha_entry.get())
        webbrowser.open("https://www8.receita.fazenda.gov.br/SimplesNacional/controleAcesso/Autentica.aspx?id=60")
        time.sleep(4)
        pyautogui.keyDown('shift')
        pyautogui.press('tab', presses=10)
        pyautogui.keyUp('shift')
        time.sleep(2)
        pyautogui.write(cnpj, interval=0.25)
        time.sleep(2)
        pyautogui.write(cpf, interval=0.25)
        time.sleep(2)
        pyautogui.write(cod_acesso, interval=0.25)
    def login_simples_parcelamento(self):
        cnpj = self.cnpj_entry.get()
        cpf = self.cpf_entry.get()
        cod_acesso = self.senha_entry.get()
        # self.options = webdriver.ChromeOptions()
        # self.options.add_argument("--start-maximized")
        # self.nav = webdriver.Chrome(options=self.options)
        # self.nav.implicitly_wait(10)
        # self.nav.get('https://www8.receita.fazenda.gov.br/SimplesNacional/controleAcesso/Autentica.aspx?id=37')
        # self.nav.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder_txtCNPJ"]').send_keys(self.cnpj_entry.get())
        # self.nav.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder_txtCPFResponsavel"]').send_keys(self.cpf_entry.get())
        # self.nav.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder_txtCodigoAcesso"]').send_keys(self.senha_entry.get())
        webbrowser.open("https://www8.receita.fazenda.gov.br/SimplesNacional/controleAcesso/Autentica.aspx?id=37")
        time.sleep(4)
        pyautogui.keyDown('shift')
        pyautogui.press('tab', presses=10)
        pyautogui.keyUp('shift')
        time.sleep(2)
        pyautogui.write(cnpj, interval=0.25)
        time.sleep(2)
        pyautogui.write(cpf, interval=0.25)
        time.sleep(2)
        pyautogui.write(cod_acesso, interval=0.25)
    def consultar_cnpj(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.implicitly_wait(30)
        self.nav.get('https://www.sintegraws.com.br/entrar')
        self.nav.find_element(By.XPATH, '//*[@id="usuario"]').send_keys('menesescontabil@outlook.com')
        self.nav.find_element(By.XPATH, '//*[@id="password"]').send_keys('Lima1234@')
        self.nav.find_element(By.XPATH, '//*[@id="login-form"]/div/div[3]/div[4]/div/button').click()
        self.nav.find_element(By.XPATH,
                         '/html/body/div[1]/div[2]/div[1]/div/div[1]/div/div/div/div/div[2]/div/div/div[1]/input').send_keys(self.cnpj_entry.get())

        self.nav.find_element(By.XPATH, '//*[@id="nav-receita-tab"]').click()
    def login_sefaz_entrada(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.options.add_argument('ignore-certificate-errors')
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.implicitly_wait(10)
        self.nav.get(
            'https://nfe.sefaz.ba.gov.br/servicos/NFENC/SSL/ASLibrary/Login?ReturnUrl=%2fservicos%2fnfenc%2fModulos%2fAutenticado%2fRestrito%2fNFENC_consulta_destinatario.aspx')
        #self.nav.find_element(By.XPATH, '//*[@id="details-button"]').click()
        #self.nav.find_element(By.XPATH, '//*[@id="proceed-link"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="PHCentro_userLogin"]').send_keys(self.usuario_entry.get())
        self.nav.find_element(By.XPATH, '//*[@id="PHCentro_userPass"]').send_keys(self.senha2_entry.get())
        pyautogui.press('enter')
        time.sleep(3)
        pyautogui.press(['tab', 'tab', 'tab', 'enter'])
        self.nav.find_element(By.XPATH, '//*[@id="rbt_filtro3"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoInicial"]').send_keys(pdmpf)
        time.sleep(1)
        if self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoInicial"]').text == "":
            self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoInicial"]').send_keys(pdmpf)
        self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoFinal"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoFinal"]').send_keys(udmpf)
        if self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoFinal"]').text == "":
            self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoFinal"]').send_keys(udmpf)
        time.sleep(0.5)
        self.nav.find_element(By.XPATH, '//*[@id="AplicarFiltro"]').click()
    def login_sefaz_saida(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.implicitly_wait(10)
        self.nav.get(
            'https://nfe.sefaz.ba.gov.br/servicos/NFENC/SSL/ASLibrary/Login?ReturnUrl=%2fservicos%2fnfenc%2fModulos%2fAutenticado%2fRestrito%2fNFENC_consulta_emitente.aspx')
        self.nav.find_element(By.XPATH, '//*[@id="details-button"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="proceed-link"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="PHCentro_userLogin"]').send_keys(self.usuario_entry.get())
        self.nav.find_element(By.XPATH, '//*[@id="PHCentro_userPass"]').send_keys(self.senha2_entry.get())
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.press(['tab', 'tab', 'tab', 'enter'])
        self.nav.find_element(By.XPATH, '//*[@id="rbt_filtro3"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoInicial"]').send_keys(pdmpf)
        time.sleep(1)
        if self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoInicial"]').text == "":
            self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoInicial"]').send_keys(pdmpf)
        self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoFinal"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoFinal"]').send_keys(udmpf)
        if self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoFinal"]').text == "":
            self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoFinal"]').send_keys(udmpf)
        time.sleep(0.5)
        self.nav.find_element(By.XPATH, '//*[@id="AplicarFiltro"]').click()
    def Gestao_iss(self):
        login = '00693222522'
        senha = 'LIMA2020'
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.implicitly_wait(10)
        self.nav.get('https://portoseguroba.gestaoiss.com.br/')
        self.nav.find_element(By.XPATH, '//*[@id="Login"]').send_keys(login)
        self.nav.find_element(By.XPATH, '//*[@id="Senha"]').send_keys(senha)
        self.nav.find_element(By.XPATH, '//*[@id="botao-logar"]').click()
    def consultar_cnd(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.implicitly_wait(10)
        self.nav.get('https://portoseguro.saatri.com.br/')
        self.nav.find_element(By.XPATH, '//*[@id="Pint_TipoCertidao_chosen"]/a/div/b').click()
        pyautogui.press('tab')
        self.nav.find_element(By.XPATH, '// *[ @ id = "txt_CpfCnpjCertidao"]').send_keys(self.cnpj_entry.get())
        self.nav.find_element(By.XPATH, '//*[@id="btn_ConsultarContribuinteCnd"]').click(self.cnpj_entry.get())
        self.nav.find_element(By.XPATH, '//*[@id="680112"]/a').click()
    def consultar_insc_est(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.implicitly_wait(10)
        self.nav.get('https://www.sefaz.ba.gov.br/scripts/cadastro/cadastroBa/consultaBa.asp')
        self.nav.find_element(By.XPATH, '/html/body/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]/p/input').send_keys(self.cnpj_entry.get())
        self.nav.find_element(By.XPATH, '/html/body/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[4]/input').click()
    def empregador_digital_aviso(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.implicitly_wait(30)
        self.nav.get('http://sd.mte.gov.br/sdweb/empregadorweb/index.jsf')
        self.nav.find_element(By.XPATH, '//*[@id="f"]/div[1]/a[1]').click()###
        self.nav.find_element(By.XPATH, '//*[@id="details-button"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="proceed-link"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="f:txtLogin"]').send_keys('BORGES_SILVA')
        self.nav.find_element(By.XPATH, '//*[@id="f:txtSenha"]').send_keys('2160275200')
        self.nav.find_element(By.XPATH, '//*[@id="f:enviarLogin"]').click()
    def empregador_digital_sem_aviso(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.implicitly_wait(30)
        self.nav.get('http://sd.mte.gov.br/sdweb/empregadorweb/index.jsf')
        self.nav.find_element(By.XPATH, '//*[@id="f"]/div[1]/a[1]').click()###
        self.nav.find_element(By.XPATH, '//*[@id="f:txtLogin"]').send_keys('BORGES_SILVA')
        self.nav.find_element(By.XPATH, '//*[@id="f:txtSenha"]').send_keys('2160275200')
        self.nav.find_element(By.XPATH, '//*[@id="f:enviarLogin"]').click()
    def controle_rescisao(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.get('https://www.notion.so/pt-br')
        self.nav.implicitly_wait(30)
        self.nav.find_element(By.XPATH, '//*[@id="notion-email-input-1"]').send_keys('menesescontabil@outlook.com')
        #self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[1]/main/section/div/div/div/div[2]/div[3]/form/div[4]').click()
        self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[1]/div/main/div/section/div/div/div/div[2]/div[1]/div[3]/form/div[3]').click()
        self.nav.find_element(By.XPATH, '//*[@id="notion-password-input-2"]').send_keys('lima1234@')
        self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[1]/div/main/div/section/div/div/div/div[2]/div[1]/div[3]/form/div[3]').click()
        #self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[1]/div[1]/div/div/div/div[4]/div[3]/div/div[2]/div').click()
        #self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[2]/div[2]/div/div[2]/div/div[1]/div/div/div[3]/div[2]').click()
        #self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[2]/div[2]/div/div[2]/div/div[2]/div/div/div[5]/div[1]/div[2]').click()
        #self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[2]/div[3]/div/div[2]/div[2]/div/div/div/div/div/div/div/div[3]/div/div/div').click()
        time.sleep(0.5)

        pyautogui.press('esc')
    def consultar_DAM(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.get('https://portoseguro.saatri.com.br/Economico')
        self.nav.implicitly_wait(30)
        self.nav.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/ul/li[1]/a/span[1]').click()
        self.nav.find_element(By.XPATH, '//*[@id="menu-esquerdo"]/li[2]/ul/li/a').click()
        self.nav.find_element(By.XPATH, '//*[@id="txt_CpfCnpj"]').click()
        time.sleep(1)
        pyautogui.press('left', presses=15)
        self.nav.find_element(By.XPATH, '//*[@id="txt_CpfCnpj"]').send_keys(self.cnpj_entry.get())
        self.nav.find_element(By.XPATH, '//*[@id="btn_EmitirTaxaAlvara"]').click()
    def consultar_alvará(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.get('https://portoseguro.saatri.com.br/Empresa/Alvara')
        self.nav.implicitly_wait(30)
        # self.nav.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[2]/ul/li[1]/a/span[1]').click()
        # self.nav.find_element(By.XPATH, '//*[@id="menu-esquerdo"]/li[2]/ul/li/a').click()
        self.nav.find_element(By.XPATH, '//*[@id="txt_CpfCnpj"]').click()
        time.sleep(1)
        pyautogui.press('left', presses=15)
        self.nav.find_element(By.XPATH, '//*[@id="txt_CpfCnpj"]').send_keys(self.cnpj_entry.get())
        self.nav.find_element(By.XPATH, '//*[@id="lnk_EmitirAlvara"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="lnk_EmitirAlvara"]').click()
    def controle_fiscal(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.get('https://www.notion.so/Controle-Simples-Nacional-91513f2a04104bf39559493bbc8a1c0c')
        self.nav.implicitly_wait(30)
        self.nav.find_element(By.XPATH, '//*[@id="notion-email-input-1"]').send_keys('menesescontabil@outlook.com')
        self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[1]/div/div/main/div/div[3]/div[1]/div[3]/form/div[3]').click()
        self.nav.find_element(By.XPATH, '//*[@id="notion-password-input-2"]').send_keys('lima1234@')
        self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[1]/div/div/main/div/div[3]/div[1]/div[3]/form/div[3]').click()
        #self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[1]/div[1]/div/div/div/div[4]/div[3]/div/div[2]/div').click()
        #self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[2]/div[2]/div/div[2]/div/div[1]/div/div/div[3]/div[2]').click()
        #self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[2]/div[2]/div/div[2]/div/div[2]/div/div/div[5]/div[1]/div[2]').click()
        #self.nav.find_element(By.XPATH, '//*[@id="notion-app"]/div/div[2]/div[3]/div/div[2]/div[2]/div/div/div/div/div/div/div/div[3]/div/div/div').click()
    def sendwhats(self):
        self.num = self.cel_entry.get()
        self.cod_pais = '+55'
        if self.num == "":
            tk.messagebox.showinfo(title='Informação', message='O campo do celular esta vazio')
        else:
            kt.sendwhatmsg_instantly(self.cod_pais + self.num, "", wait_time=30)
    def gmail(self):
        self.dest = self.cidade_entry.get()
        if self.dest == "":
            tk.messagebox.showinfo(title='Informação', message='O campo do email esta vazio')
        else:
            webbrowser.open(url='https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox?compose=new')
            time.sleep(10)
            pyautogui.write(self.dest)
            time.sleep(2)
            pyautogui.press(['tab', 'tab'])
    def Emissao_icms(self):

        info = pd.read_excel('InfoEmissaoICMS.xlsx')
        tabela = pd.read_excel('InfoEmissaoICMS.xlsx')
        insc = info['INSCRIÇÃO'][0]
        vap1 = (info['VALOR AP'][0])
        tipo = info['TIPO ICMS'][0]
        tp = 'ANTECIPAÇÃO'
        vap1 = f'{vap1:_.2f}'.replace('.', ',').replace('_', '.')
        vsp = (info['VALOR ST'][0])
        vsp = f'{vsp:_.2f}'.replace('.', ',').replace('_', '.')
        qtn = tabela["NF"]
        qtn = len(qtn)

        if tipo == tp:
            vap = vap1
            self.options = webdriver.ChromeOptions()
            self.options.add_argument("--start-maximized")
            self.nav = webdriver.Chrome(options=self.options)
            self.nav.implicitly_wait(30)
            self.nav.get('https://servicos.sefaz.ba.gov.br/sistemas/arasp/pagamento/modulos/dae/pagamento/dae_pagamento.aspx')
            self.nav.find_element(By.XPATH, '//*[@id="PHConteudo_ddl_antecipacao_tributaria"]').click()
            pyautogui.press('down')
            pyautogui.press('enter')
        else:
            vap = vsp
            self.options = webdriver.ChromeOptions()
            self.options.add_argument("--start-maximized")
            self.nav = webdriver.Chrome(options=self.options)
            self.nav.implicitly_wait(30)
            self.nav.get('https://servicos.sefaz.ba.gov.br/sistemas/arasp/pagamento/modulos/dae/pagamento/dae_pagamento.aspx')
            self.nav.find_element(By.XPATH, '//*[@id="PHConteudo_ddl_antecipacao_tributaria"]').click()
            time.sleep(1)
            pyautogui.press(['down', 'down', 'down'])
            pyautogui.press('enter')
        time.sleep(1)
        self.nav.find_element(By.XPATH, '//*[@id="PHconteudoSemAjax_txt_num_inscricao_estad"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="PHconteudoSemAjax_txt_num_inscricao_estad"]').send_keys(int(insc))
        if self.nav.find_element(By.XPATH, '//*[@id="PHconteudoSemAjax_txt_num_inscricao_estad"]').text == "":
            self.nav.find_element(By.XPATH, '//*[@id="PHconteudoSemAjax_txt_num_inscricao_estad"]').send_keys(int(insc))
        time.sleep(3)
        self.nav.find_element(By.XPATH, '//*[@id="PHconteudoSemAjax_txt_dtc_vencimento"]').click()
        time.sleep(1)

        pyautogui.write(venc)
        self.nav.find_element(By.XPATH, '//*[@id="PHconteudoSemAjax_txt_dtc_max_pagamento"]').click()
        time.sleep(1)
        pyautogui.write(venc)
        self.nav.find_element(By.XPATH, '//*[@id="PHconteudoSemAjax_txt_val_principal"]').send_keys(vap)
        self.nav.find_element(By.XPATH, '//*[@id="PHconteudoSemAjax_txt_mes_ano_referencia_6anos"]').click()
        time.sleep(1)
        pyautogui.write(ref)
        if qtn > 15:
           self.nav.find_element(By.XPATH, '//*[@id="PHconteudoSemAjax_txt_qtd_nota_fiscal"]').send_keys("15")
        else:
            self.nav.find_element(By.XPATH, '//*[@id="PHconteudoSemAjax_txt_qtd_nota_fiscal"]').send_keys(qtn)
        self.nav.find_element(By.XPATH, '//*[@id="PHconteudoSemAjax_txt_num_nota_fiscal"]').click()
        for i, TIPO in enumerate(tabela["TIPO"]):
            n_nf = tabela.loc[i, "NF"]
            ID = str(tabela.loc[i, "ID"])
            n_nf = int(n_nf)
            #time.sleep(1)
            self.nav.find_element(By.XPATH, ID).send_keys(n_nf)
            #time.sleep(1)
            #pyautogui.write('  ')
            pyautogui.press('tab')
    def Iniciar_Chat_GPT(self):
        webbrowser.open(url='https://chat.openai.com/?model=text-davinci-002-render-sha')
    def Iniciar_Chat_Bard(self):
        webbrowser.open(url='https://bard.google.com/')
    def Iniciar_Chat_LuzIa(self):
        self.num = '11972553036'
        self.cod_pais = '+55'
        if self.num == "":
            tk.messagebox.showinfo(title='Informação', message='O campo do celular esta vazio')
        else:
            kt.sendwhatmsg_instantly(self.cod_pais + self.num, "", wait_time=30)

class notas_de_fora():
    def donload_planilha(self):
        self.tb_sn = pd.read_excel('CERTIFICADOS.xlsx')
        self.emp = self.combobox_selecionar_tipo.get()
        self.var = self.tb_sn.loc[(self.tb_sn['EMPRESAS'] == self.emp)]
        self.var.to_excel('PROCV.xlsx')
        self.df = pd.read_excel('PROCV.xlsx')
        self.cert = self.df['BUSCAR'][0]
        self.caminho = caminho
        self.caminho = self.caminho.replace('\\', '/')
        self.save_folder = filedialog.askdirectory(initialdir=self.caminho, title='Por favor selecione a pasta onde esta o programa')

        self.save_folder = self.save_folder.replace('/', '\\')
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option('prefs', {'download.prompt_for_download': False,
                                                       'download.default_directory': self.save_folder})

        self.options.add_argument("--start-maximized")
        # options.add_argument('--headless')
        self.nav = webdriver.Chrome(options=self.options)

        self.nav.implicitly_wait(5)
        self.nav.get(
            'https://www.fsist.com.br/usuario/login.aspx?local=monitor-de-notas&loginhash=86Wgsa%2bhrnltb290bmI%3d')
        self.nav.find_element(By.XPATH, '//*[@id="usuario"]').send_keys('helissonfla@hotmail.com')
        self.nav.find_element(By.XPATH, '//*[@id="senha"]').send_keys('40151311')
        self.nav.find_element(By.XPATH, '//*[@id="butEntrar"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="EmpresaNome"]').click()
        self.nav.find_element(By.XPATH, self.cert).click()
        self.nav.find_element(By.XPATH, '//*[@id="ComboData"]/div[1]').click()
        self.nav.find_element(By.XPATH, '//*[@id="DataMesPassado"]').click()
        try:
            self.nav.find_element(By.XPATH, '//*[@id="linhas"]/tbody/tr[1]/td[1]').click()
        except:
            self.nav.close()
            tk.messagebox.showinfo(title='Informação', message='Não houve notas fiscais no mês passdo !!')
            return
        self.nav.find_element(By.XPATH, '//*[@id="opcoes"]/div[3]/button').click()
        self.nav.find_element(By.XPATH, '//*[@id="msgsim"]').click()
        time.sleep(5)
        self.nav.close()

        # tkinter.messagebox.showinfo(title='Informação', message='Arquivo baixado com sucesso!!'+ emp)
        def rename_file(file):
            file_name, file_extension = os.path.splitext(file)
            file_name_numbers = re.findall(r'\d+', file_name)
            print(file_name)
            print(file_name_numbers)
            if not file_name_numbers:
                return file

            file_name_numbers = 'NF ENTRADA.xlsx'

            return f'{file_name_numbers}'

        def file_loop(root, dirs, files):
            for file in files:
                if not re.search(r'\.xlsx$', file):
                    continue

                new_file_name = rename_file(file)
                old_file_full_path = os.path.join(root, file)
                new_file_full_path = os.path.join(root, new_file_name)
                print(old_file_full_path)
                print(new_file_full_path)
                print(new_file_full_path)
                print(f'Movendo arquivo "{file}" para "{new_file_name}"')
                shutil.move(old_file_full_path, new_file_full_path)

        def main_loop():
            for root, dirs, files in os.walk(self.save_folder):
                file_loop(root, dirs, files)

        if __name__ == '__main__':
            main_loop()
        tk.messagebox.showinfo(title='Informação', message='Arquivo baixado com sucesso!! ' + self.emp)
        self.tabela1 = pd.read_excel('NF ENTRADA.xlsx')
        self.tabela2 = self.tabela1.loc[(self.tabela1['UF'] != 'BA')]
        self.tabela2.to_excel('NF DE FORA.xlsx')
    def download_planilha_auternativo(self):

        self.tb_sn = pd.read_excel('CERTIFICADOS.xlsx')
        self.emp = self.combobox_selecionar_tipo.get()
        self.var = self.tb_sn.loc[(self.tb_sn['EMPRESAS'] == self.emp)]
        self.var.to_excel('PROCV.xlsx')
        self.df = pd.read_excel('PROCV.xlsx')
        self.cert = self.df['BUSCAR'][0]
        self.login = int(self.df['LOGIN '][0])
        self.senha = self.df['SENHA'][0]
        self.caminho = caminho
        self.caminho = self.caminho.replace('\\', '/')
        self.save_folder = filedialog.askdirectory(initialdir=self.caminho,
                                                   title='Por favor selecione a pasta onde esta o programa')
        if self.save_folder == "":
            return
        self.save_folder = self.save_folder.replace('/', '\\')
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option('prefs', {'download.prompt_for_download': False,
                                                       'download.default_directory': self.save_folder})


        self.options.add_argument("--start-maximized")
        # options.add_argument('--headless')

        self.nav = webdriver.Chrome(options=self.options)
        self.nav.implicitly_wait(10)
        self.nav.get(
            'https://nfe.sefaz.ba.gov.br/servicos/NFENC/SSL/ASLibrary/Login?ReturnUrl=%2fservicos%2fnfenc%2fModulos%2fAutenticado%2fRestrito%2fNFENC_consulta_destinatario.aspx')
        self.nav.find_element(By.XPATH, '//*[@id="details-button"]').click()
        self.nav.find_element(By.XPATH, '//*[@id="proceed-link"]').click()
        time.sleep(1)
        self.nav.find_element(By.XPATH, '//*[@id="PHCentro_userLogin"]').send_keys(self.login)
        time.sleep(1)
        self.nav.find_element(By.XPATH, '//*[@id="PHCentro_userPass"]').send_keys(self.senha)
        pyautogui.press('enter')
        time.sleep(3)
        pyautogui.press(['tab', 'tab', 'tab', 'enter'])
        self.nav.find_element(By.XPATH, '//*[@id="rbt_filtro3"]').click()
        time.sleep(1)
        pyautogui.press('tab')
        #self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoInicial"]').click()
        time.sleep(2)
        self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoInicial"]').send_keys(pdmpf)
        time.sleep(1)
        if self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoInicial"]').text == "":
            self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoInicial"]').send_keys(pdmpf)
        self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoFinal"]').click()
        time.sleep(1)
        self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoFinal"]').send_keys(udmpf)
        if self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoFinal"]').text == "":
            self.nav.find_element(By.XPATH, '//*[@id="txtPeriodoFinal"]').send_keys(udmpf)

        self.nav.find_element(By.XPATH, '//*[@id="AplicarFiltro"]').click()
        try:
            self.nav.find_element(By.XPATH, '//*[@id="btn_GerarPlanilha"]').click()
            time.sleep(6)
        except:

            tk.messagebox.showinfo(title='Informação',
                                   message='Não ha notas fiscais emitidas para o periodo informado')

            return

        try:
            df_new = pd.read_csv('rpt_cons_dest.csv', sep=";")
        except:
            self.nav.close()
            tk.messagebox.showinfo(title='Informação', message='As planilhas da ALLCLEAN E L&C devem ser baixadas pelo Fsist!!')
            os.remove("rpt_cons_dest.csv")
            return

        #df_new = pd.read_csv('rpt_cons_dest.csv', sep=";")
        GFG = pd.ExcelWriter('NF ENTRADA.xlsx')
        df_new.to_excel(GFG, index=False)
        GFG.close()
        tabela1 = pd.read_excel('NF ENTRADA.xlsx')
        tabela2 = tabela1.loc[(tabela1['UF Emit.'] != ' BA')]
        tabela2.to_excel('NF DE FORA.xlsx')

        os.remove("rpt_cons_dest.csv")
        self.nav.close()
        tk.messagebox.showinfo(title='Informação', message='Arquivo baixado com sucesso!! ' + self.emp)
    def donload_pdf(self, init_folder=None):
        self.tb_sn = pd.read_excel('CERTIFICADOS.xlsx')
        self.emp = self.combobox_selecionar_tipo.get()
        self.var = self.tb_sn.loc[(self.tb_sn['EMPRESAS'] == self.emp)]
        self.var.to_excel('PROCV.xlsx')
        self.df = pd.read_excel('PROCV.xlsx')
        self.cert = self.df['CERTIFICADO'][0]
        self.save_folder = filedialog.askdirectory(title="Por favor selecione a pasta para salvar")
        self.save_folder = self.save_folder.replace('/', '\\')
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option('prefs', {'download.prompt_for_download': False,
                                                       'download.default_directory': self.save_folder})
        self.options.add_argument("--start-maximized")
        # options.add_argument('--headless')
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.implicitly_wait(30)
        self.nav.get('https://www.fsist.com.br/xmls-compartilhados')
        self.nav.find_element(By.XPATH,
                              '//*[@id="ContentPlaceHolder1_PanPublico"]/table/tbody/tr/td/div/table/tbody/tr/td['
                              '2]/div[2]/div/table/tbody/tr/td[2]').click()
        time.sleep(0.5)
        self.nav.find_element(By.XPATH, '//*[@id="FormAcessoInstalar"]/div/div/div/table/tbody/tr/td[2]/button').click()
        # time.sleep(0.5)
        self.nav.find_element(By.XPATH, '//*[@id="FormAcesso"]/div[1]/div/div[1]/table/tbody/tr/td[2]').click()
        time.sleep(1)
        pyautogui.press('left')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(3)
        pyautogui.press('enter')
        time.sleep(2)

        try:
            self.nav.find_element(By.XPATH, self.cert).click()

        except:
            self.nav.minimize_window()
            time.sleep(1)
            pyautogui.press('enter')
            self.nav.maximize_window()
            time.sleep(1)
            self.nav.find_element(By.XPATH, self.cert).click()

        self.tabela1 = pd.read_excel('NF ENTRADA.xlsx')
        self.tabela2 = self.tabela1.loc[(self.tabela1['UF'] != 'BA')]
        self.tabela2.to_excel('NF DE FORA.xlsx')# procv
        self.tabela = pd.read_excel('NF DE FORA.xlsx')
        for i, Numero in enumerate(self.tabela["Numero"]):
            self.Chave = self.tabela.loc[i, "Chave"]
            time.sleep(1)
            try:
                self.nav.find_element(By.XPATH, '//*[@id="busca"]').send_keys(self.Chave)
            except:
                time.sleep(1)
                self.nav.find_element(By.XPATH, '//*[@id="butCadastroConfirmar"]').click()
                time.sleep(1)
                self.nav.find_element(By.XPATH, '//*[@id="busca"]').send_keys(self.Chave)

            time.sleep(1)
            self.nav.find_element(By.XPATH, '//*[@id="ButBuscaTipo1"]').click()
            time.sleep(1)
            self.nav.find_element(By.XPATH, '//*[@id="DivResultado"]/div[2]/table/tbody/tr/td[1]/span').click()
            time.sleep(1)
            self.nav.find_element(By.XPATH, '//*[@id="ButPDF"]').click()
            time.sleep(1)
            self.nav.find_element(By.XPATH, '//*[@id="ButDownlaodSair"]').click()
            time.sleep(1)
            self.nav.find_element(By.XPATH, '//*[@id="busca"]').clear()
        self.nav.close()
        # self.tabela.to_excel('NF DE FORA.xlsx')
        self.total_nf = str(len(self.tabela))
        print(self.total_nf)
        tk.messagebox.showinfo(title='Informação', message='Processo concluido total de nota fiscal baixada  ' + self.total_nf)
    def download_pdf_auternativo(self):
        tabela = pd.read_excel('NF DE FORA.xlsx')
        tabcont = tabela.columns.size
        chave_n_loc =[]
        if tabcont == 14:
            doc = "Numero"
            chave = "Chave"
        else:
            doc = "Numero NF-e"
            chave = "Chave de Acesso"

        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.nav = webdriver.Chrome(options=self.options)
        self.nav.implicitly_wait(60*5)
        for i, numero in enumerate(tabela[doc]):
           try:
                n_nf = tabela.loc[i, chave]
                n_nf = re.sub('[^0-9]', '', n_nf)
                urlSite = 'https://consultadanfe.com/'
                print(n_nf)
                self.nav.get(urlSite)
                self.nav.find_element(By.XPATH, '//*[@id="chave"]').send_keys(n_nf)
                self.nav.find_element(By.XPATH, '//*[@id="form_one"]/button').click()
                #self.nav.find_element(By.XPATH, '//*[@id="downloadDanfePdf"]').click()
                tk.messagebox.showinfo(title='Informação', message='Para ir para a proxima nota fiscal clickar em OK ')
           except:
               chave_n_loc.append(n_nf)
               tk.messagebox.showinfo(title='Informação', message=f'erro a chave {n_nf}              não foi encontrada.'
                                                                  f' Para ir para a proxima nota fiscal clickar em OK')


            # with sync_playwright() as p:
            #     browser = p.chromium.launch(headless=False)
            #     page = browser.new_page()
            #     page.goto("https://meudanfe.com.br/#")
            #     page.locator('xpath=//*[@id="chaveAcessoBusca"]').type(n_nf)
            #     page.locator('xpath=//*[@id="searchDiv"]/div/div/div/div/form/div[2]/div[2]/div[2]/a').click()
            #     page.wait_for_selector('input#natOperacao', timeout=600000)
            #     with page.expect_popup() as popup_info:
            #         page.locator('xpath=//*[@id="downloadDanfePdf"]').click()
            #     popup = popup_info.value
            #     popup.wait_for_load_state()
            #     print(popup.title())
            #     page.wait_for_timeout(5000)
            #     page.goto("https://www.arealme.com/click-speed-test/pt/")
            #     page.wait_for_selector('button#clickarena', timeout=600000)
            #     browser.close()


        df_chave = pd.DataFrame(chave_n_loc)
        df_chave.to_csv('Chaves com erro na emição do PDF.txt', sep='\t', index=False)
        tk.messagebox.showinfo(title='Informação', message='Processo finalidao')
    def Load_excel_data(self):
        """If the file selected is valid this will load the file into the Treeview"""
        self.file_path = self.label_file["text"]
        try:
            self.excel_filename = r"\\Desktop-1v4f59g\meneses\MENESES CONT 2\NF ENTRADA.xlsx".format(self.file_path)
            if self.excel_filename[-4:] == ".xlsx":
                self.df2 = pd.read_csv(self.excel_filename)
            else:
                self.df2 = pd.read_excel(self.excel_filename)

        except ValueError:
            self.msg = tk.messagebox.showerror("Information", "The file you have chosen is invalid")
            return None
        except FileNotFoundError:
            self.msg1 = tk.messagebox.showerror("Information", f"Caminho não encontrado")
            return None

        self.clear_data()
        self.tv1["column"] = list(self.df2.columns)
        self.tv1["show"] = "headings"
        for column in self.tv1["columns"]:
            self.tv1.heading(column, text=column)  # let the column heading = column name

        self.df2_rows = self.df2.to_numpy().tolist()  # turns the dataframe into a list of lists
        for row in self.df2_rows:
            self.tv1.insert("", "end",
                            values=row)  # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
        return None
    def Load_excel_fora_do_estado(self):
        """If the file selected is valid this will load the file into the Treeview"""
        self.file_path = self.label_file["text"]
        try:
            self.excel_filename = r"\\Desktop-1v4f59g\meneses\MENESES CONT 2\NF DE FORA.xlsx".format(self.file_path)
            if self.excel_filename[-4:] == ".xlsx":
                self.df3 = pd.read_csv(self.excel_filename)
            else:
                self.df3 = pd.read_excel(self.excel_filename)

        except ValueError:
            self.tk.messagebox.showerror("Information", "The file you have chosen is invalid")
            return None
        except FileNotFoundError:
            self.tk.messagebox.showerror("Information", f"No such file as {self.file_path}")
            return None

        self.clear_data()
        self.tv1["column"] = list(self.df3.columns)
        self.tv1["show"] = "headings"
        for column in self.tv1["columns"]:
            self.tv1.heading(column, text=column)  # let the column heading = column name

        self.df3_rows = self.df3.to_numpy().tolist()  # turns the dataframe into a list of lists
        for row in self.df3_rows:
            self.tv1.insert("", "end",
                            values=row)  # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
        return None
    def clear_data(self):
        self.tv1.delete(*self.tv1.get_children())
        return None
    def open_plan_icms(self):
        self.path = ('ICMS.xlsm')
        os.startfile(self.path)

class application(Funcs, Funcs2, Funcs3, web_sn, Funcs4,Funcs5, notas_de_fora, Validadores):
    def __init__(self):
        self.janela = janela
        self.ValidarEntrada()
        self.tela_principa()
        self.Menus()
        janela.mainloop()

    def tela_principa(self):
        self.janela.title('MENESES CONTABILIDADE')
        self.janela.configure(background='#1e3743')
        self.janela.geometry('800x600')
        self.janela.resizable(True, True)
        self.img = PhotoImage(file='LOGO EM PNG 800x600.png')
        self.label_imagem = Label(janela, image=self.img, background='#1e3743').pack()
    def Menus(self):
        menubar = Menu(self.janela)
        self.janela.config(menu=menubar)
        filemenu = Menu(menubar, tearoff=0)
        filemenu2 = Menu(menubar, tearoff=0)
        filemenu3 = Menu(menubar, tearoff=0)
        filemenu4 = Menu(menubar, tearoff=0)
        filemenu5 = Menu(menubar, tearoff=0)
        filemenu6 = Menu(menubar, tearoff=0)
        filemenu7 = Menu(menubar, tearoff=0)


        #def Quit(): self.janela.destroy()

        menubar.add_cascade(label="COD.ACESSO", menu=filemenu)
        menubar.add_cascade(label="AGENDA", menu=filemenu2)
        menubar.add_cascade(label="CONSULTAS", menu=filemenu3)
        menubar.add_cascade(label="IMPOSTOS", menu=filemenu4)
        menubar.add_cascade(label="DEPARTAMENTO PESSOAL", menu=filemenu5)
        menubar.add_cascade(label="FISCAL", menu=filemenu6)
        menubar.add_cascade(label="AJUDA", menu=filemenu7)
        #menubar.add_cascade(label="SAIR")

        filemenu.add_command(label="SIMPLES NACIONAL", command=self.janela3)
        filemenu.add_command(label="SEFAZ", command=self.janela4)
        filemenu.add_command(label="GESTÃO ISS", command=self.Gestao_iss)
        #filemenu.add_command(label="SAIR", command=quit)
        filemenu2.add_command(label="CONTATOS", command=self.janela2)
        filemenu3.add_command(label="CNPJ", command=self.janela7)
        filemenu3.add_command(label="PREFEITURA",command=self.janela8)
        filemenu3.add_command(label="INSCRIÇÂO ESTADUAL", command=self.janela9)
        filemenu3.add_command(label="SOLICITAÇÂO DE ARQUIVOS", command=self.janela10)
        #filemenu4.add_command(label="EMPREGADOR DIGITAL S/A", command=self.empregador_digital_sem_aviso)
        filemenu5.add_command(label="EMPREGADOR WEB", command=self.empregador_digital_aviso)
        filemenu4.add_command(label="NOTAS FISCAIS PARA O ICMS", command=self.janela6)
        filemenu5.add_command(label="RECISÃO", command=self.controle_rescisao)
        filemenu6.add_command(label="CONTROLE ICMS", command=self.controle_fiscal)
        filemenu7.add_command(label="ASSISTENTE VIRTUAL", command=self.janela11)
        #filemenu5.add_command(label="SEGURO DESEMPREGO", command=self.controle_SD)

    def janela2(self):
        self.root = Toplevel()
        self.tela()
        self.frames_da_tela()
        self.widgets_frame1()
        self.lista_frame2()
        self.montaTabelas()
        self.select_lista()
        self.root.transient(self.janela)
        self.root.resizable(True, True)
        self.root.focus_force()
        self.root.grab_set()
    def ValidarEntrada(self):
        self.vcmd = (self.janela.register(self.validadores_entry), "%P")
    def tela(self):
        self.root.title("Cadastro de Clientes")
        self.root.configure(background='#1e3743')
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        self.root.maxsize(width=900, height=700)
        self.root.minsize(width=500, height=400)
    def frames_da_tela(self):
        self.frame_1 = Frame(self.root, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.46)

        self.frame_2 = Frame(self.root, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)
    def widgets_frame1(self):
        ### Criação do botao limpar
        self.bt_limpar = Button(self.frame_1, text="Limpar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.limpa_cliente)
        self.bt_limpar.place(relx=0.2, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao buscar
        self.bt_limpar = Button(self.frame_1, text="Buscar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.busca_cliente)
        self.bt_limpar.place(relx=0.3, rely=0.1, relwidth=0.1, relheight=0.15)

        ### Criação do botao Whatsapp
        self.bt_limpar = Button(self.frame_1, text="Whatsapp", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.sendwhats)
        self.bt_limpar.place(relx=0.4, rely=0.1, relwidth=0.1, relheight=0.15)

        ### Criação do botao gmail
        self.bt_limpar = Button(self.frame_1, text="Email", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.gmail)
        self.bt_limpar.place(relx=0.5, rely=0.1, relwidth=0.1, relheight=0.15)

        ### Criação do botao novo
        self.bt_limpar = Button(self.frame_1, text="Novo", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.add_cliente)
        self.bt_limpar.place(relx=0.6, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao alterar
        self.bt_limpar = Button(self.frame_1, text="Alterar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.alterar_cliente)
        self.bt_limpar.place(relx=0.7, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao apagar
        self.bt_limpar = Button(self.frame_1, text="Apagar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.deleta_cliente)
        self.bt_limpar.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.15)

        ## Criação da label e entrada do codigo
        self.lb_codigo = Label(self.frame_1, text="Código", bg='#dfe3ee', fg='#107db2')
        self.lb_codigo.place(relx=0.05, rely=0.05)

        self.codigo_entry = Entry(self.frame_1, validate= "key", validatecommand= self.vcmd )
        self.codigo_entry.place(relx=0.05, rely=0.15, relwidth=0.08)

        ## Criação da label e entrada do nome
        self.lb_nome = Label(self.frame_1, text="Nome", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.35)

        self.nome_entry = Entry(self.frame_1)
        self.nome_entry.place(relx=0.05, rely=0.45, relwidth=0.8)

        ## Criação da label e entrada do celular
        self.lb_nome = Label(self.frame_1, text="Celular", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.6)

        self.cel_entry = Entry(self.frame_1)
        self.cel_entry.place(relx=0.05, rely=0.7, relwidth=0.2)

        ## Criação da label e entrada telefone
        self.lb_nome = Label(self.frame_1, text="Telefone", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.3, rely=0.6)

        self.fone_entry = Entry(self.frame_1)
        self.fone_entry.place(relx=0.3, rely=0.7, relwidth=0.2)

        ## Criação da label e entrada email
        self.lb_nome = Label(self.frame_1, text="Email", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.55, rely=0.6)

        self.cidade_entry = Entry(self.frame_1)
        self.cidade_entry.place(relx=0.55, rely=0.7, relwidth=0.3)
    def lista_frame2(self):
        self.listaCli = ttk.Treeview(self.frame_2, height=3,
                                     column=("col1", "col2", "col3", "col4", "col5"))
        self.listaCli.heading("#0", text="")
        self.listaCli.heading("#1", text="Codigo")
        self.listaCli.heading("#2", text="Nome")
        self.listaCli.heading("#3", text="Celular")
        self.listaCli.heading("#4", text="Telefone")
        self.listaCli.heading("#5", text="E-mail")
        self.listaCli.column("#0", width=1)
        self.listaCli.column("#1", width=50)
        self.listaCli.column("#2", width=200)
        self.listaCli.column("#3", width=125)
        self.listaCli.column("#4", width=125)
        self.listaCli.column("#4", width=125)
        self.listaCli.place(relx=0.01, rely=0.1, relwidth=0.95, relheight=0.85)

        self.scroolLista = Scrollbar(self.frame_2, orient='vertical')
        self.listaCli.configure(yscroll=self.scroolLista.set)
        self.scroolLista.place(relx=0.96, rely=0.1, relwidth=0.04, relheight=0.85)
        self.listaCli.bind("<Double-1>", self.OnDoubleClick)

    def janela3(self):
        self.root3 = Toplevel()
        self.listaEMP = None
        self.tela_sn()
        self.frames_da_tela_sn()
        self.widgets_frame1_sn()
        self.lista_frame2_sn()
        self.montaTabelas1()
        self.select_lista1()
        self.root3.transient(self.janela)
        self.root3.resizable(True, True)
        self.root3.focus_force()
        self.root3.grab_set()
    def tela_sn(self):
        self.root3.title("Cadastro e Login Simples Nacional")
        self.root3.configure(background='#1e3743')
        self.root3.geometry("700x500")
        self.root3.resizable(True, True)
        self.root3.maxsize(width=900, height=700)
        self.root3.minsize(width=500, height=400)
    def frames_da_tela_sn(self):
        self.frame_1 = Frame(self.root3, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.46)

        self.frame_2 = Frame(self.root3, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)
    def widgets_frame1_sn(self):
        ### Criação do botao consultar cnpj
        #self.bt_limpar = Button(self.frame_1, text="CNPJ", bd=2, bg='#107db2', fg='white'
                                #, font=('verdana', 8, 'bold'), command=self.consultar_cnpj)
        #self.bt_limpar.place(relx=0.2, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao buscar
        self.bt_buscar = Button(self.frame_1, text="Buscar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.busca_empresa)
        self.bt_buscar.place(relx=0.2, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao login simples nacional pgdas
        self.bt_buscar = Button(self.frame_1, text="PGDAS", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.login_simples)
        self.bt_buscar.place(relx=0.3, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao login simples nacional parcelamento
        self.bt_buscar = Button(self.frame_1, text="Parcelamento", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 6, 'bold'), command=self.login_simples_parcelamento)
        self.bt_buscar.place(relx=0.5, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao novo
        self.bt_limpar = Button(self.frame_1, text="Novo", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.add_empresa)
        self.bt_limpar.place(relx=0.6, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao alterar
        self.bt_limpar = Button(self.frame_1, text="Alterar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.altera_empresa)
        self.bt_limpar.place(relx=0.7, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao apagar
        self.bt_limpar = Button(self.frame_1, text="Apagar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.deleta_empresa)
        self.bt_limpar.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.15)

        ## Criação da label e entrada do codigo
        self.lb_codigo = Label(self.frame_1, text="Código", bg='#dfe3ee', fg='#107db2')
        self.lb_codigo.place(relx=0.05, rely=0.05)

        self.codigo1_entry = Entry(self.frame_1)
        self.codigo1_entry.place(relx=0.05, rely=0.15, relwidth=0.08)

        ## Criação da label e entrada do nome
        self.lb_nome = Label(self.frame_1, text="Empresa", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.35)

        self.empresa_entry = Entry(self.frame_1)
        self.empresa_entry.place(relx=0.05, rely=0.45, relwidth=0.8)

        ## Criação da label e entrada do telefone
        self.lb_nome = Label(self.frame_1, text="CNPJ", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.6)

        self.cnpj_entry = Entry(self.frame_1)
        self.cnpj_entry.place(relx=0.05, rely=0.7, relwidth=0.2)


        ## Criação da label e entrada da cidade
        self.lb_nome = Label(self.frame_1, text="CPF", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.35, rely=0.6)

        self.cpf_entry = Entry(self.frame_1)
        self.cpf_entry.place(relx=0.35, rely=0.7, relwidth=0.2)

        self.lb_nome = Label(self.frame_1, text="Senha", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.65, rely=0.6)

        self.senha_entry = Entry(self.frame_1)
        self.senha_entry.place(relx=0.65, rely=0.7, relwidth=0.2)
    def lista_frame2_sn(self):
        self.listaEMP = ttk.Treeview(self.frame_2, height=3,
                                     column=("col1", "col2", "col3", "col4", "col5"))
        self.listaEMP.heading("#0", text="")
        self.listaEMP.heading("#1", text="Codigo")
        self.listaEMP.heading("#2", text="Nome")
        self.listaEMP.heading("#3", text="CNPJ")
        self.listaEMP.heading("#4", text="CPF")
        self.listaEMP.heading("#5", text="Senha")
        self.listaEMP.column("#0", width=1)
        self.listaEMP.column("#1", width=50)
        self.listaEMP.column("#2", width=200)
        self.listaEMP.column("#3", width=100)
        self.listaEMP.column("#4", width=100)
        self.listaEMP.column("#5", width=100)
        self.listaEMP.place(relx=0.01, rely=0.1, relwidth=0.95, relheight=0.85)

        self.scroolLista = Scrollbar(self.frame_2, orient='vertical')
        self.listaEMP.configure(yscroll=self.scroolLista.set)
        self.scroolLista.place(relx=0.96, rely=0.1, relwidth=0.04, relheight=0.85)
        self.listaEMP.bind("<Double-1>", self.OnDoubleClick1)

    def janela4(self):
        self.root4 = Toplevel()
        self.listaEMP = None
        self.tela_sf()
        self.frames_da_tela_sf()
        self.widgets_frame1_sf()
        self.lista_frame2_sf()
        self.montaTabelas2()
        self.select_lista2()
        self.root4.transient(self.janela)
        self.root4.resizable(True, True)
        self.root4.focus_force()
        self.root4.grab_set()
    def tela_sf(self):
        self.root4.title("Cadastro e Login Sefaz")
        self.root4.configure(background='#1e3743')
        self.root4.geometry("700x500")
        self.root4.resizable(True, True)
        self.root4.maxsize(width=900, height=700)
        self.root4.minsize(width=500, height=400)
    def frames_da_tela_sf(self):
        self.frame_1 = Frame(self.root4, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.46)

        self.frame_2 = Frame(self.root4, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)
    def widgets_frame1_sf(self):
        ### Criação do botao limpar
        self.bt_limpar = Button(self.frame_1, text="Limpar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.limpa_empresa_sf)
        self.bt_limpar.place(relx=0.2, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao buscar
        self.bt_buscar = Button(self.frame_1, text="Buscar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.busca_empresa_sf)
        self.bt_buscar.place(relx=0.3, rely=0.1, relwidth=0.1, relheight=0.15)
        ### criação do botão login sefaz notas de entrada
        self.bt_buscar = Button(self.frame_1, text="Entrada", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.login_sefaz_entrada)
        self.bt_buscar.place(relx=0.4, rely=0.1, relwidth=0.1, relheight=0.15)
        ### criação do botão login sefaz notas de saida
        self.bt_buscar = Button(self.frame_1, text="saida", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.login_sefaz_saida)
        self.bt_buscar.place(relx=0.5, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao novo
        self.bt_novo = Button(self.frame_1, text="Novo", bd=2, bg='#107db2', fg='white'
                              , font=('verdana', 8, 'bold'), command=self.add_empresa_sf)
        self.bt_novo.place(relx=0.6, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao alterar
        self.bt_alterar = Button(self.frame_1, text="Alterar", bd=2, bg='#107db2', fg='white'
                                 , font=('verdana', 8, 'bold'), command=self.altera_empresa_sf)
        self.bt_alterar.place(relx=0.7, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao apagar
        self.bt_apagar = Button(self.frame_1, text="Apagar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.deleta_empresa_sf)
        self.bt_apagar.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.15)

        ## Criação da label e entrada do codigo
        self.lb_codigo = Label(self.frame_1, text="Código", bg='#dfe3ee', fg='#107db2')
        self.lb_codigo.place(relx=0.05, rely=0.05)

        self.codigo2_entry = Entry(self.frame_1)
        self.codigo2_entry.place(relx=0.05, rely=0.15, relwidth=0.08)

        ## Criação da label e entrada do nome
        self.lb_nome = Label(self.frame_1, text="Empresa", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.35)

        self.empresa2_entry = Entry(self.frame_1)
        self.empresa2_entry.place(relx=0.05, rely=0.45, relwidth=0.8)

        ## Criação da label e entrada do usuario
        self.lb_nome = Label(self.frame_1, text="Usuario", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.6)

        self.usuario_entry = Entry(self.frame_1)
        self.usuario_entry.place(relx=0.05, rely=0.7, relwidth=0.2)

        ## Criação da label e entrada da senha
        self.lb_nome = Label(self.frame_1, text="Senha", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.65, rely=0.6)

        self.senha2_entry = Entry(self.frame_1)
        self.senha2_entry.place(relx=0.65, rely=0.7, relwidth=0.2)
    def lista_frame2_sf(self):
        self.listaEMP = ttk.Treeview(self.frame_2, height=3,
                                     column=("col1", "col2", "col3", "col4"))
        self.listaEMP.heading("#0", text="")
        self.listaEMP.heading("#1", text="Codigo")
        self.listaEMP.heading("#2", text="Nome")
        self.listaEMP.heading("#3", text="Usuario")
        self.listaEMP.heading("#4", text="Senha")
        self.listaEMP.column("#0", width=1)
        self.listaEMP.column("#1", width=50)
        self.listaEMP.column("#2", width=200)
        self.listaEMP.column("#3", width=100)
        self.listaEMP.column("#4", width=100)
        self.listaEMP.place(relx=0.01, rely=0.1, relwidth=0.95, relheight=0.85)

        self.scroolLista = Scrollbar(self.frame_2, orient='vertical')
        self.listaEMP.configure(yscroll=self.scroolLista.set)
        self.scroolLista.place(relx=0.96, rely=0.1, relwidth=0.04, relheight=0.85)
        self.listaEMP.bind("<Double-1>", self.OnDoubleClick2)

    def janela5(self):
        self.root5 = Toplevel()
        self.listaEMP = None
        self.tela_ISS()
        self.frames_da_tela_ISS()
        self.widgets_frame1_ISS()
        self.lista_frame2_ISS()
        self.montaTabelas3()
        self.select_lista3()
        self.root5.transient(self.janela)
        self.root5.resizable(True, True)
        self.root5.focus_force()
        self.root5.grab_set()
    def tela_ISS(self):
        self.root5.title("Cadastro e Login ISS intel")
        self.root5.configure(background='#1e3743')
        self.root5.geometry("700x500")
        self.root5.resizable(True, True)
        self.root5.maxsize(width=900, height=700)
        self.root5.minsize(width=500, height=400)
    def frames_da_tela_ISS(self):
        self.frame_1 = Frame(self.root5, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.46)

        self.frame_2 = Frame(self.root5, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)
    def widgets_frame1_ISS(self):
        ### Criação do botao limpar
        self.bt_limpar = Button(self.frame_1, text="Limpar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.limpa_empresa_ISS)
        self.bt_limpar.place(relx=0.2, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao buscar
        self.bt_buscar = Button(self.frame_1, text="Buscar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'),command=self.busca_empresa_ISS)
        self.bt_buscar.place(relx=0.3, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do login iss intel
        self.bt_buscar = Button(self.frame_1, text="Gestão ISS", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.login_iss)
        self.bt_buscar.place(relx=0.5, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao novo
        self.bt_novo = Button(self.frame_1, text="Novo", bd=2, bg='#107db2', fg='white'
                              , font=('verdana', 8, 'bold'), command=self.add_empresa_ISS)
        self.bt_novo.place(relx=0.6, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao alterar
        self.bt_alterar = Button(self.frame_1, text="Alterar", bd=2, bg='#107db2', fg='white'
                                 , font=('verdana', 8, 'bold'), command=self.altera_empresa_ISS)
        self.bt_alterar.place(relx=0.7, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao apaga
        self.bt_apagar = Button(self.frame_1, text="Apagar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.deleta_empresa_ISS)
        self.bt_apagar.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.15)

        ## Criação da label e entrada do codigo
        self.lb_codigo = Label(self.frame_1, text="Código", bg='#dfe3ee', fg='#107db2')
        self.lb_codigo.place(relx=0.05, rely=0.05)

        self.codigo3_entry = Entry(self.frame_1)
        self.codigo3_entry.place(relx=0.05, rely=0.15, relwidth=0.08)

        ## Criação da label e entrada do nome
        self.lb_nome = Label(self.frame_1, text="Empresa", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.35)

        self.empresa3_entry = Entry(self.frame_1)
        self.empresa3_entry.place(relx=0.05, rely=0.45, relwidth=0.8)

        ## Criação da label e entrada do usuario
        self.lb_nome = Label(self.frame_1, text="Usuario", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.6)

        self.login_entry = Entry(self.frame_1)
        self.login_entry.place(relx=0.05, rely=0.7, relwidth=0.2)

        ## Criação da label e entrada da senha
        self.lb_nome = Label(self.frame_1, text="Senha", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.65, rely=0.6)

        self.senha3_entry = Entry(self.frame_1)
        self.senha3_entry.place(relx=0.65, rely=0.7, relwidth=0.2)
    def lista_frame2_ISS(self):
        self.listaEMP = ttk.Treeview(self.frame_2, height=3, column=("col1", "col2", "col3", "col4"))
        self.listaEMP.heading("#0", text="")
        self.listaEMP.heading("#1", text="Codigo")
        self.listaEMP.heading("#2", text="Nome")
        self.listaEMP.heading("#3", text="Login")
        self.listaEMP.heading("#4", text="Senha")
        self.listaEMP.column("#0", width=1)
        self.listaEMP.column("#1", width=50)
        self.listaEMP.column("#2", width=200)
        self.listaEMP.column("#3", width=100)
        self.listaEMP.column("#4", width=100)
        self.listaEMP.place(relx=0.01, rely=0.1, relwidth=0.95, relheight=0.85)

        self.scroolLista = Scrollbar(self.frame_2, orient='vertical')
        self.listaEMP.configure(yscroll=self.scroolLista.set)
        self.scroolLista.place(relx=0.96, rely=0.1, relwidth=0.04, relheight=0.85)
        self.listaEMP.bind("<Double-1>", self.OnDoubleClick3)

    def janela6(self):
        self.root6 = Toplevel()
        self.root6.title('Download notas ficais')
        self.root6.geometry("800x600")  # set the root dimensions
        self.root6.pack_propagate(False)  # tells the root to not let the widgets inside it determine its size.
        self.root6.resizable(0, 0)  # makes the root window fixed in size.
        self.root6.transient(self.janela)
        self.root6.resizable(False, False)
        self.root6.configure(background='#1e3743')


        #self.root6.focus_force()
        #self.root6.grab_set()
        # Frame for TreeView
        self.frame1 = atk.Frame3d(self.root6,bg="#1e3743")
        self.frame1.place(height=490, width=800, rely=0.18, relx=0)

        # Frame for open file dialog
        self.file_frame = atk.Frame3d(self.root6,bg="#1e3743")
        self.file_frame.place(height=100, width=800, rely=0.01, relx=0)

        # Buttons
        self.button1 = tk.Button(self.file_frame, text="Planilha Fsist", font=("georgia"), bg="#1e3743", fg="white", bd=0,
        highlightthickness=0,border=4,borderwidth=5, command=self.donload_planilha)
        self.button1.place(rely=0.55, relx=0.335)

        self.button2 = tk.Button(self.file_frame, text="visualizar planilha",font=("georgia"),bg="#1e3743",fg="white",bd=0,
        highlightthickness=0,border=4,borderwidth=5, command=self.Load_excel_data)
        self.button2.place(rely=0.55, relx=0.15)

        self.button3 = tk.Button(self.file_frame, text="Planilha Sefaz",font=("georgia"),bg="#1e3743",fg="white",bd=0,
        highlightthickness=0,border=4,borderwidth=5, command=self.download_planilha_auternativo)
        self.button3.place(rely=0.55, relx=0.48)

        self.button4 = tk.Button(self.file_frame, text="NF baixadas", font=("georgia"), bg="#1e3743", fg="white", bd=0,
        highlightthickness=0,border=4, borderwidth=5, command=self.Load_excel_fora_do_estado)
        self.button4.place(rely=0.55, relx=0.01)

        self.button5 = tk.Button(self.file_frame, text="Baixar PDF", font=("georgia"), bg="#1e3743", fg="white", bd=0,
        highlightthickness=0, border=4,borderwidth=5, command=self.download_pdf_auternativo)
        self.button5.place(rely=0.55, relx=0.63)

        self.button6 = tk.Button(self.file_frame, text="Lançamento", font=("georgia"), bg="#1e3743", fg="white", bd=0,
        highlightthickness=0,border=4,borderwidth=5, command=self.open_plan_icms)
        self.button6.place(rely=0.55, relx=0.758)

        self.button7 = tk.Button(self.file_frame, text="Emissão", font=("georgia"), bg="#1e3743", fg="white", bd=0,
        highlightthickness=0, border=4, borderwidth=5, command=self.Emissao_icms)
        self.button7.place(rely=0.55, relx=0.895)
        self.combobox_selecionar_tipo = ttk.Combobox(self.file_frame, values=lista_tipos)
        self.combobox_selecionar_tipo.place(rely=0.10, relx=0.16, width=250)
        # The file/file path text
        self.label_file = ttk.Label(self.file_frame, text="Selecione a empresa", background="#1e3743", foreground="white")
        self.label_file.place(rely=0.1, relx=0.01)

        ## Treeview Widget
        self.tv1 = ttk.Treeview(self.frame1)
        self.tv1.place(relheight=1,
                       relwidth=1, )  # set the height and width of the widget to 100% of its container (frame1).

        self.treescrolly = tk.Scrollbar(self.frame1, orient="vertical",
                                        command=self.tv1.yview)  # command means update the yaxis view of the widget
        self.treescrollx = tk.Scrollbar(self.frame1, orient="horizontal",
                                        command=self.tv1.xview)  # command means update the xaxis view of the widget
        self.tv1.configure(xscrollcommand=self.treescrollx.set,
                           yscrollcommand=self.treescrolly.set)  # assign the scrollbars to the Treeview Widget
        self.treescrollx.pack(side="bottom", fill="x")  # make the scrollbar fill the x axis of the Treeview widget
        self.treescrolly.pack(side="right", fill="y")  # make the scrollbar fill the y axis of the Treeview widget

    def janela7(self):
        self.root7 = Toplevel()
        self.listaEMP = None
        self.tela_const_cnpj()
        self.frames_da_tela_const_cnpj()
        self.widgets_frame1_const_cnpj()
        self.lista_frame2_const_cnpj()
        self.montaTabelas1()
        self.select_lista1()
        self.root7.transient(self.janela)
        self.root7.resizable(True, True)
        self.root7.focus_force()
        self.root7.grab_set()
    def tela_const_cnpj(self):
        self.root7.title("Consultar CNPJ")
        self.root7.configure(background='#1e3743')
        self.root7.geometry("700x500")
        self.root7.resizable(True, True)
        self.root7.maxsize(width=900, height=700)
        self.root7.minsize(width=500, height=400)
    def frames_da_tela_const_cnpj(self):
        self.frame_1 = Frame(self.root7, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.46)

        self.frame_2 = Frame(self.root7, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)
    def widgets_frame1_const_cnpj(self):
        ### Criação do botao consultar cnpj
        self.bt_cnpj = Button(self.frame_1, text="Consultar CNPJ", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.consultar_cnpj)

        self.bt_cnpj.place(relx=0.05, rely=0.15, relwidth=0.08, relheight=0.15, width=100)
        ### Criação do botao buscar
        self.bt_buscar = Button(self.frame_1, text="Buscar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'),command=self.busca_empresa)
        self.bt_buscar.place(relx=0.3, rely=0.15, relwidth=0.1, relheight=0.15)
        ### Criação do botao login simples nacional pgdas
        self.bt_buscar = Button(self.frame_1, text="PGDAS", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.login_simples)
        #self.bt_buscar.place(relx=0.4, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao login simples nacional parcelamento
        self.bt_buscar = Button(self.frame_1, text="Parcelamento", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 6, 'bold'), command=self.login_simples_parcelamento)
       # self.bt_buscar.place(relx=0.5, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao novo
        self.bt_limpar = Button(self.frame_1, text="Novo", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.add_empresa)
        #self.bt_limpar.place(relx=0.6, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao alterar
        self.bt_limpar = Button(self.frame_1, text="Alterar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.altera_empresa)
        #self.bt_limpar.place(relx=0.7, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao apagar
        self.bt_limpar = Button(self.frame_1, text="Apagar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.deleta_empresa)
        #self.bt_limpar.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.15)

        ## Criação da label e entrada do codigo
        self.lb_codigo = Label(self.frame_1, text="Código", bg='#dfe3ee', fg='#107db2')
        #self.lb_codigo.place(relx=0.05, rely=0.05)

        self.codigo1_entry = Entry(self.frame_1)
        #self.codigo1_entry.place(relx=0.05, rely=0.15, relwidth=0.08)

        ## Criação da label e entrada do nome
        self.lb_nome = Label(self.frame_1, text="Empresa", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.35)

        self.empresa_entry = Entry(self.frame_1)
        self.empresa_entry.place(relx=0.05, rely=0.45, relwidth=0.8)

        ## Criação da label e entrada do cnpj
        self.lb_nome = Label(self.frame_1, text="CNPJ", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.6)

        self.cnpj_entry = Entry(self.frame_1)
        self.cnpj_entry.place(relx=0.05, rely=0.7, relwidth=0.2)

        ## Criação da label e entrada da cidade
        self.lb_nome = Label(self.frame_1, text="CPF", bg='#dfe3ee', fg='#107db2')
        #self.lb_nome.place(relx=0.35, rely=0.6)

        self.cpf_entry = Entry(self.frame_1)
        #self.cpf_entry.place(relx=0.35, rely=0.7, relwidth=0.2)

        self.lb_nome = Label(self.frame_1, text="Senha", bg='#dfe3ee', fg='#107db2')
        #self.lb_nome.place(relx=0.65, rely=0.6)

        self.senha_entry = Entry(self.frame_1)
        #self.senha_entry.place(relx=0.65, rely=0.7, relwidth=0.2)
    def lista_frame2_const_cnpj(self):
        self.listaEMP = ttk.Treeview(self.frame_2, height=3,
                                     column=("col1", "col2", "col3", "col4", "col5"))
        self.listaEMP.heading("#0", text="")
        self.listaEMP.heading("#1", text="Codigo")
        self.listaEMP.heading("#2", text="Nome")
        self.listaEMP.heading("#3", text="CNPJ")
        self.listaEMP.heading("#4", text="CPF")
        self.listaEMP.heading("#5", text="Senha")
        self.listaEMP.column("#0", width=1)
        self.listaEMP.column("#1", width=50)
        self.listaEMP.column("#2", width=200)
        self.listaEMP.column("#3", width=100)
        self.listaEMP.column("#4", width=100)
        self.listaEMP.column("#5", width=100)
        self.listaEMP.place(relx=0.01, rely=0.1, relwidth=0.95, relheight=0.85)

        self.scroolLista = Scrollbar(self.frame_2, orient='vertical')
        self.listaEMP.configure(yscroll=self.scroolLista.set)
        self.scroolLista.place(relx=0.96, rely=0.1, relwidth=0.04, relheight=0.85)
        self.listaEMP.bind("<Double-1>", self.OnDoubleClick1)

    def janela8(self):
        self.root8 = Toplevel()
        self.listaEMP = None
        self.tela_const_cnd()
        self.frames_da_tela_const_cnd()
        self.widgets_frame1_const_cnd()
        self.lista_frame2_const_cnd()
        self.montaTabelas1()
        self.select_lista1()
        self.root8.transient(self.janela)
        self.root8.resizable(True, True)
        self.root8.focus_force()
        self.root8.grab_set()
    def tela_const_cnd(self):
        self.root8.title("Consultas do municipio")
        self.root8.configure(background='#1e3743')
        self.root8.geometry("700x500")
        self.root8.resizable(True, True)
        self.root8.maxsize(width=900, height=700)
        self.root8.minsize(width=500, height=400)
    def frames_da_tela_const_cnd(self):
        self.frame_1 = Frame(self.root8, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.46)

        self.frame_2 = Frame(self.root8, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)
    def widgets_frame1_const_cnd(self):
        ### Criação do botao consultar cnpj
        self.bt_cnpj = Button(self.frame_1, text="Consultar CND municipal", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.consultar_cnd)
        self.bt_cnpj.place(relx=0.05, rely=0.15, relwidth=0.08, relheight=0.15, width=200)
        ### Criação do botao buscar
        self.bt_buscar = Button(self.frame_1, text="Buscar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.busca_empresa)
        self.bt_buscar.place(relx=0.44, rely=0.15, relwidth=0.1, relheight=0.15)
        ### Criação do botao consultar DAM
        self.bt_buscar = Button(self.frame_1, text="DAM", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.consultar_DAM)
        self.bt_buscar.place(relx=0.545, rely=0.15, relwidth=0.1, relheight=0.15)
        ### Criação para imprimir alvará
        self.bt_buscar = Button(self.frame_1, text="ALVARÁ", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.consultar_alvará)
        self.bt_buscar.place(relx=0.65, rely=0.15, relwidth=0.1, relheight=0.15)
        ### Criação do botao novo
        self.bt_limpar = Button(self.frame_1, text="Novo", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.add_empresa)
        #self.bt_limpar.place(relx=0.6, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao alterar
        self.bt_limpar = Button(self.frame_1, text="Alterar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.altera_empresa)
        #self.bt_limpar.place(relx=0.7, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao apagar
        self.bt_limpar = Button(self.frame_1, text="Apagar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.deleta_empresa)
        #self.bt_limpar.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.15)

        ## Criação da label e entrada do codigo
        self.lb_codigo = Label(self.frame_1, text="Código", bg='#dfe3ee', fg='#107db2')
        #self.lb_codigo.place(relx=0.05, rely=0.05)

        self.codigo1_entry = Entry(self.frame_1)
        #self.codigo1_entry.place(relx=0.05, rely=0.15, relwidth=0.08)

        ## Criação da label e entrada do nome
        self.lb_nome = Label(self.frame_1, text="Empresa", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.35)

        self.empresa_entry = Entry(self.frame_1)
        self.empresa_entry.place(relx=0.05, rely=0.45, relwidth=0.8)

        ## Criação da label e entrada do cnpj
        self.lb_nome = Label(self.frame_1, text="CNPJ", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.6)

        self.cnpj_entry = Entry(self.frame_1)
        self.cnpj_entry.place(relx=0.05, rely=0.7, relwidth=0.2)

        ## Criação da label e entrada da cidade
        self.lb_nome = Label(self.frame_1, text="CPF", bg='#dfe3ee', fg='#107db2')
        #self.lb_nome.place(relx=0.35, rely=0.6)

        self.cpf_entry = Entry(self.frame_1)
        #self.cpf_entry.place(relx=0.35, rely=0.7, relwidth=0.2)

        self.lb_nome = Label(self.frame_1, text="Senha", bg='#dfe3ee', fg='#107db2')
        #self.lb_nome.place(relx=0.65, rely=0.6)

        self.senha_entry = Entry(self.frame_1)
        #self.senha_entry.place(relx=0.65, rely=0.7, relwidth=0.2)
    def lista_frame2_const_cnd(self):
        self.listaEMP = ttk.Treeview(self.frame_2, height=3,
                                     column=("col1", "col2", "col3", "col4"))
        self.listaEMP.heading("#0", text="")
        self.listaEMP.heading("#1", text="Codigo")
        self.listaEMP.heading("#2", text="Nome")
        self.listaEMP.heading("#3", text="CNPJ")
        self.listaEMP.heading("#4", text="CPF")
        #self.listaEMP.heading("#5", text="Senha")
        self.listaEMP.column("#0", width=1)
        self.listaEMP.column("#1", width=50)
        self.listaEMP.column("#2", width=200)
        self.listaEMP.column("#3", width=100)
        self.listaEMP.column("#4", width=100)
        #self.listaEMP.column("#5", width=100)
        self.listaEMP.place(relx=0.01, rely=0.1, relwidth=0.95, relheight=0.85)

        self.scroolLista = Scrollbar(self.frame_2, orient='vertical')
        self.listaEMP.configure(yscroll=self.scroolLista.set)
        self.scroolLista.place(relx=0.96, rely=0.1, relwidth=0.04, relheight=0.85)
        self.listaEMP.bind("<Double-1>", self.OnDoubleClick1)

    def janela9(self):
        self.root9 = Toplevel()
        self.listaEMP = None
        self.tela_const_insc_est()
        self.frames_da_tela_insc_est()
        self.widgets_frame1_const_insc_est()
        self.lista_frame2_const_insc_est()
        self.montaTabelas1()
        self.select_lista1()
        self.root9.transient(self.janela)
        self.root9.resizable(True, True)
        self.root9.focus_force()
        self.root9.grab_set()
    def tela_const_insc_est(self):
        self.root9.title("Consutar Inscrição Estadual")
        self.root9.configure(background='#1e3743')
        self.root9.geometry("700x500")
        self.root9.resizable(True, True)
        self.root9.maxsize(width=900, height=700)
        self.root9.minsize(width=500, height=400)
    def frames_da_tela_insc_est(self):
        self.frame_1 = Frame(self.root9, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.46)

        self.frame_2 = Frame(self.root9, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)
    def widgets_frame1_const_insc_est(self):
        ### Criação do botao consultar cnpj
        self.bt_insc_est = Button(self.frame_1, text="Consultar inscrição estadual", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.consultar_insc_est)
        self.bt_insc_est.place(relx=0.05, rely=0.15, relwidth=0.08, relheight=0.15, width=200)
        ### Criação do botao buscar
        self.bt_buscar = Button(self.frame_1, text="Buscar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.busca_empresa)
        self.bt_buscar.place(relx=0.44, rely=0.15, relwidth=0.1, relheight=0.15)
        ### Criação do botao login simples nacional pgdas
        self.bt_buscar = Button(self.frame_1, text="PGDAS", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.login_simples)
        #self.bt_buscar.place(relx=0.4, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao login simples nacional parcelamento
        self.bt_buscar = Button(self.frame_1, text="Parcelamento", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 6, 'bold'), command=self.login_simples_parcelamento)
       # self.bt_buscar.place(relx=0.5, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao novo
        self.bt_limpar = Button(self.frame_1, text="Novo", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.add_empresa)
        #self.bt_limpar.place(relx=0.6, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao alterar
        self.bt_limpar = Button(self.frame_1, text="Alterar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.altera_empresa)
        #self.bt_limpar.place(relx=0.7, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao apagar
        self.bt_limpar = Button(self.frame_1, text="Apagar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.deleta_empresa)
        #self.bt_limpar.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.15)

        ## Criação da label e entrada do codigo
        self.lb_codigo = Label(self.frame_1, text="Código", bg='#dfe3ee', fg='#107db2')
        #self.lb_codigo.place(relx=0.05, rely=0.05)

        self.codigo1_entry = Entry(self.frame_1)
        #self.codigo1_entry.place(relx=0.05, rely=0.15, relwidth=0.08)

        ## Criação da label e entrada do nome
        self.lb_nome = Label(self.frame_1, text="Empresa", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.35)

        self.empresa_entry = Entry(self.frame_1)
        self.empresa_entry.place(relx=0.05, rely=0.45, relwidth=0.8)

        ## Criação da label e entrada do cnpj
        self.lb_nome = Label(self.frame_1, text="CNPJ", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.6)

        self.cnpj_entry = Entry(self.frame_1)
        self.cnpj_entry.place(relx=0.05, rely=0.7, relwidth=0.2)

        ## Criação da label e entrada da cidade
        self.lb_nome = Label(self.frame_1, text="CPF", bg='#dfe3ee', fg='#107db2')
        #self.lb_nome.place(relx=0.35, rely=0.6)

        self.cpf_entry = Entry(self.frame_1)
        #self.cpf_entry.place(relx=0.35, rely=0.7, relwidth=0.2)

        self.lb_nome = Label(self.frame_1, text="Senha", bg='#dfe3ee', fg='#107db2')
        #self.lb_nome.place(relx=0.65, rely=0.6)

        self.senha_entry = Entry(self.frame_1)
        #self.senha_entry.place(relx=0.65, rely=0.7, relwidth=0.2)
    def lista_frame2_const_insc_est(self):
        self.listaEMP = ttk.Treeview(self.frame_2, height=3,
                                     column=("col1", "col2", "col3", "col4"))
        self.listaEMP.heading("#0", text="")
        self.listaEMP.heading("#1", text="Codigo")
        self.listaEMP.heading("#2", text="Nome")
        self.listaEMP.heading("#3", text="CNPJ")
        self.listaEMP.heading("#4", text="CPF")
        #self.listaEMP.heading("#5", text="Senha")
        self.listaEMP.column("#0", width=1)
        self.listaEMP.column("#1", width=50)
        self.listaEMP.column("#2", width=200)
        self.listaEMP.column("#3", width=100)
        self.listaEMP.column("#4", width=100)
        #self.listaEMP.column("#5", width=100)
        self.listaEMP.place(relx=0.01, rely=0.1, relwidth=0.95, relheight=0.85)

        self.scroolLista = Scrollbar(self.frame_2, orient='vertical')
        self.listaEMP.configure(yscroll=self.scroolLista.set)
        self.scroolLista.place(relx=0.96, rely=0.1, relwidth=0.04, relheight=0.85)
        self.listaEMP.bind("<Double-1>", self.OnDoubleClick1)

    def janela10(self):
        self.root10 = Toplevel()
        self.listaEMP = None
        self.tela_comb_email()
        self.frames_da_tela_comb_email()
        self.widgets_frame1_comb_email()
        self.lista_frame2_comb_email()
        self.montaTabelas4()
        self.select_lista4()
        self.root10.transient(self.janela)
        self.root10.resizable(True, True)
        self.root10.focus_force()
        self.root10.grab_set()
    def tela_comb_email(self):
        self.root10.title("Solicitar arquivos por e-mail")
        self.root10.configure(background='#1e3743')
        self.root10.geometry("800x600")
        self.root10.resizable(True, True)
        self.root10.maxsize(width=900, height=700)
        self.root10.minsize(width=500, height=400)
    def frames_da_tela_comb_email(self):
        self.frame_1 = Frame(self.root10, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.46)

        self.frame_2 = Frame(self.root10, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)
    def widgets_frame1_comb_email(self):
        ### Criação do botao buscar
        self.bt_buscar = Button(self.frame_1, text="Buscar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'),command=self.busca_solicitar_arquivo)
        self.bt_buscar.place(relx=0.3, rely=0.1, relwidth=0.1, relheight=0.15)

        ### Criação do botao enviar
        self.bt_limpar = Button(self.frame_1, text="Enviar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'),command=self.solicita_arquivo_email)
        self.bt_limpar.place(relx=0.2, rely=0.1, relwidth=0.1, relheight=0.15)

        ### Criação do botao novo
        self.bt_limpar = Button(self.frame_1, text="Novo", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.add_solicitar_arquivo)
        self.bt_limpar.place(relx=0.6, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao alterar
        self.bt_limpar = Button(self.frame_1, text="Alterar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.altera_solicitar_arquivo)
        self.bt_limpar.place(relx=0.7, rely=0.1, relwidth=0.1, relheight=0.15)
        ### Criação do botao apagar
        self.bt_limpar = Button(self.frame_1, text="Apagar", bd=2, bg='#107db2', fg='white'
                                , font=('verdana', 8, 'bold'), command=self.deleta_solicitar_arquivo)
        self.bt_limpar.place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.15)

        ## Criação da label e entrada do codigo
        self.lb_codigo = Label(self.frame_1, text="Código", bg='#dfe3ee', fg='#107db2')
        self.lb_codigo.place(relx=0.05, rely=0.05)
        self.codigo4_entry = Entry(self.frame_1)
        self.codigo4_entry.place(relx=0.05, rely=0.15, relwidth=0.08)

        ## Criação da label e entrada do empresa
        self.lb_nome = Label(self.frame_1, text="Empresa", bg='#dfe3ee', fg='#107db2')
        self.lb_nome.place(relx=0.05, rely=0.35)
        self.empresa4_entry = Entry(self.frame_1)
        self.empresa4_entry.place(relx=0.05, rely=0.45, relwidth=0.5)

        ## Criação da label e entrada do email
        self.lb_email = Label(self.frame_1, text="E-mail", bg='#dfe3ee', fg='#107db2')
        self.lb_email.place(relx=0.05, rely=0.55)
        self.email_entry = Entry(self.frame_1)
        self.email_entry.place(relx=0.05, rely=0.65, relwidth=0.5)

        ## Criação da label e entrada da observação
        self.lb_obs = Label(self.frame_1, text="Observação", bg='#dfe3ee', fg='#107db2')
        self.lb_obs.place(relx=0.05, rely=0.75)
        self.obs_entry = Entry(self.frame_1)
        self.obs_entry.place(relx=0.05, rely=0.85, relwidth=0.5)

        ## Criação da label data
        self.lb_data = Label(self.frame_1, text="Dia do mês", bg='#dfe3ee', fg='#107db2')
        self.lb_data.place(relx=0.65, rely=0.35)
        self.data_entry = Entry(self.frame_1)
        self.data_entry.place(relx=0.65, rely=0.45, relwidth=0.2)

        ## Criação da label tipo de arquivo
        self.lb_tipo_arquivo = Label(self.frame_1, text="Tipo do arquivo", bg='#dfe3ee', fg='#107db2')
        self.lb_tipo_arquivo.place(relx=0.65, rely=0.55)
        self.tipo_arquivo_entry = Entry(self.frame_1)
        self.tipo_arquivo_entry.place(relx=0.65, rely=0.65, relwidth=0.2)

        ## Criação da label status
        self.lb_status = Label(self.frame_1, text="Status", bg='#dfe3ee', fg='#107db2')
        self.lb_status.place(relx=0.65, rely=0.75)
        self.status_entry = ttk.Combobox(self.frame_1, values=list_status)
        self.status_entry.place(relx=0.65, rely=0.85, relwidth=0.2)
    def lista_frame2_comb_email(self):
        self.listaEMP = ttk.Treeview(self.frame_2, height=3,
                                     column=("col1", "col2", "col3", "col4", "col5", "col6", "col7"))
        self.listaEMP.heading("#0", text="")
        self.listaEMP.heading("#1", text="Cod")
        self.listaEMP.heading("#2", text="EMPRESA")
        self.listaEMP.heading("#3", text="TIPO DE ARQUIVO")
        self.listaEMP.heading("#4", text="DATA")
        self.listaEMP.heading("#5", text="STATUS")
        self.listaEMP.heading("#6", text="E-MAIL")
        self.listaEMP.heading("#7", text="OBSERVAÇÃO")

        self.listaEMP.column("#0", width=1)
        self.listaEMP.column("#1", width=10)
        self.listaEMP.column("#2", width=100)
        self.listaEMP.column("#3", width=80)
        self.listaEMP.column("#4", width=50)
        self.listaEMP.column("#5", width=50)
        self.listaEMP.column("#6", width=100)
        self.listaEMP.column("#7", width=100)
        self.listaEMP.place(relx=0.01, rely=0.1, relwidth=0.95, relheight=0.85)

        self.scroolLista = Scrollbar(self.frame_2, orient='vertical')
        self.listaEMP.configure(yscroll=self.scroolLista.set)
        self.scroolLista.place(relx=0.96, rely=0.1, relwidth=0.04, relheight=0.85)
        self.listaEMP.bind("<Double-1>", self.OnDoubleClick4)

    def janela11(self):
        self.root11 = Toplevel()
        self.root11.title('Assistente Virtual')
        self.root11.configure(background='#1e3743')
        self.root11.geometry("400x400")  # set the root dimensions
        self.root11.pack_propagate(False)  # tells the root to not let the widgets inside it determine its size.
        self.root11.resizable(0, 0)  # makes the root window fixed in size.
        self.root11.transient(self.janela)
        self.root11.resizable(True, True)
        self.root11.focus_force()
        self.root11.grab_set()
        # Frame da label
        self.frame1 = tk.LabelFrame(self.root11, text="Assistente virtua Meneses Contabilide")
        self.frame1.place(height=450, width=399, rely=0.18, relx=0)
        #
        # Cor de label
        self.file_frame = tk.LabelFrame(self.root11, text="Assistente virtua Meneses Contabilide", bd=4, bg='#dfe3ee',
                                        highlightbackground='#759fe6', highlightthickness=3)
        self.file_frame.place(height=100, width=399, rely=0.01, relx=0)
        #
        # Botão para executar o assistente
        self.button1 = tk.Button(self.file_frame, text="Abrir Chat GPT", bd=2, bg='#107db2', fg='white',
                                 font=('verdana', 8, 'bold'), command=self.Iniciar_Chat_GPT)
        self.button1.place(rely=0.20, relx=0.01)
        # Botão para abrir navegador no chat
        self.button2 = tk.Button(self.file_frame, text="Abrir Chat Bard", bd=2, bg='#107db2', fg='white',
                                 font=('verdana', 8, 'bold'), command=self.Iniciar_Chat_Bard)
        self.button2.place(rely=0.20, relx=0.325)
        # Botão para o chat
        self.button3 = tk.Button(self.file_frame, text="Abrir Chat LuzIa", bd=2, bg='#107db2', fg='white',
                                 font=('verdana', 8, 'bold'), command=self.Iniciar_Chat_LuzIa)
        self.button3.place(rely=0.20, relx=0.65)
        # Texto de instruções sobre o chat
        texto_instrucoes = "Bem-vindo ao Chatbot!\n" \
                           "Eu sou uma inteligência artificial desenvolvida para interagir com os usuários e fornecer informações e assistência em diversas tarefas, como pesquisas, resolução de problemas e outras atividades relacionadas ao meu conhecimento e treinamento. Estou sempre aprendendo e me aprimorando para oferecer um serviço cada vez melhor e mais útil.\n" \
                           "Clique no botão 'Executar assistente' para começar."

        # Label para exibir as instruções
        label_instrucoes = tk.Label(self.root11, text=texto_instrucoes, wraplength=400)
        label_instrucoes.pack(pady=150)

        #self.root11.protocol("WM_DELETE_WINDOW", self.parar_chat)


application()

