#!/usr/bin/env python
# coding: utf-8

# In[91]:


import pandas as pd
import customtkinter
from tkinter import *
from tkinter import ttk
import tkinter  
from openpyxl import load_workbook
from datetime import datetime
from tkinter import messagebox
import re


# In[92]:


class Login():
    
    def __init__(self,janela,df_usuarios,aba_usuarios):
        self.janela = janela
        self.df_usuarios = df_usuarios
        self.aba_usuarios = aba_usuarios
        self.configuracao_janela_loguin()
        self.tela_loguin()
        
        
    def configuracao_janela_loguin(self):
        customtkinter.set_appearance_mode("dark")
        customtkinter.set_default_color_theme("dark-blue")
        largura_tela = self.janela.winfo_screenwidth()
        altura_tela = self.janela.winfo_screenheight()
        pos_x = int(largura_tela/2 - 400/2)
        pos_y = int(altura_tela/2 - 300/2)
        self.janela.geometry(f'{400}x{300}+{pos_x}+{pos_y}')
        self.janela.title('Login')
        self.janela.resizable('False','False')


    def tela_loguin(self):
        for widget in self.janela.winfo_children():
            widget.destroy()
        texto_loguin = customtkinter.CTkLabel(janela,text='Usuário', font=('Arial Arrow',15,'bold'))
        texto_loguin.place(x=115,y=40)
        texto_senha = customtkinter.CTkLabel(janela,text='Senha', font=('Arial Arrow',15,'bold'))
        texto_senha.place(x=115,y=100)
        self.entry_loguin = customtkinter.CTkEntry(self.janela, placeholder_text='Insira o seu usuário aqui',font=('Arial Arrow', 12), width=170)
        self.entry_loguin.place(x=110,y=65)
        self.entry_senha = customtkinter.CTkEntry(self.janela, placeholder_text='Insira a sua senha aqui', font=('Arial Arrow', 12), width=170,show="*")
        self.entry_senha.place(x=110,y=125)
        self.botao_loguin = customtkinter.CTkButton(self.janela,text='Login',width=80,command=self.fazer_login)
        self.botao_loguin.place(x=110,y=200)
        self.botao_cadastro = customtkinter.CTkButton(self.janela,text='Cadastre-se',width=80,command=self.tela_cadastro)
        self.botao_cadastro.place(x=200,y=200)
        self.checkbox = customtkinter.CTkCheckBox(self.janela,text='Mostrar Senha',font=('Arial Arrow',12,'bold'),command=self.mostrar_senha)
        self.checkbox.place(x=110,y=160)
    
    
    def tela_cadastro(self):
        for widget in self.janela.winfo_children():
            widget.destroy()
        texto_loguin_cadastro = customtkinter.CTkLabel(janela,text='Usuário', font=('Arial Arrow',15,'bold'))
        texto_loguin_cadastro.place(x=115,y=40)
        texto_senha_cadastro = customtkinter.CTkLabel(janela,text='Senha', font=('Arial Arrow',15,'bold'))
        texto_senha_cadastro.place(x=115,y=100)
        texto_confirmar_senha = customtkinter.CTkLabel(janela,text='Confirmar Senha', font=('Arial Arrow',15,'bold'))
        texto_confirmar_senha.place(x=110, y=155)
        self.entry_loguin_cadastro = customtkinter.CTkEntry(self.janela, placeholder_text='Insira o seu usuário aqui',font=('Arial Arrow', 12), width=170)
        self.entry_loguin_cadastro.place(x=110,y=65)
        self.entry_senha_cadastro = customtkinter.CTkEntry(self.janela, placeholder_text='Crie a sua senha aqui', font=('Arial Arrow', 12), width=170,show="*")
        self.entry_senha_cadastro.place(x=110,y=125)
        self.entry_confirmar_senha = customtkinter.CTkEntry(self.janela, placeholder_text='Confirme a sua senha aqui', font=('Arial Arrow', 12), width=170,show="*")
        self.entry_confirmar_senha.place(x=110,y=180)
        self.botao_cadastro = customtkinter.CTkButton(self.janela,text='Cancelar',width=80,command=self.tela_loguin)
        self.botao_cadastro.place(x=110,y=220)
        self.botao_cancelar = customtkinter.CTkButton(self.janela,text='Cadastrar',width=80,command=self.cadastrar_usuario)
        self.botao_cancelar.place(x=200,y=220)
        
    
    def salvar_base_de_dados(self):
        with pd.ExcelWriter('usuarios.xlsx') as writer:
            self.df_usuarios.to_excel(writer,sheet_name=self.aba_usuarios,index=False)
            
    
    def mostrar_senha(self):
        if self.checkbox.get():
            self.entry_senha.configure(show='')
        else:
            self.entry_senha.configure(show='*')
                 
                
    def verificar_senha(self,senha):
        return all([len(senha) >= 8, re.search("[A-Z]", senha), re.search("[0-9]{3,}", senha), re.search("[@#$%^&+=]", senha)])

    
    def cadastrar_usuario(self):
        usuario = self.entry_loguin_cadastro.get()
        senha =self.entry_senha_cadastro.get()
        confirmar_senha = self.entry_confirmar_senha.get()
        cadastro = {'User': usuario,'Senha':senha}
        if usuario in self.df_usuarios['User'].tolist():#transforma pf.Series em uma lista e olha item por item
            tkinter.messagebox.showinfo(title='Erro de Cadastro', message='Esse nome de usuário já existe, tente outro.')
            return
        if not self.verificar_senha(senha):
            mensagem = '''
            A sua senha deve conter:
            - Deve ter no mínimo 8 caracteres
            - Deve ter pelo menos uma letra maiúscula
            - Deve ter pelo menos 3 números
            - Deve ter pelo menos um caractere especial (@, #, $, %, ^, &, +, ou =)
            '''
            tkinter.messagebox.showinfo(title='Erro de Cadastro', message=mensagem)
            return
        if senha != confirmar_senha:
            tkinter.messagebox.showinfo(title='Erro de Cadastro',message='As senhas não condizem. Certifique-se que são iguais.') 
            return
        tkinter.messagebox.showinfo(title='Cadastro',message='Cadastro feito com sucesso!') 
        df_cadastro = pd.DataFrame(cadastro, index=[0])
        self.df_usuarios = pd.concat([self.df_usuarios, df_cadastro], ignore_index=True)
        self.salvar_base_de_dados()
        self.tela_loguin()
        
        
    def fazer_login(self):
        usuario = self.entry_loguin.get()
        senha = self.entry_senha.get()
        if not usuario in self.df_usuarios['User'].tolist():
            tkinter.messagebox.showinfo(title='Erro de Login', message='Usuário não encontrado.')
            return
        if senha != self.df_usuarios.loc[self.df_usuarios['User'] == usuario, 'Senha'].values[0]:
            tkinter.messagebox.showinfo(title='Erro de Login', message='Senha incorreta.')
            return
        tkinter.messagebox.showinfo(title='Login', message='Login feito com sucesso!.')
        self.abrir_controle_estoque()
        
    
    def abrir_controle_estoque(self):
        self.janela.destroy() # Fecha a janela de login
        controle_estoque = Controle_De_Estoque(self.janela,aba_estoque,aba_transacoes)
        controle_estoque.janela.mainloop()


# In[93]:


class Controle_De_Estoque():
    
    def __init__(self,janela,aba_estoque,aba_transacoes):
        self.janela = janela
        self.aba_estoque = aba_estoque
        self.aba_transacoes = aba_transacoes
        self.df_estoque = pd.read_excel('Controle de Estoque.xlsx', sheet_name=self.aba_estoque,engine='openpyxl')
        self.df_transacoes = pd.read_excel('Controle de Estoque.xlsx', sheet_name=self.aba_transacoes,engine='openpyxl')
        self.configuracao_janela()
        self.janela_principal_estoque()
        self.tv_estoque()
        self.janela_principal_transacoes()
        self.tv_transacoes()
        self.display_tv_transacoes()
        self.display_tv_estoque()
    
    
    def configuracao_janela(self):
        customtkinter.set_appearance_mode("dark")
        customtkinter.set_default_color_theme("dark-blue")
        self.janela= customtkinter.CTk()
        self.janela.resizable('False','False')
        frame0 = tkinter.Frame(self.janela,width=955,height=50, bg='#1F538D')
        frame0.place(x=0,y=0)
        texto = customtkinter.CTkLabel(self.janela,text='Controle de Estoque', font=('Arial Arrow',30,'bold'), fg_color='#1F538D')
        texto.pack(padx=10, pady=5)
        self.janela.title('Controle de Estoque')
        self.janela.update_idletasks() # atualiza a janela antes de redimensioná-la
        largura_tela = self.janela.winfo_screenwidth()
        altura_tela = self.janela.winfo_screenheight()
        pos_x = int(largura_tela/2 - 955/2)
        pos_y = int(altura_tela/2 - 540/2)
        self.janela.geometry(f'{955}x{540}+{pos_x}+{pos_y}')
        
        
    def abrir_estoque(self):
        classe_estoque = Controle_De_Produtos(janela,aba_estoque,aba_transacoes)
    
    
    def janela_principal_estoque(self):
        frame = customtkinter.CTkFrame(self.janela,width=340,height=460, fg_color='#1A1A1A',border_color='#767171',border_width=1)
        frame.place(x=10,y=70)
        botao_controle_produtos = customtkinter.CTkButton(self.janela,text='Controle de Produtos',width=100,height=50, font=('Arial Arrow',12,'bold'),command=self.abrir_estoque)
        botao_controle_produtos.place(x=100,y=150)
        botao_procurar_produto = customtkinter.CTkButton(self.janela,text='Procurar',width=100)
        botao_procurar_produto.place(x=240,y=240)
        texto1 = customtkinter.CTkLabel(self.janela,text='Estoque Disponível', font=('Arial Arrow',18,'bold'))
        texto1.place(x=20,y=75)
        texto2 = customtkinter.CTkLabel(self.janela,text='Produto:', font=('Arial Arrow',15))
        texto2.place(x=20,y=240)
        procurar_produto = customtkinter.CTkEntry(self.janela,placeholder_text='Insira o nome do produto',font=('Arial Arrow',12),width=150)
        procurar_produto.place(x=80,y=240)
        
    
    def tv_estoque(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("Treeview", background="#D3D3D3", fieldbackground ='#D3D3D3')
        self.style.map('Treeview',background=[('selected','#1F538D')])
        self.tv_estoque= ttk.Treeview(self.janela,columns=('ID','Produto','Quantidade','Valor'),show='headings', style='Treeview')
        self.tv_estoque.column('ID',minwidth=0,width=70,anchor=N)
        self.tv_estoque.column('Produto',minwidth=0,width=140,anchor=N)
        self.tv_estoque.column('Quantidade',minwidth=0,width=105,anchor=N)
        self.tv_estoque.column('Valor', width=0, stretch=False)
        self.tv_estoque.heading('ID',text='ID')
        self.tv_estoque.heading('Produto',text='Produto')
        self.tv_estoque.heading('Quantidade',text='Quantidade')
        self.tv_estoque.place(x=20,y=290)
        
        
    def display_tv_estoque(self):
        self.tv_estoque.delete(*self.tv_estoque.get_children())
        for index, row in self.df_estoque.iterrows():
            self.tv_estoque.insert('', 'end', values=(row['ID'], row['Produto'], row['Quantidade'], row['Valor']))
        
    
    def janela_principal_transacoes(self):
        frame2 = customtkinter.CTkFrame(self.janela,width=585,height=460, fg_color='#1A1A1A',border_color='#767171',border_width=1)
        frame2.place(x=360,y=70)
        texto3 = customtkinter.CTkLabel(self.janela,text='Registro de Transações', font=('Arial Arrow',18,'bold'))
        texto3.place(x=370,y=75)
        texto4 = customtkinter.CTkLabel(self.janela,text='Produto:', font=('Arial Arrow',15))
        texto4.place(x=370,y=120)
        texto5 = customtkinter.CTkLabel(self.janela,text='Quantidade:', font=('Arial Arrow',15))
        texto5.place(x=580,y=120)
        texto6 = customtkinter.CTkLabel(self.janela,text='Tipo:', font=('Arial Arrow',15))
        texto6.place(x=755,y=120)
        texto7 = customtkinter.CTkLabel(self.janela,text='Valor:', font=('Arial Arrow',15))
        texto7.place(x=370,y=170)
        texto8 = customtkinter.CTkLabel(self.janela,text='Data:', font=('Arial Arrow',15))
        texto8.place(x=520,y=170)
        texto9 = customtkinter.CTkLabel(self.janela,text='ID:', font=('Arial Arrow',15))
        texto9.place(x=675,y=170)
        texto10 = customtkinter.CTkLabel(self.janela,text='Tipo de Transação', font=('Arial Arrow',17))
        texto10.place(x=370,y=210)
        lista_produto = customtkinter.CTkEntry(self.janela,font=('Arial Arrow',12),width=140)
        lista_produto.place(x=430,y=120)
        quantidade_produto = customtkinter.CTkEntry(self.janela,font=('Arial Arrow',12),width=80)
        quantidade_produto.place(x=665,y=120)
        tipo = customtkinter.CTkComboBox(self.janela)
        tipo.place(x=795,y=120)
        tipo.set('')
        tipo.configure(values= ['Compra','Venda'])
        valor_produto = customtkinter.CTkEntry(self.janela,font=('Arial Arrow',12),width=100)
        valor_produto.place(x=415,y=170)
        data = customtkinter.CTkEntry(self.janela,font=('Arial Arrow',12),width=100)
        data.place(x=560,y=170)
        id_produto = customtkinter.CTkEntry(self.janela,font=('Arial Arrow',12),width=100)
        id_produto.place(x=695,y=170)
        adicionar_produto = customtkinter.CTkButton(self.janela,text='Adicionar',width=50)
        adicionar_produto.place(x=805,y=170)
        excluir_produto = customtkinter.CTkButton(self.janela,text='Excluir',width=60)
        excluir_produto.place(x=875,y=170)
        botao_selecionar_compras = customtkinter.CTkButton(self.janela,text='Compras',width=70)
        botao_selecionar_compras.place(x=370,y=240)
        botao_selecionar_vendas = customtkinter.CTkButton(self.janela,text='Vendas',width=70)
        botao_selecionar_vendas.place(x=445,y=240)
        botao_selecionar_todas = customtkinter.CTkButton(self.janela,text='Todas',width=70)
        botao_selecionar_todas.place(x=520,y=240)
        
        
    def tv_transacoes(self):
        self.tv_transacoes= ttk.Treeview(self.janela,columns=('ID Transação','ID','Produto','Quantidade','Tipo','Data','Valor'),show='headings', style='Treeview')
        self.tv_transacoes.column('ID Transação',minwidth=0,width=80,anchor=N)
        self.tv_transacoes.column('ID',minwidth=0,width=70,anchor=N)
        self.tv_transacoes.column('Produto',minwidth=0,width=80,anchor=N)
        self.tv_transacoes.column('Quantidade',minwidth=0,width=80,anchor=N)
        self.tv_transacoes.column('Tipo',minwidth=0,width=80,anchor=N)
        self.tv_transacoes.column('Data',minwidth=0,width=70,anchor=N)
        self.tv_transacoes.column('Valor',minwidth=0,width=100,anchor=N)
        self.tv_transacoes.heading('ID Transação',text='ID Transação')
        self.tv_transacoes.heading('ID',text='ID')
        self.tv_transacoes.heading('Produto',text='Produto')
        self.tv_transacoes.heading('Quantidade',text='Quantidade')
        self.tv_transacoes.heading('Tipo',text='Tipo')
        self.tv_transacoes.heading('Data',text='Data')
        self.tv_transacoes.heading('Valor',text='Valor')
        self.tv_transacoes.place(x=370,y=290)
    
    
    def display_tv_transacoes(self):
        self.tv_transacoes.delete(*self.tv_transacoes.get_children())
        for index, row in self.df_transacoes.iterrows():
            self.tv_transacoes.insert('', 'end', values=(row['ID Transação'],row['ID'], row['Produto'], row['Quantidade'], row['Tipo'], row['Data'], row['Valor']))
    
    
    


# In[94]:


class Controle_De_Produtos():
    
    def __init__(self,janela,aba_estoque,aba_transacoes):
        self.janela2 = customtkinter.CTkToplevel()
        self.aba_estoque = aba_estoque
        self.aba_transacoes = aba_transacoes
        self.df_estoque = pd.read_excel('Controle de Estoque.xlsx', sheet_name=self.aba_estoque,engine='openpyxl')
        self.df_transacoes = pd.read_excel('Controle de Estoque.xlsx', sheet_name=self.aba_transacoes,engine='openpyxl')
        self.configuracao_janela2()
        self.janela_controle_produtos()
        self.tv_controle_produtos()
        self.display_tv_controle_produtos()
        
        
    
    def configuracao_janela2(self):
        customtkinter.set_appearance_mode("dark")
        customtkinter.set_default_color_theme("dark-blue")
        self.janela2.geometry("500x400")
        self.janela2.title('Controle de Produtos')
        self.janela2.resizable('False','False')
        texto = customtkinter.CTkLabel(self.janela2, text='Controle de Produtos', font=('Arial Arrow', 20))
        texto.pack(padx=10, pady=10)
        self.janela2.attributes("-topmost", True)
        self.janela2.lift()
        
    
    
    def janela_controle_produtos(self):
        texto11 = customtkinter.CTkLabel(self.janela2, text='Produto:', font=('Arial Arrow', 15))
        texto11.place(x=20, y=50)
        texto12 = customtkinter.CTkLabel(self.janela2, text='Valor:', font=('Arial Arrow', 15))
        texto12.place(x=210, y=50)
        texto13 = customtkinter.CTkLabel(self.janela2, text='ID:', font=('Arial Arrow', 15))
        texto13.place(x=20, y=90)
        texto14 = customtkinter.CTkLabel(self.janela2, text='Quantidade:', font=('Arial Arrow', 15))
        texto14.place(x=170, y=90)
        self.controle_protudos_produto = customtkinter.CTkEntry(self.janela2, font=('Arial Arrow', 12), width=120)
        self.controle_protudos_produto.place(x=80, y=50)
        self.controle_protudos_valor = customtkinter.CTkEntry(self.janela2, font=('Arial Arrow', 12), width=120)
        self.controle_protudos_valor.place(x=250, y=50)
        self.controle_protudos_id = customtkinter.CTkEntry(self.janela2,placeholder_text='ID Automático' ,font=('Arial Arrow', 12), width=120)
        self.controle_protudos_id.configure(state='disabled')
        self.controle_protudos_id.place(x=45, y=90)
        self.controle_protudos_quantidade = customtkinter.CTkEntry(self.janela2,font=('Arial Arrow', 12), width=120)
        self.controle_protudos_quantidade.place(x=252, y=90)
        self.controle_produtos_adicionar = customtkinter.CTkButton(self.janela2, text='Adicionar', width=100,command=self.adicionar_produtos)
        self.controle_produtos_adicionar.place(x=380, y=50)
        self.controle_produtos_remover = customtkinter.CTkButton(self.janela2, text='Remover', width=100)
        self.controle_produtos_remover.place(x=380, y=90)
        
        
    def tv_controle_produtos(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#D3D3D3", fieldbackground='#D3D3D3')
        style.map('Treeview', background=[('selected', '#1F538D')])
        self.tv_controle_produtos= ttk.Treeview(self.janela2,columns=('ID','Produto','Quantidade','Valor'),show='headings', style='Treeview')
        self.tv_controle_produtos.column('ID',minwidth=0,width=75,anchor=N)
        self.tv_controle_produtos.column('Produto',minwidth=0,width=140,anchor=N)
        self.tv_controle_produtos.column('Quantidade',minwidth=0,width=120,anchor=N)
        self.tv_controle_produtos.column('Valor',minwidth=0,width=120,anchor=N)
        self.tv_controle_produtos.heading('ID',text='ID')
        self.tv_controle_produtos.heading('Produto',text='Produto')
        self.tv_controle_produtos.heading('Quantidade',text='Quantidade')
        self.tv_controle_produtos.heading('Valor',text='Valor')
        self.tv_controle_produtos.place(x=20,y=150)
    
    
    def display_tv_controle_produtos(self):
        self.tv_controle_produtos.delete(*self.tv_controle_produtos.get_children())
        for index, row in self.df_estoque.iterrows():
            self.tv_controle_produtos.insert('', 'end', values=(row['ID'], row['Produto'], row['Quantidade'], row['Valor']))
    
    
    def salvar_base_de_dados(self):
        with pd.ExcelWriter('Controle de Estoque.xlsx') as writer:
            self.df_transacoes.to_excel(writer, sheet_name=self.aba_transacoes, index=False)
            self.df_estoque.to_excel(writer, sheet_name=self.aba_estoque, index=False)
    
    
    def adicionar_produtos(self):
        proximo_id = 'SKU' + str(len(self.df_estoque) + 1).zfill(4)
        valor_formatado = f"R$ {float(self.controle_protudos_valor.get()):.2f}"
        valor_formatado = valor_formatado.replace(".",",")
        novo_produto = {'ID':proximo_id,'Produto': self.controle_protudos_produto.get(),'Quantidade':self.controle_protudos_quantidade.get(),'Valor': valor_formatado}
        df_novo_produto = pd.DataFrame(novo_produto, index=[0])
        self.df_estoque = pd.concat([self.df_estoque, df_novo_produto], ignore_index=True)
        self.salvar_base_de_dados()
        self.display_tv_controle_produtos()
        controle_estoque = Controle_De_Estoque(janela,aba_estoque,aba_transacoes)
        controle_estoque.df_estoque = pd.read_excel('Controle de Estoque.xlsx', sheet_name=aba_estoque,engine='openpyxl')
        controle_estoque.display_tv_estoque()
        tkinter.messagebox.showinfo(title='Cadastro', message='Cadastro feito com sucesso!', parent=self.janela2)
        self.controle_protudos_produto.delete(0, END)
        self.controle_protudos_valor.delete(0, END)
        self.controle_protudos_id.delete(0, END)
        self.controle_protudos_quantidade.delete(0,END)


# In[95]:


aba_transacoes = 'Transacoes'
aba_estoque = 'Estoque'
aba_usuarios = 'usuarios'
df_transacoes = pd.read_excel('Controle de Estoque.xlsx', sheet_name=aba_transacoes,engine='openpyxl')
df_estoque = pd.read_excel('Controle de Estoque.xlsx', sheet_name=aba_estoque,engine='openpyxl')
df_usuarios = pd.read_excel('Usuarios.xlsx',sheet_name=aba_usuarios,engine='openpyxl')
janela= customtkinter.CTk()
Login(janela,df_usuarios,aba_usuarios)
janela.mainloop()

