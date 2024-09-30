import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import imaplib
import time
import pandas as pd
import tkinter as tk
from tkinter import messagebox

# Carregar a planilha de emails
file_path = 'C:/...' #caminho de lista com emails e entidades
emails_df = pd.read_excel(file_path)

def buscar_emails_por_entidade(nome_entidade):
    # Buscar os e-mails da entidade
    resultado = emails_df[emails_df['Unnamed: 0'].str.contains(nome_entidade, case=False, na=False)]
    if not resultado.empty:
        emails = resultado['E-MAIL NFS'].values[0]
        return emails.split(';')  # Separar os e-mails
    else:
        return []

def enviar_email(destinatarios, cco_list, assunto, corpo_email):
    remetente = 'teste@gmail.com.br' #Inserir seu email e senha
    senha = 'senha'

    # Configuração da mensagem
    mensagem = MIMEMultipart()
    mensagem['From'] = remetente
    mensagem['To'] = ', '.join(destinatarios)
    mensagem['Subject'] = assunto
    mensagem['Bcc'] = ', '.join(cco_list)

    mensagem.attach(MIMEText(corpo_email, 'plain'))

    try:
        # Envio do e-mail
        with smtplib.SMTP('email.seudominio.com.br', 587) as servidor: #coloque o site do email, ou o dominio
            servidor.starttls()
            servidor.login(remetente, senha)
            servidor.send_message(mensagem)
            print(f'E-mail enviado para {destinatarios} com sucesso!')
        
        # Salvar na pasta "Enviados"
        salvar_na_pasta_enviados(remetente, senha, mensagem)
        messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")

    except Exception as e:
        print(f'Erro ao enviar e-mail: {e}')
        messagebox.showerror("Erro", f"Erro ao enviar e-mail: {e}")

def salvar_na_pasta_enviados(remetente, senha, mensagem):
    try:
        mail = imaplib.IMAP4_SSL('email.seudominio.com.br')
        mail.login(remetente, senha)
        mail.select('"INBOX.Sent"')
        mail.append('"INBOX.Sent"', '', imaplib.Time2Internaldate(time.time()), str(mensagem).encode('utf-8'))
        mail.logout()
        print('E-mail salvo na pasta Enviados com sucesso!')

    except Exception as e:
        print(f'Erro ao salvar e-mail na pasta Enviados: {e}')
        messagebox.showerror("Erro", f"Erro ao salvar e-mail na pasta Enviados: {e}")

def enviar_email_selecionado():
    # Obter a entidade selecionada
    try:
        entidade = listbox.get(listbox.curselection())
        destinatarios = buscar_emails_por_entidade(entidade)
        
        if destinatarios:
            cco_list = ['...', '...', '...']#Agentes ocutos do email
            assunto = 'Assunto Teste'#Assunto do Email
            corpo_email = """Bom dia, Loremipsu""" #Corpo do Email
            enviar_email(destinatarios, cco_list, assunto, corpo_email)
        else:
            messagebox.showwarning("Aviso", "Entidade não possui e-mails registrados.")
    except:
        messagebox.showwarning("Aviso", "Nenhuma entidade selecionada!")

# Interface gráfica com tkinter
root = tk.Tk()
root.title("Enviar E-mail para Entidades")

# Frame com scrollbar
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

scrollbar = tk.Scrollbar(frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Listbox com scrollbar
listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, selectmode=tk.SINGLE, width=50, height=20)
for entidade in emails_df['Unnamed: 0']:
    listbox.insert(tk.END, entidade)

listbox.pack(side=tk.LEFT, fill=tk.BOTH)
scrollbar.config(command=listbox.yview)

# Botão para enviar e-mail
btn_enviar = tk.Button(root, text="Enviar E-mail", command=enviar_email_selecionado)
btn_enviar.pack(pady=10)

root.mainloop()
