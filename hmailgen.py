import os
import random
import string
import win32com.client
import tkinter as tk
from tkinter import messagebox
from configparser import ConfigParser
from datetime import datetime

def generate_email(domain, email_length):
    username = ''.join(random.choices(string.ascii_lowercase, k=email_length))
    return f"{username}@{domain}"

def generate_password(password_length, random_password, default_password):
    if random_password:
        chars = string.ascii_letters + string.digits + string.punctuation
        return ''.join(random.choices(chars, k=password_length))
    else:
        return default_password

def add_email_to_hmailserver(email, password, domain, hmail_username, hmail_password):
    hmailapp = win32com.client.Dispatch("hMailServer.Application")
    hmailapp.Authenticate(hmail_username, hmail_password)
    domain_obj = hmailapp.Domains.ItemByName(domain)
    account = domain_obj.Accounts.Add()
    account.Address = email
    account.Password = password
    account.Active = True
    account.Save()

def generate_emails():
    domain = domain_entry.get()
    hmail_username = hmail_username_entry.get()
    hmail_password = hmail_password_entry.get()
    email_length = length_slider_email.get()
    password_length = length_slider_password.get()
    random_password = password_var.get() == 0
    default_password = password_table_entry.get()
    num_emails = int(quantity_entry.get())

    config = ConfigParser()
    config.read('config.ini')
    config['Settings'] = {
        'Domain': domain,
        'Username': hmail_username,
        'Password': hmail_password,
        'EmailLength': email_length,
        'PasswordLength': password_length,
        'RandomPassword': random_password,
        'DefaultPassword': default_password
    }
    with open('config.ini', 'w') as configfile:
        config.write(configfile)

    result_folder = 'resultados'
    if not os.path.exists(result_folder):
        os.makedirs(result_folder)

    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    result_file = f"{result_folder}/resultados_{timestamp}.txt"

    with open(result_file, 'w') as f:
        for i in range(num_emails):
            email = generate_email(domain, email_length)
            password = generate_password(password_length, random_password, default_password)

            add_email_to_hmailserver(email, password, domain, hmail_username, hmail_password)
            print(f"E-mail adicionado ao hMailServer: {email}:{password}")

            f.write(f"{email}:{password}\n")

            result_text.insert(tk.END, f"E-mail adicionado ao hMailServer: {email}:{password}\n")
    
    messagebox.showinfo("Concluído", "E-mails gerados e adicionados ao hMailServer. Resultados salvos.")

def toggle_password_options():
    if password_var.get() == 0:
        length_slider_password.configure(state="normal")
        password_table_entry.configure(state="disabled")
    else:
        length_slider_password.configure(state="disabled")
        password_table_entry.configure(state="normal")

# Interface gráfica usando Tkinter
window = tk.Tk()
window.title("hMailGen - By Guicker")

# Carregar configurações existentes
config = ConfigParser()
config.read('config.ini')

# Label e Entry para o domínio
domain_label = tk.Label(window, text="Domínio:")
domain_label.pack()
domain_entry = tk.Entry(window)
if 'Settings' in config:
    domain_entry.insert(tk.END, config.get('Settings', 'Domain'))
domain_entry.pack()

# Label e Entry para o usuário do hMailServer
hmail_username_label = tk.Label(window, text="Usuário hMailServer:")
hmail_username_label.pack()
hmail_username_entry = tk.Entry(window)
if 'Settings' in config:
    hmail_username_entry.insert(tk.END, config.get('Settings', 'Username'))
hmail_username_entry.pack()

# Label e Entry para a senha do hMailServer
hmail_password_label = tk.Label(window, text="Senha hMailServer:")
hmail_password_label.pack()
hmail_password_entry = tk.Entry(window, show="*")
if 'Settings' in config:
    hmail_password_entry.insert(tk.END, config.get('Settings', 'Password'))
hmail_password_entry.pack()

# Label e Slider para o tamanho do e-mail
email_length_label = tk.Label(window, text="Tamanho do E-mail:")
email_length_label.pack()
length_slider_email = tk.Scale(window, from_=3, to=10, orient=tk.HORIZONTAL)
if 'Settings' in config:
    length_slider_email.set(int(config.get('Settings', 'EmailLength')))
length_slider_email.pack()

# Label e Slider para o tamanho da senha
password_length_label = tk.Label(window, text="Tamanho da Senha:")
password_length_label.pack()
length_slider_password = tk.Scale(window, from_=3, to=10, orient=tk.HORIZONTAL)
if 'Settings' in config:
    length_slider_password.set(int(config.get('Settings', 'PasswordLength')))
length_slider_password.pack()

# Checkbox para senha padrão e entrada para senha
password_var = tk.IntVar()
password_checkbox = tk.Checkbutton(window, text="Senha Padrão:", variable=password_var, command=toggle_password_options)
password_checkbox.pack()
password_table_entry = tk.Entry(window, show="*", state="disabled")
if 'Settings' in config:
    password_table_entry.insert(tk.END, config.get('Settings', 'DefaultPassword'))
password_table_entry.pack()

# Desabilitar a barra de tamanho da senha quando a opção de senha padrão é selecionada
toggle_password_options()

# Label e Entry para a quantidade de e-mails
quantity_label = tk.Label(window, text="Quantidade de E-mails:")
quantity_label.pack()
quantity_entry = tk.Entry(window)
quantity_entry.pack()

# Botão para gerar e-mails
generate_button = tk.Button(window, text="Gerar E-mails", command=generate_emails)
generate_button.pack()

# Texto de resultado
result_text = tk.Text(window, height=10, width=50)
result_text.pack()

window.mainloop()

