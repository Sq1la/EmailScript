import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk

def send_email(smtp_server, port, login, password, from_addr, to_addrs, subject, body, attachments):
    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = ', '.join(to_addrs)
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    for file in attachments:
        attachment = open(file, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {file}')
        msg.attach(part)

    try:
        server = smtplib.SMTP(smtp_server, port)
        server.starttls()
        server.login(login, password)
        server.sendmail(from_addr, to_addrs, msg.as_string())
        messagebox.showinfo("Успех", "Письмо успешно отправлено!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось отправить письмо. Ошибка: {e}")
    finally:
        server.quit()

def send_email_from_gui():
    smtp_service = smtp_var.get().strip().lower()
    if smtp_service == 'yandex':
        smtp_server = 'smtp.yandex.com'
        port = 587
    elif smtp_service == 'mail':
        smtp_server = 'smtp.mail.ru'
        port = 587
    else:
        smtp_server = smtp_entry.get()
        port = int(port_entry.get())

    login = email_entry.get()
    password = password_entry.get()  # Пароль вводится в открытом виде
    from_addr = login
    to_addrs = recipients_entry.get().split(',')
    subject = subject_entry.get()
    body = body_text.get("1.0", tk.END)
    attachments = attachment_list

    send_email(smtp_server, port, login, password, from_addr, to_addrs, subject, body, attachments)

def add_attachment():
    files = filedialog.askopenfilenames()
    for file in files:
        attachment_list.append(file)
        attachment_listbox.insert(tk.END, file)

def remove_attachment():
    selected_indices = attachment_listbox.curselection()
    if not selected_indices:
        messagebox.showwarning("Предупреждение", "Выберите файл для удаления")
        return
    for index in reversed(selected_indices):
        attachment_listbox.delete(index)
        del attachment_list[index]

# Создание графического интерфейса
root = tk.Tk()
root.title("Отправка писем")
root.geometry('600x700')
root.resizable(True, True)  # Устанавливаем возможность изменять размер окна

style = ttk.Style()
style.configure('TFrame', background='#f0f0f0')
style.configure('TLabel', background='#f0f0f0', font=('Helvetica', 12))
style.configure('TButton', font=('Helvetica', 12), padding=10)
style.configure('TEntry', font=('Helvetica', 12), padding=5)

mainframe = ttk.Frame(root, padding="20 20 20 20", style='TFrame')
mainframe.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

ttk.Label(mainframe, text="Почтовый сервис (yandex или mail):", style='TLabel').grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
smtp_var = tk.StringVar()
smtp_entry = ttk.Entry(mainframe, textvariable=smtp_var, style='TEntry')
smtp_entry.grid(row=0, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

ttk.Label(mainframe, text="Адрес SMTP сервера:", style='TLabel').grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
smtp_entry = ttk.Entry(mainframe, style='TEntry')
smtp_entry.grid(row=1, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

ttk.Label(mainframe, text="Порт SMTP сервера:", style='TLabel').grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
port_entry = ttk.Entry(mainframe, style='TEntry')
port_entry.grid(row=2, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

ttk.Label(mainframe, text="Ваш адрес электронной почты:", style='TLabel').grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
email_entry = ttk.Entry(mainframe, style='TEntry')
email_entry.grid(row=3, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

ttk.Label(mainframe, text="Пароль от электронной почты:", style='TLabel').grid(row=4, column=0, padx=10, pady=5, sticky=tk.W)
password_entry = ttk.Entry(mainframe, show="*", style='TEntry')
password_entry.grid(row=4, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

ttk.Label(mainframe, text="Адреса получателей (через запятую):", style='TLabel').grid(row=5, column=0, padx=10, pady=5, sticky=tk.W)
recipients_entry = ttk.Entry(mainframe, style='TEntry')
recipients_entry.grid(row=5, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

ttk.Label(mainframe, text="Тема письма:", style='TLabel').grid(row=6, column=0, padx=10, pady=5, sticky=tk.W)
subject_entry = ttk.Entry(mainframe, style='TEntry')
subject_entry.grid(row=6, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

ttk.Label(mainframe, text="Текст письма:", style='TLabel').grid(row=7, column=0, padx=10, pady=5, sticky=tk.W)
body_text = tk.Text(mainframe, height=10, width=40, font=('Helvetica', 12))
body_text.grid(row=7, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

attachment_list = []
ttk.Label(mainframe, text="Прикрепленные файлы:", style='TLabel').grid(row=8, column=0, padx=10, pady=5, sticky=tk.W)
attachment_listbox = tk.Listbox(mainframe, height=5, width=40, selectmode=tk.MULTIPLE, font=('Helvetica', 12))
attachment_listbox.grid(row=8, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

buttons_frame = ttk.Frame(mainframe, style='TFrame')
buttons_frame.grid(row=9, column=1, padx=10, pady=5, sticky=(tk.W, tk.E))

add_attachment_button = ttk.Button(buttons_frame, text="Добавить файл", command=add_attachment, style='TButton')
add_attachment_button.pack(side=tk.LEFT, padx=5)

remove_attachment_button = ttk.Button(buttons_frame, text="Удалить файл", command=remove_attachment, style='TButton')
remove_attachment_button.pack(side=tk.LEFT, padx=5)

send_button = ttk.Button(mainframe, text="Отправить", command=send_email_from_gui, style='TButton')
send_button.grid(row=10, column=1, padx=10, pady=10, sticky=tk.E)

# Установка веса для адаптивной компоновки
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)
mainframe.grid_rowconfigure(7, weight=1)
mainframe.grid_columnconfigure(1, weight=1)

root.mainloop()