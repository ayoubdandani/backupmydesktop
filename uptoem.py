import os
import shutil
import threading
import zipfile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk


def create_zip_file():
    try:
        zip_button.config(state=tk.DISABLED)
        progress_label.config(text='Creating zip file...')
        progress_bar.start()

        zip_file_name = 'Desktop_Backup.zip'
        max_size = 25 * 1024 * 1024
        with zipfile.ZipFile(zip_file_name, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for file_name in os.listdir(os.path.expanduser('~') + '/Desktop'):
                if os.path.isfile(os.path.join(os.path.expanduser('~') + '/Desktop', file_name)):
                    extension = os.path.splitext(file_name)[1][1:].strip().lower()
                    if extension in ['doc', 'csv', 'xml', 'docx', 'html'] and os.path.getsize(os.path.join(os.path.expanduser('~') + '/Desktop', file_name)) < max_size:
                        zip_file.write(os.path.join(os.path.expanduser('~') + '/Desktop', file_name), file_name)

        progress_bar.stop()
        progress_label.config(text='Zip file created successfully.')

        recipient = recipient_entry.get().strip()
        sender_email = sender_entry.get().strip()
        sender_password = password_entry.get().strip()

        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = recipient
        message['Subject'] = 'Desktop Backup'
        body = 'Please find attached the zip file containing a backup of your desktop files.'
        message.attach(MIMEText(body))

        with open(zip_file_name, 'rb') as file:
            attach = MIMEApplication(file.read(), _subtype='zip')
            attach.add_header('Content-Disposition', 'attachment', filename=zip_file_name)
            message.attach(attach)

        server = smtplib.SMTP(smtp_server_entry.get().strip(), int(port_entry.get().strip()))
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient, message.as_string())
        server.quit()

        messagebox.showinfo(title='Success', message='Email sent successfully.')
        zip_button.config(state=tk.NORMAL)
        progress_label.config(text='Email sent successfully.')
        progress_bar.stop()

    except FileNotFoundError:
        messagebox.showerror(title='Error', message='The Desktop_Backup.zip file could not be found.')
        zip_button.config(state=tk.NORMAL)
        progress_label.config(text='Error.')
        progress_bar.stop()

    except zipfile.BadZipFile:
        messagebox.showerror(title='Error', message='The Desktop_Backup.zip file is invalid.')
        zip_button.config(state=tk.NORMAL)
        progress_label.config(text='Error.')
        progress_bar.stop()

    except smtplib.SMTPAuthenticationError:
        messagebox.showerror(title='Error', message='Authentication error. Please check your email and password.')
        zip_button.config(state=tk.NORMAL)

    finally:
        os.remove(zip_file_name)


def execute_backup():
    threading.Thread(target=create_zip_file).start()


def exit_program():
    window.destroy()


window = tk.Tk()
window.title('Desktop Backup')
window.geometry('400x500')
window.resizable(False, False)

recipient_label = ttk.Label(window, text='Recipient email:')
recipient_label.pack(pady=10)
recipient_entry = ttk.Entry(window, width=40)
recipient_entry.pack()
sender_label = ttk.Label(window, text='Sender email:')
sender_label.pack(pady=10)
sender_entry = ttk.Entry(window, width=40)
sender_entry.pack()

password_label = ttk.Label(window, text='Sender password:')
password_label.pack(pady=10)
password_entry = ttk.Entry(window, width=40, show='*')
password_entry.pack()

smtp_server_label = ttk.Label(window, text='SMTP server:')
smtp_server_label.pack(pady=10)
smtp_server_entry = ttk.Entry(window, width=40)
smtp_server_entry.pack()

port_label = ttk.Label(window, text='Port:')
port_label.pack(pady=10)
port_entry = ttk.Entry(window, width=40)
port_entry.pack()

backup_button = ttk.Button(window, text='Backup', command=execute_backup)
backup_button.pack(pady=20)

zip_button = ttk.Button(window, text='Create zip file', command=create_zip_file)
zip_button.pack(pady=10)


#Progess Bar 
progress_label = ttk.Label(window, text='')
progress_label.pack(pady=10)
progress_bar = ttk.Progressbar(window, orient=tk.HORIZONTAL, length=200, mode='indeterminate')
progress_bar.pack(pady=10)

exit_button = ttk.Button(window, text='Exit', command=exit_program)
exit_button.pack(pady=20)

window.mainloop()

