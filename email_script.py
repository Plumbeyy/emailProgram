import pandas as pd
from tkinter import *
from tkinter import filedialog
from appscript import app, k

def read_data(file_name):
    data = pd.read_excel(file_name, engine='openpyxl')
    names = list(data.Name)
    emails = list(data.Email)
    email_list = zip(names, emails)
    return email_list

def send_email(email_body, email_list):
    outlook = app('Microsoft Outlook')
    for name, email in email_list:
        msg = outlook.make(
            new=k.outgoing_message,
            with_properties={
                k.subject: 'Test Email',
                k.plain_text_content: f'Dear {name},' + f'\n\n{email_body}'
            }
        )

        msg.make(
            new=k.recipient,
            with_properties={
                k.email_address: {
                    k.name: name,
                    k.address: email
                }
            }
        )
        msg.send()

def run_program(file, body):
    curr_list = read_data(file)
    curr_body = '\n\n' + body
    send_email(curr_body, curr_list)

# curr_list = read_data('funbook.xlsx')
# curr_body = f'I have pain in my stomach'
# send_email(curr_body, curr_list)

def file_clicked():
    global file_select
    file_select = filedialog.askopenfilename()
    selected_label.config(text="Selected File: " + file_select)
    doit_button.config(text="Send Emails", command=lambda: run_program(file_select, mail_body_text.get("1.0", "end-1c")))

main = Tk()

main.title("Email Sender")
main.geometry('500x400')
main.eval('tk::PlaceWindow . center')

file_label = Label(text="Excel File Select")
file_button = Button(text="Select File", command=file_clicked)

file_label.pack()
file_button.pack()

file_select = None

selected_label = Label(text="No file selected")
selected_label.pack()

mail_body_label = Label(text="Email Body")
mail_body_text = Text(height=20, width=300)

mail_body_label.pack()
mail_body_text.pack()

doit_button = Button(text="No file selected", command=None)
doit_button.pack()


main.mainloop()
# msg.open()
# msg.activate()