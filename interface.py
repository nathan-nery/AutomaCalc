import customtkinter
import time
from tkinter.filedialog import askopenfilename
from CTkMessagebox import CTkMessagebox

customtkinter.set_appearance_mode("light")

customtkinter.set_default_color_theme("green")

root = customtkinter.CTk()

root.geometry("500x600")

frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

label = customtkinter.CTkLabel(master=frame, text="Formatador de Planilhas", font=("",18))
label.pack(pady=12, padx=10)

def get_rad():
    my_label.configure(text=radio_var.get())
    global name
    name = radio_var.get()
    
# Botoes
radio_var = customtkinter.StringVar(value="other")

botao1 = customtkinter.CTkRadioButton(master=frame, text="Hapvida", variable= radio_var, value='Hapvida', command=get_rad)
botao1.pack(pady=12, padx=10)

botao2 = customtkinter.CTkRadioButton(master=frame, text="NS1", variable= radio_var, value='NS1', command=get_rad)
botao2.pack(pady=12, padx=10)

my_label = customtkinter.CTkLabel(root, text="")
my_label.pack(pady=10)

"""  Botão rádio para definir qual tipo de formatação deveráser utilizada
label = customtkinter.CTkLabel(master=frame, text="Nome da planilha", )
label.pack(pady=12, padx=10)

def radio():
    if radio_var.get() == "Hapvida":
        
radio_var = tkinter.StringVar(value="other")

botao1 = customtkinter.CTkRadioButton(master=frame, text="Hapvida", variable= radio_var, value='Hapvida')
botao1.pack(pady=12, padx=10)

botao2 = customtkinter.CTkRadioButton(master=frame, text="NS1", variable= radio_var, value='ns1')
botao2.pack(pady=12, padx=10)

botao3 = customtkinter.CTkRadioButton(master=frame, text="GAL", variable= radio_var, value='gal')
botao3.pack(pady=12, padx=10)

botao4 = customtkinter.CTkRadioButton(master=frame, text="Baracchini", variable= radio_var, value='barac')
botao4.pack(pady=12, padx=10)

botao5 = customtkinter.CTkRadioButton(master=frame, text="Listagem de Sinan", variable= radio_var, value='listagem')
botao5.pack(pady=12, padx=10)

"""
#insertName = customtkinter.CTkEntry(master=frame, placeholder_text= "Nome da Planilha")
#insertName.pack(pady=12, padx=10)

insertDate = customtkinter.CTkEntry(master=frame, placeholder_text= "Data do e-mail")
insertDate.pack(pady=12, padx=10)

insertTime = customtkinter.CTkEntry(master=frame, placeholder_text= "Hora do e-mail")
insertTime.pack(pady=12, padx=10)

def obterValor():
    date = insertDate.get()
    time = insertTime.get()
    global titulo
    titulo = name + ' - ' + date + ' - ' + time

def searchFile():
    global file
    file = askopenfilename()

def show_checkmark():
    # Show some positive message with the checkmark icon
    message_box = CTkMessagebox(message="Planilha ajustada!", icon="check", option_1="Fechar")
    message_box.bind("<Option-1>", lambda event: root.destroy())

def start():
    obterValor()
    import script
    script.activate(titulo, name)
    show_checkmark()

buttonSearch = customtkinter.CTkButton(master=frame, text="Buscar arquivo", command=searchFile)
buttonSearch.pack(pady=12, padx=10)

button = customtkinter.CTkButton(master=frame, text="Ajustar Planilha", command=start)
button.pack(pady=12, padx=10)

root.mainloop()


