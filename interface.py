import tkinter
import customtkinter
from tkinter.filedialog import askopenfilename


customtkinter.set_appearance_mode("light")

customtkinter.set_default_color_theme("green")

root = customtkinter.CTk()

root.geometry("500x600")

frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

label = customtkinter.CTkLabel(master=frame, text="Formatador de Planilhas", font=("",18))
label.pack(pady=12, padx=10)

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
insertName = customtkinter.CTkEntry(master=frame, placeholder_text= "Nome da Planilha")
insertName.pack(pady=12, padx=10)

insertDate = customtkinter.CTkEntry(master=frame, placeholder_text= "Data do e-mail")
insertDate.pack(pady=12, padx=10)

insertTime = customtkinter.CTkEntry(master=frame, placeholder_text= "Hora do e-mail")
insertTime.pack(pady=12, padx=10)

def obterValor():
    name = insertName.get()
    date = insertDate.get()
    time = insertTime.get()
    global titulo
    titulo = name + ' - ' + date + ' - ' + time

def searchFile():
    global file
    file = askopenfilename()

def start():
    obterValor()
    import script
    script.activate(titulo)
    

buttonSearch = customtkinter.CTkButton(master=frame, text="Buscar arquivo", command=searchFile)
buttonSearch.pack(pady=12, padx=10)

button = customtkinter.CTkButton(master=frame, text="Pronto!", command=start)
button.pack(pady=12, padx=10)

root.mainloop()


