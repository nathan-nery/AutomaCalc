import customtkinter
# import script

customtkinter.set_appearance_mode("dark")

customtkinter.set_default_color_theme("dark-blue")

root = customtkinter.CTk()

root.geometry("500x350")

frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

label = customtkinter.CTkLabel(master=frame, text="Formatador de Planilhas")
label.pack(pady=12, padx=10)

insertName = customtkinter.CTkEntry(master=frame, placeholder_text= "Nome da planilha")
insertName.pack(pady=12, padx=10)

insertDate = customtkinter.CTkEntry(master=frame, placeholder_text= "Data do e-mail")
insertDate.pack(pady=12, padx=10)

insertTime = customtkinter.CTkEntry(master=frame, placeholder_text= "Hora do e-mail")
insertTime.pack(pady=12, padx=10)

def obterValor():
    name = insertName.get()
    date = insertDate.get()
    time = insertTime.get()
    print(name + '-' + date + '-' + time)
    global titulo
    titulo = name + '-' + date + '-' + time

# Ideia, para inserir a localização, abrir uma interface de busca de arquivo.

button = customtkinter.CTkButton(master=frame, text="Pronto!", command=obterValor)
button.pack(pady=12, padx=10)

root.mainloop()


