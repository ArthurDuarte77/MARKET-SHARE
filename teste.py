import tkinter as tk
from tkinter import messagebox

def show_selected():
    selected = []
    if var_jfa.get():
        selected.append("JFA")
    if var_usina.get():
        selected.append("Usina")
    if var_taramps.get():
        selected.append("Taramps")
    if selected:
       print(selected)

# Criação da janela principal
root = tk.Tk()
root.title("Seleção de Modelos")

# Variáveis para armazenar os estados dos checkbuttons
var_jfa = tk.BooleanVar()
var_usina = tk.BooleanVar()
var_taramps = tk.BooleanVar()

# Criando os checkbuttons
check_jfa = tk.Checkbutton(root, text="JFA", variable=var_jfa)
check_usina = tk.Checkbutton(root, text="Usina", variable=var_usina)
check_taramps = tk.Checkbutton(root, text="Taramps", variable=var_taramps)

# Botão para confirmar a seleção
btn_confirm = tk.Button(root, text="Confirmar", command=show_selected)

# Posicionando os elementos na janela
check_jfa.pack(anchor="w")
check_usina.pack(anchor="w")
check_taramps.pack(anchor="w")
btn_confirm.pack(pady=10)

# Inicia o loop principal da aplicação
root.mainloop()
