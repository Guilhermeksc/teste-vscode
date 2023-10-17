import tkinter as tk

def adicionar_rotulo_principal(quadro, texto, linha):
    fonte_negrito = format_fonte_negrito()
    rotulo = tk.Label(quadro, text=texto, font=fonte_negrito, anchor=tk.W, bg="#FFFFFF")
    rotulo.grid(row=linha, column=0, padx=(60, 0), sticky=tk.W)

def format_fonte_negrito():
    return ("Verdana", 14, "bold")
