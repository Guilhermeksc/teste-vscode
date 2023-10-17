import tkinter as tk
from tkinter import Canvas, ttk
from botoes.botao_ata import atualizar_conteudo_gerar_atas
from botoes.botao_licitacao import atualizar_conteudo_licitacao
from botoes.config_botao import carregar_imagem_redimensionada
from diretorios import *
from config.configuracoes import CanvasBackgroundManager, BotaoPersonalizado
from botoes.botao_config_diretorios import atualizar_conteudo_config
from PIL import Image, ImageTk

# Configuração da janela principal
root = tk.Tk()
root.title("Projeto Automação em Licitações")

# Obtendo dimensões da tela
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Ajustando a janela para as dimensões da tela
root.geometry("1400x760+0+0")

def criar_gradiente_plano_fundo(width, height, color1, color2):
    # Cria uma nova imagem com o tamanho especificado
    base = Image.new('RGB', (width, height), color1)
    top = Image.new('RGB', (width, height), color2)
    mask = Image.new('L', (width, height))
    mask_data = []
    for y in range(height):
        mask_data.extend([int(255 * (y / height))] * width)
    mask.putdata(mask_data)
    base.paste(top, (0, 0), mask)
    return base

# Permitindo que a janela seja redimensionada
root.resizable(True, True)

# Criando o gradiente com as dimensões da tela
gradient = criar_gradiente_plano_fundo(screen_width, screen_height, "#000C7B", "#000000")
gradient_photo = ImageTk.PhotoImage(gradient)
gradient_label = tk.Label(root, image=gradient_photo)
gradient_label.place(x=0, y=0, relwidth=1, relheight=1)

canvas_height = 710
canvas = Canvas(root, bg="white", bd=1, highlightthickness=0) 
canvas.place(x=250, y=25, width=1130, height=canvas_height)

# Aqui criamos a instância do gerenciador de plano de fundo
background_manager = CanvasBackgroundManager(canvas)
background_manager.set_inicio_background()

def criar_botoes():
    comandos_botoes = {
        "Início": background_manager.set_inicio_background,
        "Licitação": lambda t="Licitação": atualizar_conteudo_licitacao(t, canvas),
        "Gerar Atas": lambda t="Gerar Atas": atualizar_conteudo_gerar_atas(t, canvas),
        "Configurações": lambda t="Configurações": atualizar_conteudo_config(t, canvas)
    }

    for i, (text, command) in enumerate(comandos_botoes.items()):
        BotaoPersonalizado(root, text, command, x=5, y=25 + (i * (50 + 25)), width=235, height=50)

criar_botoes()

root.mainloop()
