import tkinter as tk
from docx.shared import Pt

def adicionar_traco(quadro, linha):
    canvas_traco = tk.Canvas(quadro, bg="#FFFFFF", width=485, height=2, bd=0, highlightthickness=0)  # Ajuste a largura do Canvas
    canvas_traco.create_line(10, 1, 485, 1, fill="#C7C7C7")  # Ajuste as coordenadas conforme necessário
    canvas_traco.grid(row=linha, column=0, sticky=tk.W, pady=(5, 5))  # Adiciona um pequeno espaçamento vertical

def adicionar_traco_preto(quadro, linha, comprimento=410):
    canvas_traco = tk.Canvas(quadro, bg="#FFFFFF", width=comprimento, height=2, bd=0, highlightthickness=0)
    canvas_traco.create_line(0, 0, comprimento, 0, fill="#000000")
    canvas_traco.grid(row=linha, column=0, sticky=tk.W, pady=(5, 5))

def criar_label(quadro, text, row, column=0, padx=10, pady=0, sticky=tk.W):
    """Helper function to create a label and place it on the grid."""
    lbl = tk.Label(quadro, text=text, font=("Arial", 12, "bold"), bg="#FFFFFF")
    lbl.grid(row=row, column=column, padx=padx, pady=pady, sticky=sticky)
    return lbl

def obter_fonte_13():
    return ("Calibri", 13)

def adicionar_rotulo_instrucao(quadro, texto, linha):
    fonte_negrito = obter_fonte_negrito()
    rotulo = tk.Label(quadro, text=texto, font=fonte_negrito, anchor="w", bg="#FFFFFF")
    rotulo.grid(row=linha, column=0, sticky=tk.W, pady=5)

def adicionar_rotulo_descricao(quadro, texto, linha, padx_value=(0, 0)):
    fonte_normal = obter_fonte_normal()
    rotulo = tk.Label(quadro, text=texto, font=fonte_normal, anchor="w", justify=tk.LEFT, bg="#FFFFFF")
    rotulo.grid(row=linha, column=0, columnspan=2, sticky=tk.W, pady=1, padx=padx_value)

def adicionar_rotulo_diretorio(quadro, diretorio, linha):
    fonte_normal = obter_fonte_normal()
    rotulo = tk.Label(quadro, text=diretorio, font=fonte_normal, anchor="w", justify=tk.LEFT, bg="#FFFFFF")
    rotulo.grid(row=linha, column=0, sticky=tk.W, pady=1)

def carregar_imagem_redimensionada(caminho):
    imagem_original = tk.PhotoImage(file=caminho)
    metade_largura = imagem_original.width() // 2
    metade_altura = imagem_original.height() // 2
    return imagem_original.subsample(2, 2)

def carregar_imagem_xlsx(caminho):
    imagem_original = tk.PhotoImage(file=caminho)
    metade_largura = imagem_original.width() // 8
    metade_altura = imagem_original.height() // 8
    return imagem_original.subsample(8, 8)

def carregar_imagem_icone(caminho, tamanho):
    imagem_original = tk.PhotoImage(file=caminho)
    
    # Obter as dimensões originais da imagem
    largura_original, altura_original = imagem_original.width(), imagem_original.height()
    
    # Obter os fatores de redução
    fator_largura, fator_altura = tamanho
    
    # Usar subsample para reduzir a imagem com base nos fatores fornecidos
    imagem_final = imagem_original.subsample(int(fator_largura), int(fator_altura))
    
    return imagem_final

def criar_botao_imagem(quadro, imagem, *, comando=None):
    botao = tk.Button(quadro, image=imagem, bg="lavender", command=comando)
    botao.image = imagem
    return botao

def obter_fonte_negrito():
    return ("Verdana", 14, "bold")

def obter_fonte_normal():
    return ("Verdana", 14)

def adicione_texto_formatado(paragraph, text, bold=False):
    run = paragraph.add_run(text)
    run.bold = bold
    font = run.font
    font.name = 'Calibri'
    font.size = Pt(12)

def criar_etiqueta_com_imagem(quadro, caminho_imagem, tamanho=(32, 32), comando=None, **opcoes_grid):
    def on_click(event):
        event.widget.config(bg="gray", relief="sunken")
        if comando:
            comando()

    def on_release(event):
        event.widget.config(bg="#FFFFFF", relief="flat")

    img = carregar_imagem_icone(caminho_imagem, tamanho)
    label_img = tk.Label(quadro, image=img, bg="#FFFFFF", bd=2)
    label_img.image = img
    label_img.grid(**opcoes_grid)

    label_img.bind("<Button-1>", on_click)
    label_img.bind("<ButtonRelease-1>", on_release)

    return label_img
