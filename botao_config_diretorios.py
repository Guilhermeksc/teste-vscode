import tkinter as tk
from diretorios import *
from botoes.config_botao import criar_etiqueta_com_imagem, adicionar_rotulo_descricao, adicionar_traco
from botoes.config_formatacao import adicionar_rotulo_principal
from config.configuracoes import CanvasBackgroundManager

def atualizar_conteudo_config(texto_botao, canvas):
    gerenciador = CanvasBackgroundManager(canvas)
    quadro, _  = gerenciador.set_licitacao_background(canvas, texto_botao)

    config_diretorios(quadro)


def config_diretorios(quadro):
    adicionar_rotulo_principal(quadro, "Configurações de Diretórios", 0)

    adicionar_traco(quadro, 1)    
    adicionar_rotulo_descricao(quadro, "Atualizar Pasta (Termo de Homologação)", 2, padx_value=(70, 0))
    botao_update_pdf_dir = criar_etiqueta_com_imagem(quadro=quadro, caminho_imagem=ICONS_DIR / "rede.png", row=2, tamanho=(2, 2), column=0, padx=(10, 0), sticky=tk.W, comando=lambda: update_dir("Selecione o novo diretório para PDF_DIR", "PDF_DIR", DATABASE_DIR / "pasta_pdf"))

    adicionar_traco(quadro, 3)
    adicionar_rotulo_descricao(quadro, "Atualizar Pasta (SICAF)", 4, padx_value=(70, 0))
    botao_update_sicaf_dir = criar_etiqueta_com_imagem(quadro=quadro, caminho_imagem=ICONS_DIR / "rede.png", row=4, tamanho=(2, 2), column=0, padx=(10, 0), sticky=tk.W, comando=lambda: update_dir("Selecione o novo diretório para SICAF_DIR", "SICAF_DIR", DATABASE_DIR / "pasta_sicaf"))

    adicionar_traco(quadro, 5)  
    adicionar_rotulo_descricao(quadro, "Atualizar Pasta (Templates)", 6, padx_value=(70, 0))
    botao_update_pdf_dir = criar_etiqueta_com_imagem(quadro=quadro, caminho_imagem=ICONS_DIR / "rede.png", row=6, tamanho=(2, 2), column=0, padx=(10, 0), sticky=tk.W, comando=lambda: update_dir("Selecione o novo diretório para PASTA_TEMPLATE", "PASTA_TEMPLATE", BASE_DIR / "template"))

    adicionar_traco(quadro, 7)
    adicionar_rotulo_descricao(quadro, "Atualizar Pasta (Relatório)", 8, padx_value=(70, 0))
    botao_update_pdf_dir = criar_etiqueta_com_imagem(quadro=quadro, caminho_imagem=ICONS_DIR / "rede.png", row=8, tamanho=(2, 2), column=0, padx=(10, 0), sticky=tk.W, comando=lambda: update_dir("Selecione o novo diretório para RELATORIO_PATH", "RELATORIO_PATH", DATABASE_DIR / "relatorio"))

    adicionar_traco(quadro, 9)
    adicionar_rotulo_descricao(quadro, "Atualizar Pasta (Database)", 10, padx_value=(70, 0))
    botao_update_pdf_dir = criar_etiqueta_com_imagem(quadro=quadro, caminho_imagem=ICONS_DIR / "rede.png", row=10, tamanho=(2, 2), column=0, padx=(10, 0), sticky=tk.W, comando=lambda: update_dir("Selecione o novo diretório para DATABASE_DIR", "DATABASE_DIR", BASE_DIR / "database"))
  