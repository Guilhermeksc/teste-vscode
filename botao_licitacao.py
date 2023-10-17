import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from diretorios import *
from botoes.config_botao import adicionar_traco_preto, criar_label, obter_fonte_13, adicionar_traco, adicione_texto_formatado, adicionar_rotulo_descricao, adicionar_rotulo_diretorio, adicionar_rotulo_instrucao, criar_botao_imagem, carregar_imagem_redimensionada, carregar_imagem_icone, obter_fonte_negrito, obter_fonte_normal
from config.configuracoes import CanvasBackgroundManager
import pandas as pd
import os

df_licitacao = None
tabela_licitacao = None
df_pregao_escolhido = None

def atualizar_conteudo_licitacao(texto_botao, canvas):
    gerenciador = CanvasBackgroundManager(canvas)
    quadro, quadro_novo = gerenciador.set_licitacao_background(canvas, texto_botao)

    licitacao(quadro)
    processos(quadro_novo)
    exibir_dataframe(quadro)

def criar_planilha_vazia():
    colunas = [
        "num_pregao", "ano_pregao", "nup", "objeto", "uasg", 
        "orgao_responsavel", "sigla_om", "setor_responsavel", "portaria",
        "coordenador_planejamento", "grad_coord_plan", "memb1_plan",
        "grad_memb1_plan", "memb2_plan", "grad_memb2_plan", "email",
        "telefone", "srp", "variavel_q", "variavel_e", "variavel_t"
    ]

    # Crie um DataFrame vazio com as colunas fornecidas
    df = pd.DataFrame(columns=colunas)

    file_path = filedialog.asksaveasfilename(initialdir=BASE_DIR, defaultextension=".xlsx", initialfile="planilha_vazia.xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])

    if not file_path:  # o usuário cancelou o diálogo
        return

    # Salve o DataFrame como um arquivo .xlsx
    df.to_excel(file_path, index=False)

def selecionar_arquivo_e_carregar(quadro):
    global df_licitacao, df_licitacao_carregado

    # Restringindo a seleção de arquivos para .xlsx diretamente no diálogo
    arquivo = filedialog.askopenfilename(initialdir=BASE_DIR, 
                                     filetypes=[("Spreadsheet files", ("*.xlsx", "*.xls", "*.ods")), 
                                                ("Excel files", ("*.xlsx", "*.xls")), 
                                                ("LibreOffice files", "*.ods")])

    if arquivo:
        df_licitacao_carregado = pd.read_excel(arquivo)
        df_licitacao = df_licitacao_carregado[['num_pregao', 'ano_pregao', 'objeto', 'sigla_om']]
        exibir_dataframe(quadro, df_licitacao)

nomes_das_colunas = {
    'num_pregao': 'Nº',
    'ano_pregao': 'Ano',
    'objeto': 'Objeto',
    'sigla_om': 'OM'
}

ajustar_colunas_especificas = {
    'num_pregao': 30,
    'ano_pregao': 45,
    'sigla_om': 95 
}

def exibir_dataframe(quadro, df_licitacao=None):
    global tabela_licitacao

    if df_licitacao is None:
        return

    # Filtering the dataframe columns
    df_licitacao = df_licitacao[['num_pregao', 'ano_pregao', 'objeto', 'sigla_om']]

    # Limpar o Treeview existente
    for i in tabela_licitacao.get_children():
        tabela_licitacao.delete(i)

    tabela_licitacao["columns"] = df_licitacao.columns.tolist()

    for col in df_licitacao.columns:
        display_name = nomes_das_colunas.get(col, col)
        tabela_licitacao.heading(col, text=display_name)
        tabela_licitacao.column(col, width=ajustar_colunas_especificas.get(col, 280))

    tabela_licitacao["show"] = "headings"

    for _, row in df_licitacao.iterrows():
        tabela_licitacao.insert("", "end", values=row.tolist())
        
    # Adicionando uma barra de rolagem vertical
    scrollbar = ttk.Scrollbar(quadro, orient="vertical", command=tabela_licitacao.yview)
    scrollbar.grid(row=3, column=0, sticky='ns', padx=470)  
    tabela_licitacao.configure(yscrollcommand=scrollbar.set)

    # Configurar a fonte do Treeview
    fonte = obter_fonte_13()
    style = ttk.Style()
    style.configure("Treeview", font=fonte)

    tabela_licitacao.bind("<Double-1>", lambda e: escolher_item(quadro))

escolha_anterior_label=None

rótulos_anteriores = []

def escolher_item(quadro):
    global df_licitacao_carregado, df_pregao_escolhido, rótulos_anteriores

    # Safety check to ensure df_licitacao_carregado is not None
    if df_licitacao_carregado is None:
        messagebox.showwarning("Aviso", "Por favor, carregue o arquivo primeiro!")
        return

    item_selecionado = tabela_licitacao.selection()
    
    # If no item is selected, show a warning message and return
    if not item_selecionado:
        messagebox.showwarning("Aviso", "Por favor, selecione um item primeiro!")
        return
    
    # Get the item values
    item_values = tabela_licitacao.item(item_selecionado, "values")
    
    # Extract num_pregao value
    num_pregao = int(item_values[0])
    ano_pregao = int(item_values[1])
    objeto = (item_values[2])
    sigla_om = (item_values[3])

    # Filter all rows from df_licitacao_carregado with the selected num_pregao and assign to df_pregao_escolhido
    df_pregao_escolhido = df_licitacao_carregado[df_licitacao_carregado['num_pregao'] == num_pregao]
    print(df_pregao_escolhido)
    
    for lbl in rótulos_anteriores:
        lbl.destroy()
    
    # Clear the list
    rótulos_anteriores.clear()

    label_texts = [
        f"Pregão Eletrônico nº {num_pregao}/{ano_pregao} escolhido!",
        f"Objeto: {objeto}.",
        f"OM: {sigla_om}."
    ]
           
    for idx, label_text in enumerate(label_texts, start=5):
        lbl_item_escolhido = criar_label(quadro, label_text, idx)
        rótulos_anteriores.append(lbl_item_escolhido)

def adicionar_rotulo_licitacao(quadro, texto, linha):
    fonte_negrito = obter_fonte_negrito()
    rotulo = tk.Label(quadro, text=texto, font=fonte_negrito, anchor=tk.W, bg="#FFFFFF")
    rotulo.grid(row=linha, column=0, padx=(60, 0), sticky=tk.W)

def adicionar_subrotulo_licitacao(quadro, texto, linha, fonte=None):
    rotulo = tk.Label(quadro, text=texto, font=fonte, anchor=tk.W, bg="#FFFFFF")
    rotulo.grid(row=linha, column=0, padx=(60, 0), sticky=tk.W)

from docxtpl import DocxTemplate

def criar_pasta_e_salvar_docx(df, template_path, salvar_nome):
    num_pregao = df['num_pregao'].iloc[0]
    ano_pregao = df['ano_pregao'].iloc[0]

    # Usando formatação de strings para clareza
    nome_dir_principal = f"PE {num_pregao}-{ano_pregao}"
    path_dir_principal = RELATORIO_PATH / nome_dir_principal
    if not path_dir_principal.exists():
        path_dir_principal.mkdir(parents=True)

    path_subpasta = path_dir_principal / salvar_nome
    if not path_subpasta.exists():
        path_subpasta.mkdir()

    nome_do_arquivo = f"PE {num_pregao}-{ano_pregao} - {salvar_nome}.docx"
    local_para_salvar = path_subpasta / nome_do_arquivo
    gerar_documento_com_dados(df, template_path, local_para_salvar)

def gerar_documento_com_dados(df, template_path, save_path):
    doc = DocxTemplate(template_path)
    # Converter o dataframe para um dicionário
    data = df.iloc[0].to_dict()
    # Usar o dicionário para substituir as variáveis no template
    doc.render(data)
    # Salvar o documento gerado
    doc.save(save_path)

def gerar_txt_com_dados(df, template_path, save_path, salvar_nome):
    """
    Gera um arquivo .txt substituindo as variáveis no template com os valores do dataframe.
    """
    with open(template_path, 'r', encoding='utf-8') as file:
        template_content = file.read()
    
    # Converter o dataframe para um dicionário
    data = df.iloc[0].to_dict()
    data['tipo_documento'] = salvar_nome  # Use the 'salvar_nome' argument directly
    
    # Replace the variables in the template
    content = template_content.format(**data)
    
    # Save the generated .txt file
    with open(save_path, 'w', encoding='utf-8') as file:
        file.write(content)

def efeito_visual_apos_clique(evento):
    """Aplica um efeito visual ao widget após um clique."""
    evento.widget.config(bg="gray", relief="sunken")
    evento.widget.update()  # Atualiza o widget para mostrar a mudança visual imediatamente
    evento.widget.after(100)  # Adiciona um atraso de 100ms
    evento.widget.config(bg="#FFFFFF", relief="flat")

def ao_clicar_no_icone_documento(nome_template, salvar_nome, evento):
    """Handle the click on the document icon."""
    efeito_visual_apos_clique(evento)

    if df_pregao_escolhido is None:
        root = tk.Tk()  # Cria uma janela raiz do tkinter
        root.withdraw()  # Esconde a janela raiz
        messagebox.showwarning("Aviso", "Por favor, selecione um item primeiro!")
        return  # Encerra a função para não continuar o processamento
        
    if df_pregao_escolhido is not None:
        caminho_template_docx = PASTA_TEMPLATE / f'template_{nome_template}.docx'
        criar_pasta_e_salvar_docx(df_pregao_escolhido, caminho_template_docx, salvar_nome)

        # Generating and saving the .txt file
        caminho_template_txt = PASTA_TEMPLATE / 'SIGDEM.txt'
        num_pregao = df_pregao_escolhido['num_pregao'].iloc[0]
        ano_pregao = df_pregao_escolhido['ano_pregao'].iloc[0]
        nome_do_arquivo_txt = f"PE {num_pregao}-{ano_pregao} - {salvar_nome}.txt"
        path_dir_principal = RELATORIO_PATH / f"PE {num_pregao}-{ano_pregao}"
        path_subpasta = path_dir_principal / salvar_nome
        local_para_salvar_txt = path_subpasta / nome_do_arquivo_txt
        
        gerar_txt_com_dados(df_pregao_escolhido, caminho_template_txt, local_para_salvar_txt, salvar_nome)

def ao_clicar_no_icone_pasta(evento, df):
    """Manipula o clique no ícone da pasta."""
    efeito_visual_apos_clique(evento)
    if df is None:
        messagebox.showwarning("Atenção", "Por favor, selecione um item primeiro.")
        return
    
    num_pregao = df['num_pregao'].iloc[0]
    ano_pregao = df['ano_pregao'].iloc[0]
    nome_diretorio_principal = f"PE {num_pregao}-{ano_pregao}"
    caminho_diretorio_principal = RELATORIO_PATH / nome_diretorio_principal

    if not caminho_diretorio_principal.exists():
        # Cria a pasta se ela não existir
        caminho_diretorio_principal.mkdir(parents=True)

    # Abre a pasta
    os.startfile(caminho_diretorio_principal)

def adicionar_subrotulo_e_icones(quadro, texto, row, template_nome, salvar_nome, fonte=None):
    adicionar_subrotulo_licitacao(quadro, texto, row, fonte)
    
    label_doc = criar_label_com_imagem(quadro, ICONS_DIR / "docx (1).png", row=row, tamanho=(1, 1), column=0, padx=(0, 0), sticky=tk.W)
    label_doc.bind("<Button-1>", lambda evento: ao_clicar_no_icone_documento(template_nome, salvar_nome, evento))

    label_pdf = criar_label_com_imagem(quadro, ICONS_DIR / "pdf.png", row=row, tamanho=(1, 1), column=0, padx=(30, 0), sticky=tk.W)

subrotulo_dados = [
    ("Autorização para Abertura de Licitação", 5, 'aut'),
    ("Declaração de Adequação Orçamentária", 6, 'dec_orc'),
    ("Documento de Formalização de Demanda (DFD)", 7, 'dfd'),
    ("Portaria de Equipe de Planejamento", 8, 'port'),
    ("Estudo Técnico Preliminar (ETP)", 9, 'etp'),
    ("Matriz de Riscos", 10, 'matriz'),
    ("Intenção de Registro de Preços (IRP)", 11, 'irp'),        
    ("Termo de Referência", 12, 'tr'),
    ("Edital", 13, 'edital'),
    ("CP CJACM", 14, 'cp_cjacm'),
    ("Despacho CJACM", 15, 'despacho_cjacm'),
]

def processos(quadro):
    adicionar_rotulo_licitacao(quadro, "Abrir Pasta", 0)
    label_pasta = criar_label_com_imagem(quadro, ICONS_DIR / "pasta.png", row=0, tamanho=(12,12), column=0, padx=(00, 0), sticky=tk.W)
    label_pasta.bind("<Button-1>", lambda event: ao_clicar_no_icone_pasta(event, df_pregao_escolhido))

    adicionar_traco_preto(quadro, 1) 
    adicionar_rotulo_licitacao(quadro, "Processo Completo", 3)
    
    label_processo = criar_label_com_imagem(quadro, ICONS_DIR / "processo.png", row=3, tamanho=(12,12), column=0, padx=(00, 0), sticky=tk.W)
    label_processo.bind("<Button-1>", gerar_todos_documentos)  # Bind the new function here
    adicionar_traco_preto(quadro, 4) 

    # Lista de dados para adicionar subrótulos e ícones
    fonte_calibri_13 = obter_fonte_13()

    for texto, row, nome_template in subrotulo_dados:
        adicionar_subrotulo_e_icones(quadro, texto, row, nome_template, texto, fonte=fonte_calibri_13)

    adicionar_traco_preto(quadro, 16) 
    adicionar_rotulo_licitacao(quadro, "Lista de Verificação", 17)
    from botoes.lista_de_verificacao import main as verificar_lista
    def on_click(event):
        event.widget.config(bg="gray", relief="sunken")

    label_checklist = criar_label_com_imagem(quadro, ICONS_DIR / "check-list.png", tamanho=(10,10), row=17, column=0, padx=(0, 0), sticky=tk.W)
    label_checklist.bind("<Button-1>", lambda e: (on_click(e), verificar_lista()))

    adicionar_traco_preto(quadro, 18) 
    adicionar_rotulo_licitacao(quadro, "Nota Técnica", 19)

    label_pasta = criar_label_com_imagem(quadro, ICONS_DIR / "compliant.png", tamanho=(12,12), row=19, column=0, padx=(0, 0), sticky=tk.W)

def criar_label_com_imagem(quadro, caminho_imagem, tamanho=(32, 32), comando=None, **opcoes_grid):
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

def criar_botoes_analise_para_licitacao(quadro):
    imagem_import = criar_label_com_imagem(quadro, ICONS_DIR / "import.png", row=0, tamanho=(2,2), column=0, padx=(160, 0), sticky=tk.W, comando=lambda: selecionar_arquivo_e_carregar(quadro))
    imagem_download = criar_label_com_imagem(quadro, ICONS_DIR / "planilha_vazia.png", row=0, tamanho=(2,2), column=0, padx=(360, 0), sticky=tk.W, comando=lambda: criar_planilha_vazia())

    adicionar_rotulo_descricao(quadro, "Planilha vazia:", 0, padx_value=(215, 0))
    adicionar_rotulo_descricao(quadro, "Importar dados:", 0, padx_value=(0, 0))
    adicionar_traco(quadro, 1)
    adicionar_traco(quadro, 4)  

def configurar_visualizacao_tabela_para_licitacao(quadro):
    global tabela_licitacao
    tabela_licitacao = ttk.Treeview(quadro, height=20)
    tabela_licitacao.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky='w')  # Alinhado à esquerda com sticky='w'

    # Definir as colunas e seus tamanhos mesmo quando o Treeview está vazio
    tabela_licitacao["columns"] = list(nomes_das_colunas.keys())
    for col in tabela_licitacao["columns"]:
        display_name = nomes_das_colunas.get(col, col)
        tabela_licitacao.heading(col, text=display_name)
        tabela_licitacao.column(col, width=ajustar_colunas_especificas.get(col, 280))
    tabela_licitacao["show"] = "headings"

    # Adicionar barra de rolagem
    scrollbar = ttk.Scrollbar(quadro, orient="vertical", command=tabela_licitacao.yview)
    scrollbar.grid(row=3, column=0, sticky='ns', padx=470)
    tabela_licitacao.configure(yscrollcommand=scrollbar.set)

def exibir_tabela_licitacao(quadro):
    if df_licitacao is not None:
        exibir_dataframe(quadro, df_licitacao)

def licitacao(quadro):
    criar_botoes_analise_para_licitacao(quadro)
    configurar_visualizacao_tabela_para_licitacao(quadro)
    exibir_tabela_licitacao(quadro)

def gerar_todos_documentos(evento=None):
    # Gerar todos os documentos listados em subrotulo_dados.
    if df_pregao_escolhido is None:
        messagebox.showwarning("Atenção", "Por favor, selecione um pregão primeiro.")
        return

    # Iterate sobre subrotulo_dados e gerar cada documento
    for texto, _, nome_template in subrotulo_dados:
        ao_clicar_no_icone_documento(nome_template, texto, evento)
