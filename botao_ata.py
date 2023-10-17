import tkinter as tk
from tkinter import ttk, font, messagebox
from tkinter import font as tkFont
from diretorios import *
from botoes.config_botao import  adicionar_traco_preto, adicione_texto_formatado, adicionar_rotulo_descricao, adicionar_rotulo_diretorio, adicionar_rotulo_instrucao, criar_botao_imagem, carregar_imagem_redimensionada, obter_fonte_negrito, obter_fonte_normal
import os
import pandas as pd
from num2words import num2words
import json
from pathlib import Path
from config.configuracoes import CanvasBackgroundManager
from botao_ata.processar_sicaf import convert_pdf_to_txt, processar_arquivos_sicaf
from botao_ata.utilidades import convert_pdf_to_txt, ler_arquivos_txt, obter_arquivos_txt, save_to_excel
from config.formatacao import adicionar_rotulo, criar_label_com_imagem, carregar_imagem_icone, ao_clicar_no_icone_pasta, efeito_visual_apos_clique, abrir_pasta, ao_clicar_no_label
import re
import locale
# Configurar o locale para o formato brasileiro
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

NUMERO_ATA_GLOBAL = None
GERADOR_NUMERO_ATA = None
CSV_OUTPUT_PATH = DATABASE_DIR / "dados.csv"
tr_variavel_df_carregado = None

info_template = """UASG: {}
Pregão Eletrônico: {}{}
Objeto: {}
Total de itens: {}
Não Concluídos: {}
Concluídos: {}
        Homologado: {}
        Deserto: {}
        Fracassado: {}
Valor Total Homologado: R$ {}"""

nomes_das_colunas = {
    'item_num': 'Item',
    'descricao_tr': 'Descrição',
    'situacao': 'Situação',
    'cnpj': 'CNPJ',
    'empresa': 'Empresa'
}

ajustar_colunas_especificas = {
    'item_num': 50,
    'descricao_tr': 240,
    'situacao': 210, 
    'cnpj': 190, 
    'empresa': 310
}

import tkinter.font

def atualizar_conteudo_gerar_atas(texto_botao, canvas):
    gerenciador = CanvasBackgroundManager(canvas)
    quadro, _ = gerenciador.set_licitacao_background(canvas, texto_botao)

    # Chama a função para criar os quadros e receber o quadro criado
    quadro_novo_ata = gerenciador.criar_quadro_ata(canvas)

    numero_ata_entrada = tk.Entry(quadro)
    criar_passos(quadro, numero_ata_entrada, GERADOR_NUMERO_ATA, quadro_novo_ata)

    # Modificando o INFO_TEMPLATE para inserir os 6 espaços vazios
    partes = info_template.split("{}")
    info_text = ("      ").join(partes[:-1]) + partes[-1]

    # Exibindo o template sem os valores
    bold_font = tkinter.font.Font(family="Verdana", size=12, weight="bold")
    label_info = tk.Label(quadro_novo_ata, text=info_text, font=bold_font, anchor='w', justify=tk.LEFT, bg="#FFFFFF")
    label_info.grid(row=0, column=0, sticky='w', pady=0, padx=0)

    titulo = tk.Label(quadro_novo_ata, text="Análise - Termo de Homologação", font=('Verdana', 15, 'bold'), bg="white")
    titulo.grid(row=2, column=0, sticky='ew', columnspan=2)  # Ajuste o valor de columnspan conforme necessário

    style = ttk.Style()
    style.configure('Treeview', font=('Verdana', 12, 'bold'))
    style.configure('Treeview.Heading', font=('Verdana', 12))

    # Inicialize e configure o Treeview
    global dataframe_licitacao
    dataframe_licitacao = ttk.Treeview(quadro_novo_ata, height=10)
    dataframe_licitacao["columns"] = ("item_num", "descricao_tr", "situacao")
    dataframe_licitacao.column("#0", width=0, stretch=tk.NO)  # Remova o espaço reservado para a imagem

    for col in dataframe_licitacao["columns"]:
        display_name = nomes_das_colunas.get(col, col)
        dataframe_licitacao.heading(col, text=display_name)
        dataframe_licitacao.column(col, width=ajustar_colunas_especificas.get(col, 240), stretch=tk.NO)

    # Centralizando o conteúdo das colunas "item_num"
    dataframe_licitacao.column("item_num", anchor="center")
    dataframe_licitacao.grid(row=3, column=0, pady=0, padx=10, sticky='nsew') 

    titulo = tk.Label(quadro_novo_ata, text="Análise - SICAF", font=('Verdana', 15, 'bold'), bg="white")
    titulo.grid(row=6, column=0, sticky='ew', columnspan=2)  # Posicione na linha acima do treeview

        # Inicialize e configure o Treeview
    global dataframe_sicaf
    dataframe_sicaf = ttk.Treeview(quadro_novo_ata, height=7)
    dataframe_sicaf["columns"] = ("cnpj", "empresa")
    dataframe_sicaf.column("#0", width=0, stretch=tk.NO)  # Remova o espaço reservado para a imagem

    for col in dataframe_sicaf["columns"]:
        display_name = nomes_das_colunas.get(col, col)
        dataframe_sicaf.heading(col, text=display_name)
        dataframe_sicaf.column(col, width=ajustar_colunas_especificas.get(col, 240), stretch=tk.NO)
    
    dataframe_sicaf.grid(row=7, column=0, pady=0, padx=10, sticky='nsew')


padrao_1 = (r"UASG\s+(?P<uasg>\d+)\s+-\s+(?P<orgao_responsavel>.+?)\s+PREGÃO\s+(?P<num_pregao>\d+)/(?P<ano_pregao>\d+)")
padrao_srp = r"(?P<srp>SRP - Registro de Preço|SISPP - Tradicional)"
padrao_objeto = (r"Objeto da compra:\s*(?P<objeto>.*?)\s*Entrega de propostas:")
padrao_grupo2 = (r"Item\s+(?P<item_num>\d+)\s+do\s+Grupo\s+G(?P<grupo>\d+).+?Valor\s+estimado:\s+R\$\s+(?P<valor>[\d,.]+).+?Critério\s+de\s+julgamento:\s+(?P<crit_julgamento>.+?)\s+Quantidade:\s+(?P<quantidade>\d+)\s+Unidade\s+de\s+fornecimento:\s+(?P<unidade>[^S]+?)\s+Situação:\s+(?P<situacao>Adjudicado e Homologado|Deserto e Homologado|Fracassado e Homologado)")
padrao_item2 = (r"Item\s+(?P<item_num>\d+)\s+-\s+.*?Quantidade:\s+(?P<quantidade>\d+)\s+Valor\s+estimado:\s+R\$\s+(?P<valor>[\d,.]+)\s+Unidade\s+de\s+fornecimento:\s+(?P<unidade>.+?)\s+Situação:\s+(?P<situacao>Adjudicado e Homologado|Deserto e Homologado|Fracassado e Homologado)")
padrao_3 = (
    r"Adjucado e Homologado por CPF (?P<cpf_od>\*\*\*.\d{3}.\*\*\*-\*\d{1})\s+-\s+"
    r"(?P<ordenador_despesa>.+?)\s+para\s+"
    r"(?P<empresa>.+?)\s*,\s+CNPJ\s+"
    r"(?P<cnpj>\d{2}\s*\.\s*\d{3}\s*\.\s*\d{3}\s*/\s*\d{4}\s*-\s*\d{2}),\s+"
    r"melhor lance:\s*(?:[\d,]+%\s*\()?"
    r"R\$ (?P<melhor_lance>[\d,.]+)(?:\))?(?:,\s+"
    r"valor negociado: R\$ (?P<valor_negociado>[\d,.]+))?\s+Propostas do Item"
)
padrao_4 = (r"Proposta adjudicada.*? Marca/Fabricante:(?P<marca_fabricante>.*?) Modelo/versão:(?P<modelo_versao>.*?)(?=\d{2}/\d{2}/\d{4}|\s*Valor proposta:)")

def encontre_valor_ou_NA(item, chave, match=None):
    if match:
        return match.group(chave) if match and match.group(chave) else "N/A"
    return item.get(chave, 'N/A')

def extrair_uasg_e_pregao(conteudo: str, padrao_1: str, padrao_srp: str, padrao_objeto: str) -> dict:
    match = re.search(padrao_1, conteudo)
    match2 = re.search(padrao_srp, conteudo)
    match3 = re.search(padrao_objeto, conteudo)
    
    srp_valor = match2.group("srp") if match2 else "N/A"
    objeto_valor = match3.group("objeto") if match3 else "N/A"

    if match:
        return {
            "uasg": match.group("uasg"),
            "orgao_responsavel": match.group("orgao_responsavel"),
            "num_pregao": match.group("num_pregao"),
            "ano_pregao": match.group("ano_pregao"),
            "srp": srp_valor,
            "objeto": objeto_valor
        }
    return {}

def buscar_itens(conteudo: str, padrao_grupo2: str, padrao_item2: str) -> list:
    return list(re.finditer(padrao_grupo2, conteudo)) + list(re.finditer(padrao_item2, conteudo))

def ajuste_cnpj(cnpj: str) -> str:
    cnpj = re.sub(r'\s+', '', cnpj)
    return cnpj

def processar_item(match, conteudo: str, ultima_posicao_processada: int, padrao_3: str, padrao_4: str) -> dict:
    item = match.groupdict()
    match_3 = re.search(padrao_3, conteudo[ultima_posicao_processada:])
    match_4 = re.search(padrao_4, conteudo[ultima_posicao_processada:])
    
    if match_3:
        ultima_posicao_processada += match_3.end()
    item_num_convertido = int(item.get('item_num')) if item.get('item_num', 'N/A') != 'N/A' else 'N/A'
    item_data = {
        "item_num": item_num_convertido,
        "grupo": encontre_valor_ou_NA(item, 'grupo'),
        "valor_estimado": encontre_valor_ou_NA(item, 'valor'),
        "quantidade": encontre_valor_ou_NA(item, 'quantidade'),
        "unidade": encontre_valor_ou_NA(item, 'unidade'),
        "situacao": encontre_valor_ou_NA(item, 'situacao'),
        "melhor_lance": encontre_valor_ou_NA(item, 'melhor_lance', match_3),
        "valor_negociado": encontre_valor_ou_NA(item, 'valor_negociado', match_3),
        "ordenador_despesa": encontre_valor_ou_NA(item, 'ordenador_despesa', match_3),
        "empresa": encontre_valor_ou_NA(item, 'empresa', match_3),
        "cnpj": ajuste_cnpj(encontre_valor_ou_NA(item, 'cnpj', match_3)),
        "marca_fabricante": encontre_valor_ou_NA(item, 'marca_fabricante', match_4),
        "modelo_versao": encontre_valor_ou_NA(item, 'modelo_versao', match_4),
    }
    return item_data, ultima_posicao_processada

def process_cnpj_data(cnpj_dict):
    """Converter "valor_estimado", "melhor_lance", e "valor_negociado" para float se não for possível deverá pular"""
    for field in ["valor_estimado", "melhor_lance", "valor_negociado"]:
        if isinstance(cnpj_dict[field], str):
            try:
                cnpj_dict[field] = float(cnpj_dict[field].replace(".", "").replace(",", "."))
            except ValueError:
                cnpj_dict[field] = 'N/A'

    # Convert "quantidade" to integer if possible, otherwise keep as is
    try:
        cnpj_dict["quantidade"] = int(cnpj_dict["quantidade"])
    except ValueError:
        pass

    # Ensure valor_homologado_item_unitario is defined
    if cnpj_dict["valor_negociado"] in [None, "N/A", "", "none", "null"]:
        cnpj_dict["valor_homologado_item_unitario"] = cnpj_dict["melhor_lance"]
    else:
        cnpj_dict["valor_homologado_item_unitario"] = cnpj_dict["valor_negociado"]

    # Now perform the other calculations
    if cnpj_dict["valor_estimado"] != 'N/A' and cnpj_dict["valor_homologado_item_unitario"] != 'N/A':
        try:
            cnpj_dict["valor_estimado_total_do_item"] = cnpj_dict["quantidade"] * float(cnpj_dict["valor_estimado"])
            cnpj_dict["valor_homologado_total_item"] = cnpj_dict["quantidade"] * float(cnpj_dict["valor_homologado_item_unitario"])
            cnpj_dict["percentual_desconto"] = (1 - (float(cnpj_dict["valor_homologado_item_unitario"]) / float(cnpj_dict["valor_estimado"]))) * 100
        except ValueError:
            pass
            
    return cnpj_dict

def identificar_itens_e_grupos(conteudo: str, padrao_grupo2: str, padrao_item2: str, padrao_3: str, padrao_4: str, df: pd.DataFrame) -> list:
    itens_data = []
    itens = buscar_itens(conteudo, padrao_grupo2, padrao_item2)
    ultima_posicao_processada = 0

    for match in itens:
        item_data, ultima_posicao_processada = processar_item(match, conteudo, ultima_posicao_processada, padrao_3, padrao_4)
        
        item_data = process_cnpj_data(item_data)

        itens_data.append(item_data)

    return itens_data

def create_dataframe_from_txt_files(txt_directory: str, padrao_1: str, padrao_grupo2: str, padrao_item2: str, padrao_3: str, padrao_4: str, df: pd.DataFrame):
    txt_files = obter_arquivos_txt(txt_directory)
    all_data = []
    
    for txt_file in txt_files:
        content = ler_arquivos_txt(txt_file)
        uasg_pregao_data = extrair_uasg_e_pregao(content, padrao_1, padrao_srp, padrao_objeto)
        items_data = identificar_itens_e_grupos(content, padrao_grupo2, padrao_item2, padrao_3, padrao_4, df)
        
        for item in items_data:
            all_data.append({
                **uasg_pregao_data,
                **item
            })

    df = pd.DataFrame(all_data)
    df = df.sort_values(by="item_num")

    return df

def salvar_txt_cnpj_empresa(df, output_path):
    # Drop duplicates by CNPJ column
    df_unique = df.drop_duplicates(subset='cnpj', keep='first')
    
    # Exclude rows where 'cnpj' or 'empresa' are 'N/A'
    df_filtered = df_unique[(df_unique['cnpj'] != 'N/A') & (df_unique['empresa'] != 'N/A')]
    
    # Sort the dataframe by CNPJ
    sicaf_df = df_filtered.sort_values(by='cnpj')
    
    # Save the CNPJ and Empresa data to a TXT file
    with open(output_path, 'w', encoding='utf-8') as f:
        for _, value in sicaf_df.iterrows():
            line = f"{value['cnpj']} - {value['empresa']}\n"
            f.write(line)

        # Writing the message about SICAF in the last line
        f.write(f'Salve o SICAF no diretório:\n{SICAF_DIR}.')

    return sicaf_df

TXT_OUTPUT_PATH = DATABASE_DIR / "relacao_cnpj.txt"

def save_to_dataframe(quadro_novo_ata, callback=None):
    df = pd.DataFrame()  # Initialize an empty dataframe
    df = create_dataframe_from_txt_files(str(TXT_DIR), padrao_1, padrao_grupo2, padrao_item2, padrao_3, padrao_4, df)
    
    # Lendo o arquivo tr_variavel.xlsx para atualização
    # Se o usuário carregou um DataFrame, use-o. Caso contrário, mostre uma mensagem de erro.
    if tr_variavel_df_carregado is not None:
        tr_variavel_df = tr_variavel_df_carregado
    else:
        messagebox.showerror("Erro", "Por favor, insira o arquivo corretamente.")
        return
    
    # Atualizando o DataFrame com as informações de tr_variavel.xlsx
    merged_df = pd.merge(df, tr_variavel_df, on='item_num', how='outer')

    # Salvar os DataFrames como arquivos
    EXCEL_OUTPUT_PATH = DATABASE_DIR / "relatorio.xlsx"
    merged_df.to_excel(EXCEL_OUTPUT_PATH, index=False, engine='openpyxl')
    merged_df.to_csv(CSV_OUTPUT_PATH, index=False, encoding='utf-8-sig')
    salvar_txt_cnpj_empresa(merged_df, TXT_OUTPUT_PATH)

    uasg = merged_df['uasg'].iloc[0]
    num_pregao = merged_df['num_pregao'].iloc[0]
    ano_pregao = "/" + str(merged_df['ano_pregao'].iloc[0])
    objeto = merged_df['objeto'].iloc[0][:45]
    qtde_adjudicados_homologados = len(merged_df[merged_df['situacao'] == "Adjudicado e Homologado"])
    qtde_desertos = len(merged_df[merged_df['situacao'] == "Deserto e Homologado"])
    qtde_fracassados = len(merged_df[merged_df['situacao'] == "Fracassado e Homologado"])
    valor_total_homologado = formatar_brl(merged_df['valor_homologado_total_item'].sum())
    qtde_itens_nao_concluidos = merged_df['situacao'].isna().sum()
    qtd_itens_concluidos = qtde_adjudicados_homologados + qtde_desertos + qtde_fracassados
    qtd_itens_licitados = len(merged_df['item_num'].unique())

    bold_font = tkFont.Font(family="Verdana", size=12, weight="bold")

    # Exibindo informações no quadro_novo_ata
    info_text = info_template.format(
        uasg, 
        num_pregao, 
        ano_pregao,
        objeto,
        qtd_itens_licitados,
        qtde_itens_nao_concluidos,
        qtd_itens_concluidos,
        qtde_adjudicados_homologados,
        qtde_desertos,
        qtde_fracassados,
        valor_total_homologado, 
            )

    label_info = tk.Label(quadro_novo_ata, text=info_text, font=bold_font, anchor='w', justify=tk.LEFT, bg="#FFFFFF")
    label_info.grid(row=0, column=0, sticky='w', pady=0, padx=0)
   
    DESERTO_FRACASSADO = gerar_relatorio_deserto_fracassado(merged_df)
    NAN_SITUACAO = gerar_relatorio_nan_situacao(merged_df)

    if callback:
        callback()
    return merged_df

def gerar_relatorio_deserto_fracassado(df):
    report = df[(df['situacao'] == "Deserto e Homologado") | (df['situacao'] == "Fracassado e Homologado")]
    DESERTO_FRACASSADO = DATABASE_DIR / "report_deserto_fracassado.xlsx"
    save_to_excel(report, DESERTO_FRACASSADO)
    return DESERTO_FRACASSADO

def gerar_relatorio_nan_situacao(df):
    report = df[df['situacao'].isna()]
    NAN_SITUACAO = DATABASE_DIR / "report_nan_situacao.xlsx"
    save_to_excel(report, NAN_SITUACAO)
    return NAN_SITUACAO

def open_excel_file():
    EXCEL_OUTPUT_PATH = DATABASE_DIR / "relatorio.xlsx"
    os.system(f'start excel "{EXCEL_OUTPUT_PATH}"')

def abrir_bloco_notas():
    TXT_OUTPUT_PATH = DATABASE_DIR / "relacao_cnpj.txt"
    os.startfile(TXT_OUTPUT_PATH)

def criar_pastas_com_subpastas() -> None:
    df = pd.read_csv(CSV_SICAF_PATH) 
    combinacoes = df[['num_pregao', 'ano_pregao', 'empresa']].drop_duplicates().values
    
    for num_pregao, ano_pregao, empresa in combinacoes:
        # Verifica se algum dos valores é NaN antes de prosseguir
        if pd.isna(num_pregao) or pd.isna(ano_pregao) or pd.isna(empresa):
            continue
        
        nome_dir_principal = (f"PE {num_pregao}-{ano_pregao}")
        path_dir_principal = BASE_DIR / nome_dir_principal
        
        if not path_dir_principal.exists():
            path_dir_principal.mkdir(parents=True)
        
        if empresa not in NOMES_INVALIDOS and empresa:
            nome_subpasta = (f"{empresa}")
            path_subpasta = path_dir_principal / nome_subpasta
        
            if not path_subpasta.exists():
                path_subpasta.mkdir()

from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Pt

def inserir_relacao_empresa(paragrafo, registro, cnpj):
    dados = {
        "Razão Social": registro["empresa"],
        "CNPJ": registro["cnpj"],
        "Endereço": registro["endereco"],
        "Município-UF": registro["municipio"],
        "CEP": registro["cep"],
        "Telefone": registro["telefone"],
        "E-mail": registro["email"]
    }
    
    for chave, valor in dados.items():
        adicione_texto_formatado(paragrafo, f"{chave}: ", True)
        adicione_texto_formatado(paragrafo, f"{valor};\n", False)
    
    adicione_texto_formatado(paragrafo, "Representada neste ato, por seu representante legal, o(a) Sr(a) ", False)
    adicione_texto_formatado(paragrafo, f'{registro["responsavel_legal"]}.\n', False)

def gerar_documento(NUMERO_ATA: int, df: pd.DataFrame) -> int:
    """Generate docx files based on template and CSV data, incrementing NUMERO_ATA for each doc."""
    combinacoes = df[['num_pregao', 'ano_pregao', 'empresa']].drop_duplicates().values
    
    for num_pregao, ano_pregao, empresa in combinacoes:
        nome_dir_principal = (f"PE {num_pregao}-{ano_pregao}")
        path_dir_principal = BASE_DIR / nome_dir_principal
        
        # Check for invalid names
        if empresa not in NOMES_INVALIDOS and empresa:
            nome_subpasta = (f"{empresa}")
            path_subpasta = path_dir_principal / nome_subpasta
            
            # Create subfolder if it doesn't exist
            if not path_subpasta.exists():
                path_subpasta.mkdir(parents=True)
            
            # Construct the header text
            texto_substituto = f"Nº 87000/2023-{NUMERO_ATA:03}/00\nPregão Eletrônico nº {num_pregao}/{ano_pregao}"
            
            # Find the relevant record for this document
            registro = df[df['empresa'] == empresa].iloc[0].to_dict()

            # Render the docx template with the data from the csv
            tpl = DocxTemplate(TEMPLATE_PATH)
            context = {
                "num_pregao": num_pregao,
                "ano_pregao": ano_pregao,
                "empresa": empresa,
                "numero_ata": NUMERO_ATA,
                "cabecalho": texto_substituto,
                "objeto": registro["objeto"],
                "ordenador_despesa": registro["ordenador_despesa"],
                "responsavel_legal": registro["responsavel_legal"]
            }
            tpl.render(context)
            
            # Save the rendered docx file
            nome_documento = (f"{empresa}.docx")
            path_documento = path_subpasta / nome_documento
            tpl.save(path_documento)
            
            # Increment NUMERO_ATA for the next document
            NUMERO_ATA += 1

    return NUMERO_ATA

def gerar_campos_item(item):
    return [
        (f'Item {item["item_num"]} - {item["descricao_tr"]} | Catálogo: {item["catalogo"]}', True),
        (f'Descrição: {item["descricao_detalhada_tr"]}', False),
        (f'Unidade de Fornecimento: {item["unidade"]}', False),
        (f'Marca/Fabricante: {item["marca_fabricante"]}   |   Modelo/Versão: {item["modelo_versao"]}', False),
        (f'Quantidade: {item["quantidade"]}   |   Valor Unitário: R$ {formatar_brl(item["valor_homologado_item_unitario"])}   |   Valor Total do Item: R$ {formatar_brl(item["valor_homologado_total_item"])}', False),
        (f'{"-" * 130}', False)
    ]

def inserir_relacao_itens(paragrafo, itens):
    # Primeiro, limpamos o parágrafo para remover o placeholder e qualquer outro texto.
    paragrafo.clear()

    for item in itens:
        campos = gerar_campos_item(item)
        for texto, negrito in campos:
            adicione_texto_formatado(paragrafo, texto + '\n', negrito)

    # Calculando o valor total homologado para a empresa
    valor_total = sum(float(item["valor_homologado_total_item"] or 0) for item in itens)  # tratando None como 0
    valor_extenso = valor_por_extenso(valor_total)
    
    adicione_texto_formatado(paragrafo, f'Valor total homologado para a empresa:\n', False)
    adicione_texto_formatado(paragrafo, f'R$ {formatar_brl(valor_total)} ({valor_extenso})\n', True)

def formatar_brl(valor):
    try:
        valor_float = float(valor)
    except (ValueError, TypeError):  # Se a conversão falhar ou o valor for None, usar 0.0
        valor_float = 0.0
    return locale.currency(valor_float, grouping=True, symbol=None)

def valor_por_extenso(valor):
    extenso = num2words(valor, lang='pt_BR', to='currency')
    return extenso.capitalize()

def alterar_documento_criado(caminho_documento, registro, cnpj, itens):
    # Carregando o documento real
    doc = Document(caminho_documento)
    
    # Iterando por cada parágrafo do documento
    for paragraph in doc.paragraphs:
        if '{relacao_empresa}' in paragraph.text:
            # Substituindo o marcador pelo conteúdo gerado pela função inserir_relacao_empresa
            paragraph.clear()  # Limpar o parágrafo atual
            inserir_relacao_empresa(paragraph, registro, cnpj)
        
        # Verificando o marcador {relacao_item}
        if '{relacao_item}' in paragraph.text:
            # Substituindo o marcador pelo conteúdo gerado pela função inserir_relacao_itens
            paragraph.clear()  # Limpar o parágrafo atual
            inserir_relacao_itens(paragraph, itens)
    
    # Salvando as alterações no documento
    doc.save(caminho_documento)

def processar_documentos(NUMERO_ATA: int):
    # 1. Load the CSV
    df = pd.read_csv(CSV_SICAF_PATH)
    
    # 2. Generate the documents using gerar_documento logic
    combinacoes = df[['uasg', 'num_pregao', 'ano_pregao', 'empresa']].drop_duplicates().values
    
    for uasg, num_pregao, ano_pregao, empresa in combinacoes:
        nome_dir_principal = f"PE {num_pregao}-{ano_pregao}"
        path_dir_principal = BASE_DIR / nome_dir_principal
        
        # Check for invalid names
        if empresa not in NOMES_INVALIDOS and empresa:
            nome_subpasta = f"{empresa}"
            path_subpasta = path_dir_principal / nome_subpasta
            
            # Create subfolder if it doesn't exist
            if not path_subpasta.exists():
                path_subpasta.mkdir()                   
            
            # Construct the header text
            texto_substituto = f"Nº {uasg}/2023-{NUMERO_ATA:03}/00\nPregão Eletrônico nº {num_pregao}/{ano_pregao}"
            
            # Find the relevant record for this document
            registro = df[df['empresa'] == empresa].iloc[0].to_dict()

            # Render the docx template with the data from the csv
            tpl = DocxTemplate(TEMPLATE_PATH)
            context = {
                "num_pregao": num_pregao,
                "ano_pregao": ano_pregao,
                "empresa": empresa,
                "uasg": uasg,
                "numero_ata": NUMERO_ATA,
                "cabecalho": texto_substituto,
                "objeto": registro["objeto"],
                "ordenador_despesa": registro["ordenador_despesa"],
                "responsavel_legal": registro["responsavel_legal"]
            }
            tpl.render(context)
            
            # Save the rendered docx file
            nome_documento = f"{empresa}.docx"
            path_documento = path_subpasta / nome_documento
            try:
                path_documento.parent.mkdir(parents=True, exist_ok=True)
                tpl.save(path_documento)
            except FileNotFoundError as e:
                print(f"Erro ao salvar o documento: {e}")
            
            tpl.save(path_documento)

            # Update or add the num_ata column
            if 'num_ata' not in df.columns:
                df['num_ata'] = ""

            df['num_ata'] = df['num_ata'].astype(str)
            df.loc[df['empresa'] == empresa, 'num_ata'] = f"{uasg}/2023-{NUMERO_ATA:03}/00"
            NUMERO_ATA += 1
            
    # 3. Alter the generated documents using alterar_documento_criado function
    for uasg, num_pregao, ano_pregao, empresa in combinacoes:
        if empresa not in NOMES_INVALIDOS and empresa:
            nome_dir_principal = f"PE {num_pregao}-{ano_pregao}"
            path_dir_principal = BASE_DIR / nome_dir_principal
            nome_subpasta = f"{empresa}"
            path_subpasta = path_dir_principal / nome_subpasta
            
            # Construct the path to the document
            nome_documento = f"{empresa}.docx"
            path_documento = path_subpasta / nome_documento
            
            # Find the relevant record for this document
            registro = df[df['empresa'] == empresa].iloc[0].to_dict()
            itens_relacionados = df[df['empresa'] == empresa].to_dict('records')

            # Modify the document
            alterar_documento_criado(path_documento, registro, registro["cnpj"], itens_relacionados)
    
    # 4. Save the updated df to a new CSV file
    csv_filename = f"PE {num_pregao}-{ano_pregao}.csv"
    df.to_csv(csv_filename, index=False)
    
    # 5. Save the updated df to a new Excel file
    excel_filename = f"PE {num_pregao}-{ano_pregao}.xlsx"
    df.to_excel(excel_filename, index=False)

def selecionar_termo_de_referencia_e_carregar(quadro):
    global tr_variavel_df_carregado

    # Restringindo a seleção de arquivos para .xlsx diretamente no diálogo
    arquivo = filedialog.askopenfilename(initialdir=BASE_DIR, 
                                     filetypes=[("Spreadsheet files", ("*.xlsx", "*.xls", "*.ods")), 
                                                ("Excel files", ("*.xlsx", "*.xls")), 
                                                ("LibreOffice files", "*.ods")])

    if arquivo:
        tr_variavel_df_carregado = pd.read_excel(arquivo)

def iniciar_processo():
    if NUMERO_ATA_GLOBAL is None:
        raise ValueError("NUMERO_ATA not set!")
    
    # 1. Create the necessary directories and subdirectories
    criar_pastas_com_subpastas()
    
    # 2. Generate and modify the documents
    processar_documentos(NUMERO_ATA_GLOBAL)

def criar_passos(quadro, numero_ata, gerador_numero_ata, quadro_novo_ata):
    adicionar_rotulo(quadro, "Passo 1: Importar Termo de Referência", 0, fonte="Verdana", tamanho="14", bold=True)
    imagem_import_tr = criar_label_com_imagem(quadro, ICONS_DIR / "import.png", row=0, tamanho=(2,2), column=0, padx=(440, 0), sticky=tk.W, comando=lambda: selecionar_termo_de_referencia_e_carregar(quadro))
    
    adicionar_traco_preto(quadro, 1, 710)

    #PASSO 01
    adicionar_rotulo(quadro, "Passo 2: Processar Termos de Homologação", 2, fonte="Verdana", tamanho="14", bold=True)
    adicionar_rotulo(quadro, "Insira os arquivos PDF na pasta:", 3, fonte="Verdana", tamanho="12", bold=False)
    adicionar_rotulo(quadro, PDF_DIR, 4, fonte="Verdana", tamanho="12", bold=False, padx_value=40)

    label_pasta = criar_label_com_imagem(quadro, ICONS_DIR / "pasta_novo.png", row=4, tamanho=(2,2), column=0, padx=(0, 0), sticky=tk.W)
    label_pasta.bind("<Button-1>", lambda event: ao_clicar_no_label(event, PDF_DIR))

    confirm_homolog = criar_label_com_imagem(quadro, ICONS_DIR / "check-mark2.png", row=5, tamanho=(2,2), column=0, padx=(0, 0), sticky=tk.W)
    confirm_homolog.bind("<Button-1>", lambda event: (efeito_visual_apos_clique(event), converter_pdf_e_salvar_dataframe(quadro, PDF_DIR, TXT_DIR, progress_bar, progress_var, quadro_novo_ata)))

    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(quadro, variable=progress_var, maximum=100, length=300)
    progress_bar.grid(row=5, column=0, padx=(35, 0), sticky=tk.W)

    def converter_pdf_e_salvar_dataframe(quadro, PDF_DIR, TXT_DIR, progress_bar, progress_var, quadro_novo_ata):
        convert_pdf_to_txt(quadro, PDF_DIR, TXT_DIR, progress_bar, progress_var)
        df = save_to_dataframe(quadro_novo_ata)
        exibir_dataframe_itens_homologados(quadro_novo_ata, df)

    botao_converter = tk.Button(quadro, text="Teste Rápido", command=lambda: teste_1(quadro_novo_ata))
    botao_converter.grid(row=0, column=0, padx=(490, 0), sticky=tk.W)

    def teste_1(quadro_novo_ata):
        df = save_to_dataframe(quadro_novo_ata)
        exibir_dataframe_itens_homologados(quadro_novo_ata, df)

    adicionar_traco_preto(quadro, 6, 710)
    #PASSO 02
    adicionar_rotulo(quadro, "Passo 3: Processar Dados SICAF", 7, fonte="Verdana", tamanho="14", bold=True)
 
    label_txt = criar_label_com_imagem(quadro, ICONS_DIR / "txt.png", row=8, tamanho=(2,2), column=0, padx=(0, 0), sticky=tk.W)
    label_txt.bind("<Button-1>", lambda event: (efeito_visual_apos_clique(event), abrir_bloco_notas()))

    adicionar_rotulo(quadro, "Relação de Empresas Homologadas", 8, fonte="Verdana", tamanho="12", bold=False, padx_value=40)
    adicionar_rotulo(quadro, "Insira os arquivos SICAF-PDF na pasta:", 9, fonte="Verdana", tamanho="12", bold=False)

    label_pasta_sicaf = criar_label_com_imagem(quadro, ICONS_DIR / "pasta_novo.png", row=10, tamanho=(2,2), column=0, padx=(0, 0), sticky=tk.W)
    label_pasta_sicaf.bind("<Button-1>", lambda event: ao_clicar_no_label(event, SICAF_DIR))

    adicionar_rotulo(quadro, SICAF_DIR, 10, fonte="Verdana", tamanho="12", bold=False, padx_value=40)
    
    confirm_sicaf = criar_label_com_imagem(quadro, ICONS_DIR / "check-mark2.png", row=11, tamanho=(2,2), column=0, padx=(0, 0), sticky=tk.W)
    confirm_sicaf.bind("<Button-1>", lambda event: (efeito_visual_apos_clique(event), processas_sicaf_e_salvar_dataframe(quadro, progress_bar_sicaf, progress_var_sicaf, quadro_novo_ata)))

    progress_var_sicaf = tk.DoubleVar()
    progress_bar_sicaf = ttk.Progressbar(quadro, variable=progress_var_sicaf, maximum=100, length=300)
    progress_bar_sicaf.grid(row=11, column=0, padx=(35, 0), sticky=tk.W)

    def processas_sicaf_e_salvar_dataframe(quadro, progress_bar_sicaf, progress_var_sicaf, quadro_novo_ata):
        df_final_ordered = processar_arquivos_sicaf(quadro, progress_bar_sicaf, progress_var_sicaf)
        salvar_txt_cnpj_empresa(df_final_ordered, TXT_OUTPUT_PATH)
        exibir_dataframe_sicaf(quadro_novo_ata, df_final_ordered)
   
    adicionar_traco_preto(quadro, 12, 710)
    #PASSO 03
    adicionar_rotulo(quadro, "Passo 4: Gerar Ata de Registro de Preços", 13, fonte="Verdana", tamanho="14", bold=True)
    adicionar_rotulo(quadro, "Insira o próximo número sequencial:", 14, fonte="Verdana", tamanho="12", bold=False)

    label_pasta_sicaf = criar_label_com_imagem(quadro, ICONS_DIR / "pasta_novo.png", row=15, tamanho=(2,2), column=0, padx=(0, 0), sticky=tk.W)
    label_pasta_sicaf.bind("<Button-1>", lambda event: ao_clicar_no_label(event, RELATORIO_PATH))

    adicionar_rotulo(quadro, ATA_DIR, 15, fonte="Verdana", tamanho="12", bold=False, padx_value=40)

    # Campo de entrada para o número da Ata
    fonte_tamanho_14 = font.Font(size=14)
    numero_ata = tk.Entry(quadro, font=fonte_tamanho_14, width=4, bg="#EFE4B0")  # #FFFFE0 é a cor hexadecimal para amarelo claro
    numero_ata.grid(row=14, column=0, padx=(310, 0), sticky=tk.W)

    confirm_sicaf = criar_label_com_imagem(quadro, ICONS_DIR / "check-mark2.png", row=14, tamanho=(2,2), column=0, padx=(355, 0), sticky=tk.W)
    confirm_sicaf.bind("<Button-1>", lambda event: (efeito_visual_apos_clique(event), confirmar_numero_ata(numero_ata.get())))
    
    adicionar_traco_preto(quadro, 16, 710)
   
    label_gerar_processo = criar_label_com_imagem(quadro, ICONS_DIR / "docx_64.png", texto=" Gerar Ata ", row=18, tamanho=(2,2), column=0, padx=(75, 0), sticky=tk.W, fonte="Verdana", tamanho_fonte=22, bold=True, espessura_borda=1, cor_borda="#000000", cor_bg="#b1ef8f")
    label_gerar_processo.bind("<Button-1>", lambda event: (efeito_visual_apos_clique(event), iniciar_processo()))
    
    label_relatorio = criar_label_com_imagem(quadro, ICONS_DIR / "xlsx.png", texto=" Relatorio ", row=18, tamanho=(2,2), column=0, padx=(300, 0), sticky=tk.W, fonte="Verdana", tamanho_fonte=22, bold=True, espessura_borda=1, cor_borda="#000000", cor_bg="#b1ef8f")
    label_relatorio.bind("<Button-1>", lambda event: (efeito_visual_apos_clique(event), open_excel_file()))

    # label_relatorio = criar_botao_com_imagem(quadro, ICONS_DIR / "xlsx.png", texto=" Relatorio ", row=19, tamanho=(2,2), column=0, padx=(300, 0), sticky=tk.W, fonte="Verdana", tamanho_fonte=22, bold=True, espessura_borda=1, cor_borda="#000000", cor_bg="#FFFFFF")
    # label_relatorio.bind("<Button-1>", lambda event: (efeito_visual_apos_clique(event), open_excel_file()))

# def criar_botao_com_imagem(quadro, caminho_imagem, texto="", tamanho=(32, 32), fonte="Verdana", tamanho_fonte=10, bold=False, comando=None, espessura_borda=0, cor_borda="#000000", cor_bg="#FFFFFF", **opcoes_grid):
#     img = carregar_imagem_icone(caminho_imagem, tamanho)
#     estilo_fonte = "bold" if bold else "normal"
#     fonte_config = (fonte, tamanho_fonte, estilo_fonte)
    
#     # Criando um estilo personalizado para o botão
#     style = ttk.Style()
#     style.configure("Custom.TButton", image=img, background=cor_bg, font=fonte_config, borderwidth=espessura_borda)
#     style.map("Custom.TButton",
#               background=[('active', 'black'), ('pressed', 'black')],
#               foreground=[('active', '#FFFFFF'), ('pressed', '#FFFFFF')],
#               bordercolor=[('active', 'white'), ('pressed', 'white')],
#               borderwidth=[('active', 2), ('pressed', 2)])
    
#     botao = ttk.Button(quadro, text=texto, style="Custom.TButton", command=comando)
#     botao.grid(**opcoes_grid)
    
#     return botao

    # botao_relatorio = BotaoPersonalizado_interno(
    #     master=quadro,
    #     text="Relatório",
    #     command=lambda: open_excel_file(),
    #     x=250,
    #     y=18,  # Ajuste esse valor conforme necessário
    #     width=200,  # Ajuste esse valor conforme necessário
    #     height=50, # Ajuste esse valor conforme necessário
    #     caminho_imagem=ICONS_DIR / "xlsx.png"
    # )

def seu_gerador_inicial(valor_inicial: int):
    """Gerador que fornece números de ata incrementais a partir de um número inicial."""
    numero = valor_inicial
    while True:
        valor_recebido = (yield numero)
        if valor_recebido is not None:
            numero = valor_recebido
        else:
            numero += 1

def confirmar_numero_ata(numero_ata):
    global NUMERO_ATA_GLOBAL, GERADOR_NUMERO_ATA
    
    # Update the global variable with the user's input
    NUMERO_ATA_GLOBAL = int(numero_ata)
    
    # Display the confirmation message
    mensagem = f'A próxima ata de registro de preços será "87000/2023-{numero_ata}-00"'
    # Assuming messagebox is already imported; otherwise, you'll need to import it.
    messagebox.showinfo("Confirmação", mensagem)
    
    # Update the generator
    GERADOR_NUMERO_ATA = seu_gerador_inicial(NUMERO_ATA_GLOBAL)
    next(GERADOR_NUMERO_ATA)

def situacao_modificada(situacao_original):
    if situacao_original == "Adjudicado e Homologado":
        return "Item Homologado", 'green'
    elif situacao_original == "Deserto e Homologado":
        return "Item Deserto", 'orange'
    elif situacao_original == "Fracassado e Homologado":
        return "Item Fracassado", 'orange'
    else:
        return "Item não encontrado!", 'red'

def exibir_dataframe_itens_homologados(quadro_novo_ata, df=None):
    global dataframe_licitacao
    
    if df is None:
        return

    # Crie um objeto Style
    style = ttk.Style()
    style.configure('Treeview', font=('Verdana', 13))
    style.configure('Treeview.Heading', font=('Verdana', 13))

    # Limpar o Treeview existente
    for i in dataframe_licitacao.get_children():
        dataframe_licitacao.delete(i)

    # Configurando colunas e cabeçalhos (se isso não mudar frequentemente, você pode considerar movê-lo para fora desta função)
    dataframe_licitacao["columns"] = ("item_num", "descricao_tr", "situacao")
    dataframe_licitacao.column("#0", width=0, stretch=tk.NO)  # Remova o espaço reservado para a imagem

    for col in dataframe_licitacao["columns"]:
        display_name = nomes_das_colunas.get(col, col)
        dataframe_licitacao.heading(col, text=display_name)
        dataframe_licitacao.column(col, width=ajustar_colunas_especificas.get(col, 240), stretch=tk.NO)
        
    # Centralizando o conteúdo das colunas "simbolo" e "item_num"
    dataframe_licitacao.column("item_num", anchor="center")

    for _, row in df.iterrows():
        situacao, cor_fundo = situacao_modificada(row['situacao'])
        dataframe_licitacao.insert("", "end", values=(row['item_num'], row['descricao_tr'], situacao), tags=(cor_fundo,))
        dataframe_licitacao.tag_configure(cor_fundo, background=cor_fundo)
    
    # Adicionando a barra de rolagem
    scrollbar = ttk.Scrollbar(quadro_novo_ata, orient="vertical", command=dataframe_licitacao.yview)
    scrollbar.grid(row=4, column=0, padx=(500, 0), sticky='ns')  # Posiciona a barra de rolagem à direita do Treeview

    dataframe_licitacao.configure(yscrollcommand=scrollbar.set)
    dataframe_licitacao.configure(height=10)
    dataframe_licitacao.grid(row=4, column=0, sticky='nsew')  # Posiciona o Treeview na primeira coluna

import pandas as pd

def exibir_dataframe_sicaf(quadro_novo_ata, sicaf_df):
    global dataframe_sicaf
    
    # Carregar o CSV sicaf.csv
    sicaf_output_df = pd.read_csv(CSV_SICAF_PATH)

    # Conjunto de CNPJs únicos de sicaf_output_df
    cnpj_sicaf_output = set(sicaf_output_df['cnpj'].unique())

    # Filtrar sicaf_df para remover NaN e manter apenas registros únicos
    sicaf_df = sicaf_df.dropna(subset=["cnpj", "empresa"]).drop_duplicates(subset="cnpj")

    # Crie um objeto Style
    style = ttk.Style()
    style.configure('Treeview', font=('Verdana', 13))
    style.configure('Treeview.Heading', font=('Verdana', 13))

    # Limpar o Treeview existente
    for i in dataframe_sicaf.get_children():
        dataframe_sicaf.delete(i)

    # Configurando colunas e cabeçalhos
    dataframe_sicaf["columns"] = ("cnpj", "empresa")
    dataframe_sicaf.column("#0", width=0, stretch=tk.NO)

    for col in dataframe_sicaf["columns"]:
        display_name = nomes_das_colunas.get(col, col)
        dataframe_sicaf.heading(col, text=display_name)
        dataframe_sicaf.column(col, width=ajustar_colunas_especificas.get(col, 10), stretch=tk.NO)

    # Inserir dados de sicaf_df no Treeview
    for _, row in sicaf_df.iterrows():
        if row['cnpj'] in cnpj_sicaf_output:
            situacao_fornecedor = sicaf_output_df[sicaf_output_df['cnpj'] == row['cnpj']]['situação_fornecedor'].values[0]
            if pd.isna(situacao_fornecedor) or situacao_fornecedor == '':
                dataframe_sicaf.insert("", "end", values=(row["cnpj"], row["empresa"]), tags='unmatched')
            else:
                dataframe_sicaf.insert("", "end", values=(row["cnpj"], row["empresa"]), tags='matched')
        else:
            dataframe_sicaf.insert("", "end", values=(row["cnpj"], row["empresa"]), tags='unmatched')

    # Colorindo linhas com base nas tags
    dataframe_sicaf.tag_configure('matched', background='green')
    dataframe_sicaf.tag_configure('unmatched', background='red')

    # Adicionando uma barra de rolagem
    scrollbar = ttk.Scrollbar(quadro_novo_ata, orient="vertical", command=dataframe_sicaf.yview)

    # Posicione o Treeview e a Scrollbar usando grid
    scrollbar.grid(row=7, column=0, padx=(500, 0), sticky='ns')

    dataframe_sicaf.configure(yscrollcommand=scrollbar.set)
