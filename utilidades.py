from diretorios import *
import pdfplumber
from tkinter import messagebox
from pathlib import Path
import os
import tkinter as tk

def convert_pdf_to_txt(frame, pdf_dir, txt_dir, progress_bar, progress_var):    
    # Verifica se TXT_DIR existe. Se não, cria o diretório.
    if not txt_dir.exists():
        txt_dir.mkdir(parents=True, exist_ok=True)
    else:
        # Se TXT_DIR existir, deleta todos os arquivos dentro dele.
        for file in txt_dir.iterdir():
            if file.is_file():
                file.unlink()
    
    # Inicia o processo de conversão
    pdf_files = list(pdf_dir.glob("*.pdf"))
    total_files = len(pdf_files)
    
    for index, pdf_file in enumerate(pdf_files):
        with pdfplumber.open(pdf_file) as pdf:
            texts = [page.extract_text() for page in pdf.pages]
            all_text = ' '.join(texts).replace('\n', ' ').replace('\x0c', ' ')

            txt_file = txt_dir / f"{pdf_file.stem}.txt"
            with open(txt_file, 'w', encoding='utf-8') as f:
                f.write(all_text)
        
        progress_var.set((index + 1) / total_files * 100)
        frame.update_idletasks()
    
    messagebox.showinfo("Conversão Completa", "Todos os arquivos PDF foram convertidos para TXT!")

def obter_arquivos_txt(directory: str) -> list:
    """Retorna a lista de arquivos TXT em um diretório."""
    return [os.path.join(directory, file) for file in os.listdir(directory) if file.endswith('.txt')]

def ler_arquivos_txt(file_path: str) -> str:
    """Lê o conteúdo de um arquivo TXT."""
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()

import locale
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def formatar_brl(valor):
    try:
        valor_float = float(valor)
    except (ValueError, TypeError):  # Se a conversão falhar ou o valor for None, usar 0.0
        valor_float = 0.0
    return locale.currency(valor_float, grouping=True, symbol=None)

def save_to_excel(df, filepath):
    df.to_excel(filepath, index=False, engine='openpyxl')