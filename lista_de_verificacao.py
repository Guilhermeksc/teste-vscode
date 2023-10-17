import PyPDF2
from docxtpl import DocxTemplate
import tkinter as tk
from tkinter import filedialog
from diretorios import *

TEMPLATE_CHECKLIST = BASE_DIR / 'template/checklist.docx'
LV_DIR = BASE_DIR / "Lista_de_Verificacao"

def selecionar_pdf():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    return file_path

def buscar_marcadores_no_pdf(pdf_path):
    marcadores_pdf = {
        "autorização_para_abertura_": [],
        "dfd_inicio_": [],
        "dfd_fim_": [],
        "etp_inicio_": [],
        "etp_fim_": [],
        "mr_inicio_": [],
        "mr_fim_": [],
        "dec_adeq_orc_": [],
        "port_plan_inicio_": [],
        "port_plan_fim_": [],
        "tr_inicio_": [],
        "tr_fim_": [],
        "edital_inicio_": [],
        "edital_fim_": [],
    }

    with open(pdf_path, "rb") as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        for page_number in range(len(reader.pages)):
            page = reader.pages[page_number]
            page_text = page.extract_text()
            for marcador in marcadores_pdf:
                if marcador in page_text:
                    marcadores_pdf[marcador].append(page_number + 1)
    return marcadores_pdf

def substituir_marcadores_no_docx(docx_path, marcadores):
    doc = DocxTemplate(docx_path)
    context = {}
    for marcador, paginas in marcadores.items():
        if paginas:
            context[marcador] = str(paginas[0])
        else:
            context[marcador] = ""  # ou algum valor padrão se não for encontrado
    doc.render(context)
    output_path = LV_DIR / "checklist_modificado.docx"
    doc.save(output_path)

def main():
    pdf_path = selecionar_pdf()
    if not pdf_path:
        return
    marcadores = buscar_marcadores_no_pdf(pdf_path)
    substituir_marcadores_no_docx(TEMPLATE_CHECKLIST, marcadores)

if __name__ == "__main__":
    main()