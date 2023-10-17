import subprocess
from diretorios import *

def build_executable():
    # Caminho para o script principal
    main_script = "menu.py"
    
    # Recursos para adicionar
    resources = [
        (DATABASE_DIR, "database"),
        (ICONS_DIR, "icons"),
        (PDF_DIR, "pasta_pdf"),
        (SICAF_DIR, "pasta_sicaf"),
        (PASTA_TEMPLATE, "template"),
        (SICAF_TXT_DIR, "sicaf_txt"),
        (TXT_DIR, "homolog_txt"),
        (RELATORIO_PATH, "database"),
        (TEMPLATE_PATH, "template_ata.docx"),
        (TEMPLATE_PATH_TEMP, "template_ata_temp.docx"),
    ]
    # Argumentos para o PyInstaller
    pyinstaller_args = [
        "pyinstaller",  # Isso pressupõe que pyinstaller esteja em seu PATH.
    ]
    
    # Adicionando recursos
    for src, dest in resources:
        pyinstaller_args.append(f"--add-data={src};{dest}")
    
    # Adicionando o script principal
    pyinstaller_args.append(main_script)
    
    # Chamar o PyInstaller
    subprocess.run(pyinstaller_args)

if __name__ == "__main__":
    build_executable()
