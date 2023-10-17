from pathlib import Path
from tkinter import filedialog
import numpy as np
import json

BASE_DIR = Path(__file__).resolve().parent
DATABASE_DIR = BASE_DIR / "database"
CONFIG_FILE = BASE_DIR / "config.json"

def load_config(key, default_value):
    try:
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
            return config.get(key, default_value)
    except (FileNotFoundError, json.JSONDecodeError):
        return default_value

def save_config(key, value):
    config = {}
    try:
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        pass
    config[key] = value
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

def update_dir(title, key, default_value):
    new_dir = filedialog.askdirectory(title=title)
    if new_dir:
        save_config(key, new_dir)
        return Path(new_dir)
    return default_value

ICONS_DIR = DATABASE_DIR / "icons"
DATABASE_DIR = Path(load_config("DATABASE_DIR", BASE_DIR / "database"))
PDF_DIR = Path(load_config("PDF_DIR", DATABASE_DIR / "pasta_pdf"))
SICAF_DIR = Path(load_config("SICAF_DIR", DATABASE_DIR / "pasta_sicaf"))
PASTA_TEMPLATE = Path(load_config("PASTA_TEMPLATE", BASE_DIR / "template"))
RELATORIO_PATH = Path(load_config("RELATORIO_PATH", DATABASE_DIR / "relatorio"))
TXT_DIR = PDF_DIR / "homolog_txt"
SICAF_TXT_DIR = SICAF_DIR / "sicaf_txt"
ATA_DIR = DATABASE_DIR / "atas"
TR_VAR_DIR = DATABASE_DIR / "tr_variavel.xlsx"
NOMES_INVALIDOS = ['N/A', None, 'None', 'nan', 'null', np.nan]
TEMPLATE_PATH = BASE_DIR / 'database/template_ata.docx'
TEMPLATE_PATH_TEMP = BASE_DIR / 'database/template_ata_temp.docx'
CSV_DIR = DATABASE_DIR / "dados.csv"
VARIAVEIS_DIR = DATABASE_DIR / "variaveis.xlsx"
XLSX_SICAF_PATH = DATABASE_DIR / "sicaf.xlsx"
CSV_SICAF_PATH = DATABASE_DIR / "sicaf.csv"