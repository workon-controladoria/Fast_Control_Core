from pathlib import Path
import os
from pydantic_settings import BaseSettings, SettingsConfigDict

# ========================
# CONFIGURAÇÕES DE CAMINHOS
# ========================

# Caminho absoluto para a pasta onde este arquivo (settings.py) está
_CURRENT_FILE = Path(__file__).resolve()

# Raiz do Projeto (volta 3 níveis: config -> core -> controladoria_cli -> RAIZ)
BASE_DIR = _CURRENT_FILE.parent.parent.parent

# Definição das pastas de dados relativas à raiz
DATA_DIR = BASE_DIR / "dados_aps"
OUTPUT_DIR = BASE_DIR / "output_dados_aps"

# Garante que as pastas existam
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Caminho padrão da base (sem ser fixo no seu usuário)
CAMINHO_BASE_PADRAO = DATA_DIR / "data_base_de_aps.xlsx"


# ==================================
# CONFIGURAÇÕES DE BANCO DE DADOS
# ==================================

class Settings(BaseSettings):
    """
    Carrega e valida as configurações do ambiente a partir de variáveis de ambiente
    ou de um arquivo .env.
    """

    DB_DRIVER: str = "SQL Server"
    DB_SERVER: str
    DB_NAME: str
    DB_USER: str
    DB_PASSWORD: str

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

settings = Settings()
