import pandas as pd
from pathlib import Path
from datetime import datetime

from ..core.config import settings
from ..core.services.base_dados_service import BaseDadosService

# --- CAMINHO DE SAÍDA PARA VOCÊ PREENCHER ---
# Coloque o caminho completo para a sua pasta de destino dentro das aspas.
# Exemplo: 'C:/Users/seu_usuario/Desktop/Relatorios_Fast_Control'
CAMINHO_PASTA_SAIDA = 'C:/Users/jose.santos/Desktop/CONTROLADORIA/Automacoes/Fast_Control/controladoria_cli/dados_aps/output_dados_base_aps'


def solicitar_data(tipo: str) -> tuple[int, int]:
    """Solicita e valida o ano e mês do usuário."""
    while True:
        try:
            ano = int(input(f"Insira o Ano de {tipo}: "))
            mes = int(input(f"Insira o Mês de {tipo} (1-12): "))
            if 1 <= mes <= 12:
                return ano, mes
            else:
                print("[AVISO] Mês inválido. Por favor, insira um valor entre 1 e 12.")
        except ValueError:
            print("[AVISO] Entrada inválida. Por favor, insira apenas números.")

def run_extracao_aps_interativo():
    """
    Função que guia o usuário para extrair e processar a base de dados de APs.
    """
    print("\n--- Extração da Base de APs ---")
    
    ano_inicio, mes_inicio = solicitar_data("início")
    ano_fim, mes_fim = solicitar_data("fim")

    start_date = f"{ano_inicio}{str(mes_inicio).zfill(2)}"
    end_date = f"{ano_fim}{str(mes_fim).zfill(2)}"

    print(f"\nIniciando extração de APs para o período de {start_date} a {end_date}...")

    service = BaseDadosService(db_settings=settings)
    df_final = service.get_and_process_aps(start_date=start_date, end_date=end_date)

    if df_final.empty:
        print("\n[AVISO] Nenhum dado encontrado para o período especificado.")
        input("Pressione Enter para continuar...")
        return

    if not CAMINHO_PASTA_SAIDA or not CAMINHO_PASTA_SAIDA.strip():
        print("\n[ERRO CRÍTICO] O caminho de saída não foi definido no código.")
        print("Por favor, edite o script e preencha a variável 'CAMINHO_PASTA_SAIDA'.")
        input("Pressione Enter para continuar...")
        return

    output_dir = Path(CAMINHO_PASTA_SAIDA)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"base_aps_{timestamp}.xlsx"

    print(f"\nSalvando {len(df_final)} registros em '{output_path}'...")
    df_final.to_excel(output_path, index=False, engine='xlsxwriter')
    
    print("\n[SUCESSO] Arquivo salvo com sucesso!")
    input("Pressione Enter para continuar...")