import os
from colorama import Fore, Style
from datetime import datetime

from ..core.services import provisoes_service

def run_provisoes_interativo():
    print(Fore.YELLOW + "\n--- Módulo de Provisões ---" + Style.RESET_ALL)
    print("Este módulo transforma a base horizontal no layout vertical do ERP (Provisão e Reversão).")
    
    try:
        ano = int(input(Fore.BLUE + "Ano da competência (ex: 2026): " + Style.RESET_ALL))
        mes = int(input(Fore.BLUE + "Mês da competência (1-12): " + Style.RESET_ALL))
    except ValueError:
        print(Fore.RED + "❌ Entrada inválida. Digite apenas números.")
        return

    caminho_planilha = input(Fore.BLUE + "Arraste para cá a planilha base de provisões: " + Style.RESET_ALL).strip()
    
    # Limpa aspas do Windows
    if caminho_planilha.startswith('"') and caminho_planilha.endswith('"'):
        caminho_planilha = caminho_planilha[1:-1]
        
    if not os.path.isfile(caminho_planilha):
        print(Fore.RED + "❌ Arquivo não encontrado!")
        return

    df_layout = provisoes_service.processar_planilha_provisoes(caminho_planilha, ano, mes)
    
    if df_layout is not None and not df_layout.empty:
        qtd_linhas = len(df_layout)
        print(Fore.GREEN + f"\n✅ Sucesso! {qtd_linhas} linhas geradas para o ERP.")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"Provisoes_Layout_ERP_{ano}_{mes:02d}_{timestamp}.xlsx"
        provisoes_service.salvar_layout_erp_por_empresa(df_layout, nome_arquivo)
    else:
        print(Fore.YELLOW + "⚠️ Nenhuma linha gerada. Verifique se o arquivo tem as colunas corretas (EMPRESA, CLIENTE, C.C., etc).")