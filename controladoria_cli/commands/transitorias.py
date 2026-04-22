from colorama import Fore, Style
from datetime import datetime

from ..core.services import transitorias_services

def run_extracao_transitarias_interativo():
    print(Fore.YELLOW + "--- Módulo de Transitórias: Extração de APs ---")
    valores_aps_str = input(Fore.BLUE + "Digite os valores da coluna 'APS' para filtrar, separados por vírgula: " + Style.RESET_ALL)
    if not valores_aps_str:
        print(Fore.RED + "Nenhum valor de APS fornecido. Abortando.")
        return
    valores_aps = [valor.strip() for valor in valores_aps_str.split(',')]
    df_resultado, nao_encontrados = transitorias_services.processar_extracao_aps_transitarias(valores_aps)
    if nao_encontrados:
        print(Fore.RED + f"\nAVISO: As seguintes APS não foram encontradas: {', '.join(nao_encontrados)}")
    if df_resultado is None or df_resultado.empty:
        print(Fore.YELLOW + "Nenhum dado correspondente encontrado para as APS fornecidas.")
        return
    print(Fore.GREEN + f"\n{len(df_resultado)} registros encontrados.")
    print(Fore.YELLOW + "\nPré-visualização dos Dados Finais:")
    colunas_para_mostrar = ['APS', 'NATUREZA', 'CUSTO OU DESPESA', 'conta contabil']
    colunas_existentes = [col for col in colunas_para_mostrar if col in df_resultado.columns]
    print(Fore.CYAN + str(df_resultado[colunas_existentes].head(10)))
    salvar = input(Fore.BLUE + "\nDeseja salvar o resultado em um arquivo Excel? (s/n): " + Style.RESET_ALL).strip().lower()
    if salvar == 's':
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f'extracao_transitarias_{timestamp}.xlsx'
        filename = input(Fore.BLUE + f"Digite o nome do arquivo (padrão: {default_filename}): " + Style.RESET_ALL).strip()
        if not filename:
            filename = default_filename
        transitorias_services.salvar_dataframe_transitarias(df_resultado, filename)