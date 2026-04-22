import os
from pathlib import Path
from colorama import Fore, Style
from datetime import datetime

from ..core.services import transitorias_services
from ..core.config import settings
from ..core.config.settings import CAMINHO_BASE_PADRAO

def run_extracao_transitarias_interativo():
    print(Fore.YELLOW + "\n--- Módulo de Transitórias ---" + Style.RESET_ALL)
    print("1. Extração Padrão (Digitar números das APs)")
    print("2. Conciliação Fiscal (Via planilha RAZAO_NORTE)")
    print("0. Voltar")
    
    escolha = input(Fore.BLUE + "Escolha uma opção: " + Style.RESET_ALL).strip()
    
    if escolha == '1':
        _executar_extracao_padrao()
    elif escolha == '2':
        _executar_conciliacao_fiscal()
    elif escolha == '0':
        return
    else:
        print(Fore.RED + "Opção inválida!" + Style.RESET_ALL)

def _executar_extracao_padrao():
    """Mantém a sua funcionalidade original intacta."""
    print(Fore.CYAN + "\n[ Extração Padrão ]" + Style.RESET_ALL)
    aps_input = input(Fore.BLUE + "Digite os números das APs separados por vírgula: " + Style.RESET_ALL)
    
    if not aps_input.strip():
        print(Fore.RED + "Nenhuma AP informada. Operação cancelada.")
        return
        
    lista_aps = [ap.strip() for ap in aps_input.split(',') if ap.strip()]
    
    # Chama o serviço original que você já tinha
    df_resultado, valores_nao_encontrados = transitorias_services.processar_extracao_aps_transitarias(lista_aps)
    
    if df_resultado is not None and not df_resultado.empty:
        transitorias_services.salvar_dataframe_transitarias(df_resultado)


def _executar_conciliacao_fiscal():
    """A nova funcionalidade de conciliação com a planilha do fiscal."""
    print(Fore.CYAN + "\n[ Conciliação Fiscal ]" + Style.RESET_ALL)
    
    # ===== VALIDAÇÃO INTELIGENTE DA BASE LOCAL =====
    caminho_base = CAMINHO_BASE_PADRAO
    
    if not caminho_base.exists():
        print(Fore.YELLOW + f"⚠️ Base local não encontrada em: {caminho_base}")
        caminho_base_input = input(Fore.BLUE + "Arraste para cá a sua Base de Dados de APs (Excel): " + Style.RESET_ALL).strip()
        # Limpar aspas do input (Windows permite arrastar arquivos entre aspas)
        caminho_base = Path(caminho_base_input.replace('"', ''))
        
        if not caminho_base.exists():
            print(Fore.RED + f"❌ Arquivo não encontrado: {caminho_base}. Operação cancelada." + Style.RESET_ALL)
            return
        
        print(Fore.GREEN + f"✅ Base carregada de: {caminho_base}" + Style.RESET_ALL)
    
    # ===== SOLICITAÇÃO DO ARQUIVO FISCAL =====
    print("Arraste para cá a planilha enviada pelo Fiscal (deve conter AP, NF, FORNECEDOR, VALOR).")
    
    caminho_fiscal = input(Fore.BLUE + "Caminho do arquivo: " + Style.RESET_ALL).strip()
    
    # Remove aspas caso o usuário arraste o arquivo no terminal Windows
    if caminho_fiscal.startswith('"') and caminho_fiscal.endswith('"'):
        caminho_fiscal = caminho_fiscal[1:-1]
        
    if not os.path.isfile(caminho_fiscal):
        print(Fore.RED + f"Arquivo não encontrado: {caminho_fiscal}. Operação abortada." + Style.RESET_ALL)
        return

    print(Fore.YELLOW + "\n⏳ Lendo a planilha Fiscal e cruzando com a base de dados local..." + Style.RESET_ALL)
    
    # Agora a função retorna quatro valores
    df_resultado, df_pendencias, df_overview, resumo = transitorias_services.processar_conciliacao_fiscal(caminho_fiscal)
    
    if df_resultado is None or df_resultado.empty:
        print(Fore.YELLOW + "Nenhum dado pôde ser processado." + Style.RESET_ALL)
        return

    # Feedback claro para o usuário
    print(Fore.GREEN + f"\n✅ Processamento concluído!" + Style.RESET_ALL)
    print(f"  - {resumo.get('total_aps_fiscais', 0)} APs lidas da planilha do Fiscal.")
    print(f"  - {resumo.get('encontradas', 0)} APs conciliadas na base.")
    print(f"  - {resumo.get('nao_encontradas', 0)} APs não encontradas na base local.")
    print(f"  - {resumo.get('divergencias', 0)} APs com divergência de valor.")
    
    salvar = input(Fore.BLUE + "\nDeseja salvar o resultado em um arquivo Excel? (s/n): " + Style.RESET_ALL).strip().lower()
    if salvar == 's':
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f'Transitorias_Conciliadas_{timestamp}.xlsx'
        filename = input(Fore.BLUE + f"Digite o nome do arquivo (padrão: {default_filename}): " + Style.RESET_ALL).strip()
        if not filename:
            filename = default_filename
        
        # Passamos os três DataFrames para a função de salvamento
        transitorias_services.salvar_dataframe_transitarias(
            df_resultado, 
            df_pendencias, 
            df_overview, 
            filename
        )