# controladoria_cli/commands/beneficios.py

from colorama import Fore
from ..core.services import beneficios_service

def exibir_menu_beneficios():
    """Exibe o sub-menu de opções para a área de Benefícios."""
    print(Fore.YELLOW + "\n--- Menu de Benefícios ---")
    print(Fore.YELLOW + "Selecione a ferramenta que deseja utilizar:")
    print("1 - Juntar (Concatenar) Planilhas")
    print("2 - Executar Rateio de Benefícios")
    print("0 - Voltar ao Menu Principal")
    return input("Digite sua opção: ")

def run_beneficios_interativo():
    """
    Loop principal que executa o sub-menu da área de Benefícios.
    """
    while True:
        escolha = exibir_menu_beneficios()
        
        if escolha == '1':
            try:
                # Chama a função de concatenação, agora com nome genérico
                beneficios_service.executar_concatenacao()
            except Exception as e:
                print(Fore.RED + f"[ERRO INESPERADO] Ocorreu um problema: {e}")

        elif escolha == '2':
            try:
                # Chama a função de rateio, agora com nome genérico
                beneficios_service.executar_rateio()
            except Exception as e:
                print(Fore.RED + f"[ERRO INESPERADO] Ocorreu um problema: {e}")

        elif escolha == '0':
            print("Voltando ao menu principal...")
            break
        
        else:
            print(Fore.RED + "Opção inválida, por favor tente novamente.")