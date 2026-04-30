from colorama import Fore
from .commands import base_dados
from .commands import beneficios
from .commands import transitorias
from .commands import provisoes

def exibir_menu():
    """Exibe o menu principal de opções para o usuário."""
    print(Fore.CYAN + "\n--- FAST CONTROL ---")
    print(Fore.CYAN + "Selecione o serviço que deseja utilizar:")
    print(Fore.CYAN + "1 - Base de Dados: Extrair APs") 
    print(Fore.CYAN + "2 - Benefícios")
    print(Fore.GREEN + "3 - Transitórias: Extrair APs com Mapeamento") 
    print(Fore.CYAN + "4 - Provisões")
    print(Fore.CYAN + "0 - Sair")
    return input("Digite sua opção: ")


def main():
    """Loop principal que executa a aplicação de terminal interativo."""
    while True:
        escolha = exibir_menu()
        
        if escolha == '1':
            try:
                base_dados.run_extracao_aps_interativo()
            except Exception as e:
                print(Fore.RED + f"\n[ERRO] Ocorreu um problema na extração de APs: {e}")
        

        elif escolha == '2':
            try:
                beneficios.run_beneficios_interativo()
            except Exception as e:
                print(Fore.RED + f"\n[ERRO] Ocorreu um problema no módulo de Benefícios: {e}")
        
        elif escolha == '3':
            try:
                transitorias.run_extracao_transitarias_interativo()
            except Exception as e:
                print(Fore.RED + f"\n[ERRO] Ocorreu um problema no módulo de Transitórias: {e}")
        
        elif escolha == '4':
            try:
                provisoes.run_provisoes_interativo()
            except Exception as e:
                print(Fore.RED + f"\n[ERRO] Ocorreu um problema no módulo de Provisões: {e}")

        elif escolha == '0':
            print(Fore.GREEN + "Encerrando a aplicação.")
            break
            
        else:
            print(Fore.RED + "\n[AVISO] Opção inválida. Por favor, tente novamente.")

if __name__ == "__main__":
    main()