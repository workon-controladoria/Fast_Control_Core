import pandas as pd
from colorama import Fore
import os


from ..config import account_mapping

# CAMINHO FIXO DE ENTRADA
CAMINHO_ARQUIVO_ENTRADA = 'C:/Users/jose.santos/Desktop/CONTROLADORIA/Automacoes/Fast_Control/controladoria_cli/dados_aps/data_base_de_aps.xlsx'

# --- CAMINHO DE SAÍDA PARA VOCÊ PREENCHER ---
# Coloque o caminho completo para a sua pasta 'output' dentro das aspas.
# Exemplo: 'C:/Users/jose.santos/Desktop/Automacoes/Fast_Control/output'
CAMINHO_PASTA_SAIDA = 'C:/Users/jose.santos/Desktop/CONTROLADORIA/Automacoes/Fast_Control/controladoria_cli/dados_aps/output_dados_aps'


def _determinar_conta_contabil(row):
    try:
        natureza = str(row['NATUREZA']).strip()
        classificacao = str(row['CUSTO OU DESPESA']).strip().upper()
        if 'CUSTO' in classificacao:
            return account_mapping.mapeamento_custo.get(natureza, 'Custo Não Mapeado')
        elif 'DESPESA' in classificacao:
            return account_mapping.mapeamento_despesa.get(natureza, 'Despesa Não Mapeada')
        else:
            return 'Classificação Inválida'
    except KeyError:
        return "Coluna Essencial Faltando"

def processar_extracao_aps_transitarias(valores_aps: list):

    try:
        df = pd.read_excel(CAMINHO_ARQUIVO_ENTRADA)
    except FileNotFoundError:
        print(Fore.RED + f"ERRO: Arquivo de entrada não encontrado em '{CAMINHO_ARQUIVO_ENTRADA}'")
        return None, valores_aps
    if 'APS' not in df.columns:
        print(Fore.RED + "ERRO: A coluna 'APS' não foi encontrada no arquivo.")
        return None, valores_aps
    colunas_necessarias = ['NATUREZA', 'CUSTO OU DESPESA']
    df['APS'] = df['APS'].astype(str)
    valores_aps_str = [str(v).strip() for v in valores_aps]
    df_filtrado = df[df['APS'].isin(valores_aps_str)].copy()
    valores_encontrados = list(df_filtrado['APS'].unique())
    valores_nao_encontrados = [v for v in valores_aps_str if v not in valores_encontrados]
    if all(col in df_filtrado.columns for col in colunas_necessarias):
        print(Fore.MAGENTA + "Aplicando mapeamento de contas contábeis...")
        df_filtrado['conta contabil'] = df_filtrado.apply(_determinar_conta_contabil, axis=1)
        print(Fore.GREEN + "Coluna 'conta contabil' criada com sucesso.")
    return df_filtrado, valores_nao_encontrados

# --- FUNÇÃO ALTERADA ---
def salvar_dataframe_transitarias(df: pd.DataFrame, file_name: str):
    """Salva um DataFrame no diretório de saída especificado na variável global."""
    if df is None or df.empty:
        print(Fore.YELLOW + "Nenhum dado para salvar.")
        return
        
    try:

        if not CAMINHO_PASTA_SAIDA or not CAMINHO_PASTA_SAIDA.strip():
            print(Fore.RED + "ERRO CRÍTICO: O caminho de saída não foi definido.")
            print(Fore.YELLOW + "Por favor, edite o arquivo 'services/transitorias_service.py' e preencha a variável 'CAMINHO_PASTA_SAIDA'.")
            return

        # Garante que a pasta de saída exista. Se não existir, ela será criada.
        os.makedirs(CAMINHO_PASTA_SAIDA, exist_ok=True)
            
        # Junta o caminho da pasta com o nome do arquivo para ter o caminho completo
        output_path = os.path.join(CAMINHO_PASTA_SAIDA, file_name)

        df.to_excel(output_path, index=False)
        print(Fore.CYAN + f"Dados filtrados salvos com sucesso em: {output_path}")
    except Exception as e:
        print(Fore.RED + f"Ocorreu um erro ao salvar o arquivo: {e}")