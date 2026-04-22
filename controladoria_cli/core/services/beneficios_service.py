import pandas as pd
import os
import zipfile

# IMPORT NOVO: Trazendo o mapeamento que você configurou no passo anterior
# ANTES:
# from controladoria_cli.core.config.account_mapping import MAPEAMENTO_BENEFICIOS

# DEPOIS:
from ..config.account_mapping import mapeamento_custo, mapeamento_despesa

CAMINHO_CONCATENACAO_ENTRADA = r"C:/Users/jose.santos/Desktop/CONTROLADORIA/Automacoes/Fast_Control/controladoria_cli/dados_beneficios/concatenacao"
CAMINHO_CONCATENACAO_SAIDA = os.path.join(CAMINHO_CONCATENACAO_ENTRADA, "CONCAT")

CAMINHO_RATEIO = r'C:/Users/jose.santos/Desktop/CONTROLADORIA/Automacoes/Fast_Control/controladoria_cli/dados_beneficios/rateio'
ARQUIVO_RATEIO_ENTRADA = os.path.join(CAMINHO_RATEIO, 'base_rateio.xlsx')

ARQUIVO_RATEIO_SAIDA = os.path.join(CAMINHO_RATEIO, 'rateio_beneficios_gerado.xlsx')

# --- FUNÇÃO NOVA DE CLASSIFICAÇÃO ---
def _obter_conta_contabil(nome_beneficio: str, tipo_custo_despesa: str) -> str:
    """
    Busca a conta contábil baseado no nome do benefício (que já sabemos via loop)
    e na string da coluna 'CUSTO OU DESPESA', usando os dicionários unificados.
    """
    # Exemplo: 'vr' vira 'VR'
    beneficio_upper = str(nome_beneficio).strip().upper()
    # Exemplo: ' Custo ' vira 'CUSTO'
    classificacao_upper = str(tipo_custo_despesa).strip().upper()

    # Se a linha for Custo, busca no dicionário de custo
    if 'CUSTO' in classificacao_upper:
        # Pega a conta, se não achar retorna o aviso
        return mapeamento_custo.get(beneficio_upper, "CONTA NÃO MAPEADA NO CUSTO")
        
    # Se a linha for Despesa, busca no dicionário de despesa
    elif 'DESPESA' in classificacao_upper:
        return mapeamento_despesa.get(beneficio_upper, "CONTA NÃO MAPEADA NA DESPESA")
        
    else:
        return "CONTA NÃO MAPEADA - TIPO CD INVÁLIDO"


def executar_concatenacao():
    """
    Lógica para juntar (concatenar) múltiplos arquivos Excel de um diretório.
    (MANTIDO EXATAMENTE COMO VOCÊ FEZ)
    """
    print("\n--- Executando: Junção de Planilhas ---")
    try:
        os.makedirs(CAMINHO_CONCATENACAO_SAIDA, exist_ok=True)
        
        arquivos = os.listdir(CAMINHO_CONCATENACAO_ENTRADA)
        arquivos_xlsx = [f for f in arquivos if f.endswith(".xlsx")]

        if not arquivos_xlsx:
            print("AVISO: Nenhum arquivo .xlsx encontrado no diretório de entrada.")
            return

        print(f"Encontrados {len(arquivos_xlsx)} arquivos. Processando...")
        dfs = []
        for arquivo in arquivos_xlsx:
            caminho_arquivo = os.path.join(CAMINHO_CONCATENACAO_ENTRADA, arquivo)
            try:
                df = pd.read_excel(caminho_arquivo, engine="openpyxl")
                df['Arquivo Original'] = arquivo
                dfs.append(df)
            except (zipfile.BadZipFile, Exception) as e:
                print(f"Erro ao ler o arquivo {arquivo}: {e}. Pulando.")

        if not dfs:
            print("Nenhum arquivo Excel válido pôde ser lido.")
            return
            
        df_final = pd.concat(dfs, ignore_index=True)
        nome_arquivo_saida = os.path.join(CAMINHO_CONCATENACAO_SAIDA, "arquivo_concatenado_gerado.xlsx")
        df_final.to_excel(nome_arquivo_saida, index=False)
        print(f"SUCESSO! O arquivo foi salvo em: {nome_arquivo_saida}")

    except Exception as e:
        print(f"ERRO GERAL na junção de planilhas: {e}")


def executar_rateio():
    """
    Lógica para executar o rateio em uma base de benefícios, com classificação contábil
    e exportação no layout padrão de integração (ERP).
    """
    print("\n--- Executando: Rateio de Benefícios ---")
    try:
        df_base = pd.read_excel(ARQUIVO_RATEIO_ENTRADA)
        df_base.columns = df_base.columns.str.strip()
        
        # O nome exato da coluna na sua planilha
        coluna_cd = "C / D" 
        
        if coluna_cd not in df_base.columns:
            print(f"ERRO: A coluna '{coluna_cd}' não foi encontrada na planilha base de rateio.")
            print(f"Colunas disponíveis: {list(df_base.columns)}")
            return
        
        beneficios = ['VR', 'CB', 'VA', 'OB', 'AJ', 'VT', 'PJ', 'AUX MORADIA', 'DESPESAS ACIONISTAS','D-BÔNUS', 'Taxa']

        with pd.ExcelWriter(ARQUIVO_RATEIO_SAIDA, engine='xlsxwriter') as writer:
            print("Processando cada tipo de benefício para o novo layout...")
            
            for beneficio in beneficios:
                if beneficio in df_base.columns:
                    df_temp = df_base.copy()
                    df_temp[beneficio] = pd.to_numeric(df_temp[beneficio], errors='coerce')
                    
                    # Filtra apenas quem tem valor válido
                    df_temp = df_temp[df_temp[beneficio].notna() & (df_temp[beneficio] != 0)]
                    
                    if not df_temp.empty:
                        # 1. Aplica a Regra Contábil
                        df_temp['Conta Contabil'] = df_temp[coluna_cd].apply(
                            lambda valor_cd: _obter_conta_contabil(beneficio, valor_cd)
                        )
                        
                        # 2. Agrupa os valores (adicionamos 'coluna_cd' no groupby para não perdê-la)
                        df_aggregated = df_temp.groupby(
                            ['CC', 'Conta Contabil', coluna_cd, 'NOME ARQUIVO']
                        )[beneficio].sum().reset_index()
                        
                        # 3. MOLDANDO O LAYOUT FINAL (DataFrame novo apenas com as colunas solicitadas)
                        df_layout_final = pd.DataFrame()
                        
                        # Mapeando os dados para a estrutura da imagem
                        df_layout_final['CC'] = df_aggregated['CC']
                        
                        # Criando um histórico dinâmico. Ex: "RATEIO VR - FOPAG_MARCO.xlsx"
                        df_layout_final['HISTÓRICO'] = "RATEIO " + beneficio + " - " + df_aggregated['NOME ARQUIVO'].astype(str)
                        
                        df_layout_final['VALOR'] = df_aggregated[beneficio]
                        df_layout_final['BENEFICIO'] = beneficio
                        df_layout_final['C / D'] = df_aggregated[coluna_cd]
                        df_layout_final['CONT.D'] = df_aggregated['Conta Contabil'] # Despesa/Custo entra no Débito
                        df_layout_final['CONT.C'] = "" # Deixando em branco (ou você pode preencher com conta fixa depois)
                        
                        # Salvando na aba correspondente
                        df_layout_final.to_excel(writer, sheet_name=beneficio, index=False)
                else:
                    print(f"AVISO: A coluna '{beneficio}' não foi encontrada na planilha.")

        print(f"SUCESSO! Planilha de rateio criada com o novo layout em: {ARQUIVO_RATEIO_SAIDA}")

    except FileNotFoundError:
        print(f"ERRO: O arquivo de entrada não foi encontrado em: {ARQUIVO_RATEIO_ENTRADA}")
    except Exception as e:
        print(f"ERRO GERAL no processo de rateio: {e}")