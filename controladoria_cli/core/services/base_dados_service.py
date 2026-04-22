import logging
import re
import pandas as pd
import pyodbc
from typing import List

# Configurando o logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class BaseDadosService:
    def __init__(self, db_settings):
        """Inicializa o serviço com as configurações do banco de dados."""
        self.db_settings = db_settings
        self.connection_string = (
            f"DRIVER={{{self.db_settings.DB_DRIVER}}};"
            f"SERVER={self.db_settings.DB_SERVER};"
            f"DATABASE={self.db_settings.DB_NAME};"
            f"UID={self.db_settings.DB_USER};"
            f"PWD={self.db_settings.DB_PASSWORD};"
            "TrustServerCertificate=yes;"
        )
        self.sql_sup = """
            SELECT
                TB_Titulo.NumeroTituloPrincipal, ClienteFornec AS sup_ClienteFornec,
                TB_Empresa.cgc AS sup_cnpj, TB_Titulo.CodigoEmpresa AS sup_CodigoEmpresa,
                CodigoBanco AS sup_CodigoBanco, Status AS sup_Status, DataBaixa AS sup_DataBaixa,
                DataCompetencia AS sup_DataCompetencia, ValorTitulo AS sup_ValorTitulo,
                ValorJuros AS sup_ValorJuros, ValorDesconto AS sup_ValorDesconto,
                ValorMulta AS sup_ValorMulta, Duplicata AS sup_Duplicata,
                ValorDesctoAdtoForn AS sup_ValorDesctoAdtoForn,
                TB_Titulo.NumeroTituloPrincipal AS sup_NumeroTituloPrincipal, NumeroAP AS sup_NumeroAp,
                TB_Titulo.Cheque AS sup_Cheque, Rateio AS sup_valid_rateio,
                DescHistorico AS SUP_DescHistorico, CodigoCentroCusto AS SUP_CodigoCentroCusto,
                Portador AS SUP_Portador, CodigoDespesa AS SUP_CodigoDespesa, Descricao AS SUP_Descricao
            FROM TB_Titulo
            LEFT JOIN TB_TipoDespesa ON TB_TipoDespesa.CodigoDespesa = TB_Titulo.Despesa
            LEFT JOIN TB_Empresa ON CONCAT(tb_empresa.CodigoEmpresa, tb_empresa.CodigoFilial) = CONCAT(TB_Titulo.CodigoEmpresa, TB_Titulo.CodigoFilial)
            WHERE
                STATUS = 'B' AND
                CAST(DataCompetencia as INT) BETWEEN ? AND ? AND
                NumeroTituloPrincipal = '' AND
                NumeroAP != 0
        """
        self.sql_inf = """
            SELECT
                DescHistorico AS INF_DescHistorico, CodigoCentroCusto AS INF_CodigoCentroCusto,
                DataCompetencia AS INF_DataCompetencia, ValorTitulo AS INF_ValorTitulo,
                Portador AS INF_Portador, NumeroAP AS INF_NumeroAP,
                NumeroAPPrestacaoContas AS INF_NumeroAPPrestacaoContas, CodigoDespesa AS INF_CodigoDespesa,
                Descricao AS INF_Descricao, TB_Titulo.Cheque AS INF_Cheque,
                TB_Titulo.NumeroTituloPrincipal AS inf_NumeroTituloPrincipal,
                CASE
                    WHEN NumeroTituloPrincipal = '' THEN 'LINHA PRINCIPAL - SUPERIOR'
                    ELSE 'LINHA RATEIO'
                END AS TIPO_DADO
            FROM TB_Titulo
            LEFT JOIN TB_TipoDespesa ON TB_TipoDespesa.CodigoDespesa = TB_Titulo.Despesa
            WHERE
                STATUS = 'R' AND
                NumeroAP IN ({}) AND
                CodigoBancoOri != 999 AND
                NumeroTituloPrincipal != ''
        """

    def _get_superior_data(self, conn, start_date: str, end_date: str) -> pd.DataFrame:
        """Busca os dados superiores do banco de dados."""
        logging.info(f"Executando query superior para o período de {start_date} a {end_date}.")
        # Usamos parâmetros aqui pois os valores vêm do usuário
        return pd.read_sql_query(self.sql_sup, conn, params=[start_date, end_date])

    def _get_inferior_data(self, conn, ap_numbers: list) -> pd.DataFrame:
        """Busca os dados inferiores em lotes para evitar o limite de parâmetros."""
        if not ap_numbers:
            return pd.DataFrame()

        logging.info(f"Executando query inferior para {len(ap_numbers)} APs.")
        all_inf_dfs = []
        # O limite do SQL Server para itens em uma cláusula IN é de alguns milhares.
        # Vamos limitar cada consulta a 1000 itens para segurança.
        chunk_size = 1000

        for i in range(0, len(ap_numbers), chunk_size):
            chunk = ap_numbers[i:i + chunk_size]
            
            # Converte cada item para string para a junção. Como os números vêm do banco,
            # não há risco de SQL injection aqui.
            ap_list_str = ','.join(map(str, chunk))
            
            # Formata a string SQL diretamente com os números das APs
            sql_inf_chunk = self.sql_inf.format(ap_list_str)
            
            logging.info(f"Processando lote de {len(chunk)}/{len(ap_numbers)} APs...")
            try:
                # Agora não precisamos do parâmetro 'params' pois os valores já estão na query
                df_chunk = pd.read_sql_query(sql_inf_chunk, conn)
                all_inf_dfs.append(df_chunk)
            except pyodbc.Error as e:
                logging.error(f"Erro ao executar a query para o lote de APs: {e}")
                logging.error(f"Query com falha (primeiros 500 caracteres): {sql_inf_chunk[:500]}...")
                raise e

        if not all_inf_dfs:
            logging.warning("Nenhum dado de rateio (inferior) correspondente encontrado.")
            return pd.DataFrame()

        return pd.concat(all_inf_dfs, ignore_index=True)


    def _clean_text(self, text: str) -> str:
        """Remove caracteres problemáticos de uma string."""
        if isinstance(text, str):
            cleaned_text = text.replace('▼', '')
            cleaned_text = re.sub(r'[^\x00-\x7F]+', '', cleaned_text)
            return cleaned_text
        return text

    def get_and_process_aps(self, start_date: str, end_date: str) -> pd.DataFrame:
        """
        Orquestra a extração e processamento dos dados de APs.
        """
        try:
            with pyodbc.connect(self.connection_string) as conn:
                logging.info("Conexão com o banco de dados estabelecida com sucesso.")
                
                df_sup = self._get_superior_data(conn, start_date, end_date)
                if df_sup.empty:
                    logging.warning("Nenhum dado superior encontrado para o período.")
                    return pd.DataFrame()

                ap_numbers = df_sup['sup_NumeroAp'].unique().tolist()
                df_inf = self._get_inferior_data(conn, ap_numbers)

                logging.info("Juntando e processando os dados...")
                df = pd.merge(df_sup, df_inf, how='left', left_on='sup_NumeroAp', right_on='INF_NumeroAP')

                df_ok = df[df['INF_NumeroAP'].notna()]
                df_nok = df[df['INF_NumeroAP'].isna()].copy()

                if not df_nok.empty:
                    df_nok['INF_DescHistorico'] = df_nok['SUP_DescHistorico']
                    df_nok['INF_CodigoCentroCusto'] = df_nok['SUP_CodigoCentroCusto']
                    df_nok['INF_DataCompetencia'] = df_nok['sup_DataCompetencia']
                    df_nok['INF_ValorTitulo'] = df_nok['sup_ValorTitulo']
                    df_nok['INF_Portador'] = df_nok['SUP_Portador']
                    df_nok['INF_NumeroAP'] = df_nok['sup_NumeroAp']
                    df_nok['INF_Cheque'] = df_nok['sup_Cheque']
                    df_nok['inf_NumeroTituloPrincipal'] = df_nok['sup_NumeroTituloPrincipal']
                    df_nok['TIPO_DADO'] = 'LINHA PRINCIPAL - SUPERIOR'

                df_final = pd.concat([df_nok, df_ok], ignore_index=True)

                logging.info("Limpando dados...")
                for col in df_final.columns:
                    if df_final[col].dtype == 'object':
                        df_final[col] = df_final[col].apply(self._clean_text)
                
                df_final.columns = [self._clean_text(col) for col in df_final.columns]
                
                logging.info("Processamento concluído com sucesso.")
                return df_final

        except pyodbc.Error as ex:
            sqlstate = ex.args[0]
            logging.error(f"Erro de banco de dados: {sqlstate} - {ex}")
            raise ConnectionError(f"Falha na comunicação com o banco de dados: {ex}")
        except Exception as e:
            logging.error(f"Ocorreu um erro inesperado durante o processamento de dados: {e}")
            raise

