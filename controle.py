import pandas as pd
import os

ARQUIVO_EXCEL = 'dados/estoque_validade.xlsx'
NOME_PLANILHA_ESTOQUE = 'Estoque'
NOME_PLANILHA_EXCLUIDOS = 'Excluidos'

COLUNAS_ESTOQUE = ['Produto', 'Categoria', 'Validade', 'Lote', 'Data Recebimento']
COLUNAS_EXCLUIDOS = ['Produto', 'Categoria', 'Validade', 'Lote', 'Data Recebimento', 'Data Exclusao']

def _handle_datetime_conversion(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors='coerce')

def carregar_dados():
    df_estoque = pd.DataFrame(columns=COLUNAS_ESTOQUE)
    df_excluidos = pd.DataFrame(columns=COLUNAS_EXCLUIDOS)

    if os.path.exists(ARQUIVO_EXCEL):
        excel_data = pd.read_excel(ARQUIVO_EXCEL, sheet_name=None, engine='openpyxl')
        if NOME_PLANILHA_ESTOQUE in excel_data:
            df_estoque = excel_data[NOME_PLANILHA_ESTOQUE].dropna(how='all')
            df_estoque['Validade'] = _handle_datetime_conversion(df_estoque['Validade'])
            df_estoque['Data Recebimento'] = _handle_datetime_conversion(df_estoque['Data Recebimento'])
            df_estoque['Lote'] = df_estoque['Lote'].fillna('').astype(str)
            df_estoque['Categoria'] = df_estoque['Categoria'].fillna('').astype(str)
        if NOME_PLANILHA_EXCLUIDOS in excel_data:
            df_excluidos = excel_data[NOME_PLANILHA_EXCLUIDOS].dropna(how='all')
            df_excluidos['Validade'] = _handle_datetime_conversion(df_excluidos['Validade'])
            df_excluidos['Data Recebimento'] = _handle_datetime_conversion(df_excluidos['Data Recebimento'])
            df_excluidos['Data Exclusao'] = _handle_datetime_conversion(df_excluidos['Data Exclusao'])
            df_excluidos['Lote'] = df_excluidos['Lote'].fillna('').astype(str)
            df_excluidos['Categoria'] = df_excluidos['Categoria'].fillna('').astype(str)
    return df_estoque, df_excluidos
