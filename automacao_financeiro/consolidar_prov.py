import os
from time import thread_time
from tqdm import tqdm
import pandas as pd
import warnings

# Ignorar as advertências do tipo UserWarning
warnings.filterwarnings("ignore", category=UserWarning)
pd.options.mode.chained_assignment = None

def juntar_arquivos_prov():
    # Solicita os caminhos de entrada e saída
    print('Inicializando junção arquivos Prov...')
    print()
    caminho_entrada_prov = input('Digite o caminho da pasta onde estão os gerenciais: ')
    nome_arquivo_prov = 'Informe de Faturamento Devidos_.xlsx'
    nome_arquivo_log = 'log.txt'
    caminho_saida = input('Digite o caminho onde será salvo o arquivo consolidado de PROV: ')
    print()
    arquivo_log = os.path.join(caminho_entrada_prov, nome_arquivo_log)

    # Cria um DataFrame vazio para armazenar os dados combinados
    df_principal = pd.DataFrame()

    # Lista todos os arquivos e diretórios na pasta
    for nome_arquivo in tqdm(os.listdir(caminho_entrada_prov), desc='Consolidação de arquivos', unit='arquivo'):
        caminho_arquivo = os.path.join(caminho_entrada_prov, nome_arquivo)

        if os.path.isfile(caminho_arquivo):
            try:
                abas = pd.ExcelFile(caminho_arquivo).sheet_names
                if 'Auditoria' in abas:
                    df = pd.read_excel(caminho_arquivo, sheet_name='Auditoria')

                    # Verifica se existe a coluna 'BIID'
                    if 'BIID' not in df.columns:
                        with open(arquivo_log, 'a') as arquivo:
                            arquivo.write(f"Arquivo '{nome_arquivo}' não contém a coluna 'BIID'. Ignorado.\n")
                        continue

                    # Filtra registros cujo BIID termina com 'C' ou 'c'
                    df_validos = df[df['BIID'].astype(str).str.endswith(('C', 'c'))]

                    if df_validos.empty:
                        with open(arquivo_log, 'a') as arquivo:
                            arquivo.write(f"Arquivo '{nome_arquivo}' não possui BIIDs terminando com 'C' ou 'c'. Ignorado.\n")
                        continue

                    df_principal = pd.concat([df_principal, df_validos], ignore_index=True)

                else:
                    with open(arquivo_log, 'a') as arquivo:
                        arquivo.write(f"Arquivo '{nome_arquivo}' não contém a aba 'Auditoria'. Ignorado.\n")
            except Exception as e:
                with open(arquivo_log, 'a') as arquivo:
                    arquivo.write(f"Erro ao ler o arquivo '{nome_arquivo}': {str(e)}\n")

    if not df_principal.empty:
        # Filtra as colunas desejadas
        colunas_desejadas = [
            'CS_ARID', 'CP_NAME', 'CS_ID', 'CS_NAME', 'CA_DVID', 'CS_CYID', 'CS_CPID',
            'ST_CDID', 'CD_NAME', 'BIID', 'BCID', 'BTTYPE', 'CS_ACCTSTAT', 'DP_DESC',
            'DP_DAYS', 'DP_DATE', 'ST_BTDESC', 'ITEM', 'QTDE', 'PR_UNIT', 'PR_TOTAL',
            'IRRF', 'COFINS_PIS_CSLL', 'ISS', 'FAT_MIN', 'CUST_MIN', 'PROCESSADO_PRECO',
            'PRECO_EXISTENTE', 'CA_CNTRCT', 'STATUS_LANCTO', 'FATURADO', 'NUMERO DA NOTA'
        ]

        # Apenas mantém as colunas que existem no dataframe
        colunas_existentes = [col for col in colunas_desejadas if col in df_principal.columns]
        df_final = df_principal[colunas_existentes]

        # Ajusta as colunas adicionais
        df_final['Origin'] = 'PROV'
        df_final['LEGACY_ID'] = '0'

        # Salva o DataFrame combinado em um novo arquivo Excel
        df_final.to_excel(os.path.join(caminho_saida, nome_arquivo_prov), index=False)
        print('Arquivo consolidado salvo com sucesso!')
    else:
        print('Nenhum dado foi consolidado.')


