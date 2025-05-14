import pandas as pd
import os

def consolida_arquivos_finais():
    print('Inicializando consolidação final...')
    print()
    # Solicita o caminhos
    caminho_arquivos = input('Digite o caminho dos arquivos para consolidação: \n')

    # Define os nomes dos arquivos
    arquivo_can = 'Informe de Cancelamento.xlsx'
    arquivo_prov = 'Informe de Faturamento Devidos_.xlsx'
    arquivo_per = 'Informe de Perdas.xlsx'
    arquivo_est = 'Controle de Devidos e Provisão Não Faturados.xlsx'
    arquivo_rem = 'informe de emissões.xlsx'
    arquivo_fat = 'Informe de receita diferida.xlsx'

    # Cria o caminho completo dos arquivos
    caminho_can = os.path.join(caminho_arquivos, arquivo_can)
    caminho_prov = os.path.join(caminho_arquivos, arquivo_prov)
    caminho_per = os.path.join(caminho_arquivos, arquivo_per)
    caminho_est = os.path.join(caminho_arquivos, arquivo_est)
    caminho_rem = os.path.join(caminho_arquivos,arquivo_rem)
    caminho_fat = os.path.join(caminho_arquivos,arquivo_fat)

    #Escolhendo sheet diferida
    arquivo_diferidas = pd.ExcelFile(caminho_fat)
    nome_guias_diferidas = arquivo_diferidas.sheet_names
    guias_encontradas = []


    for guia in nome_guias_diferidas:
        if 'Auditoria' in guia:
            guias_encontradas.append(guia)
    print('Encontramos as seguintes guias abaixo na base de informe de diferidas: ')
    print()
    print(f'{guias_encontradas} que contém a palavra Auditoria')

    opcao = input('Digite o nome da guia a considerar ? obs: digite sem as aspas: \n')
    print()
    print(f'guia escolhida: {opcao} \n')

    print('Juntando arquivos...')
    print()

    # Lê os arquivos como DataFrame
    df_can = pd.read_excel(caminho_can,sheet_name="Auditoria")
    df_prov = pd.read_excel(caminho_prov)
    df_per = pd.read_excel(caminho_per,sheet_name="Auditoria")
    df_est = pd.read_excel(caminho_est,sheet_name="BAIXADOS")
    df_rem = pd.read_excel(caminho_rem)
    df_fat = pd.read_excel(caminho_fat,sheet_name=opcao, header=1)

    # Adiciona as colunas adicionais
    df_can['Origin'] = 'CAN'
    df_per['Origin'] = 'PER'
    df_est['Origin'] = 'EST'
    df_rem['Origin'] = 'REM'
    df_fat['Origin'] = 'FAT'

    #transformando valor em negativo
    df_can['PR_TOTAL'] = df_can ['PR_TOTAL'] * -1
    df_per['PR_TOTAL'] = df_per['PR_TOTAL'] * -1
    df_est['PR_TOTAL'] = df_est['PR_TOTAL'] * -1

    # Concatena os DataFrames
    df_concatenado = pd.concat([df_can, df_prov, df_per,df_est, df_rem, df_fat])

    colunas_desejadas = [
        'Origin','CS_ARID', 'CP_NAME', 'CS_ID', 'CS_NAME', 'CA_DVID', 'CS_CYID', 'CS_CPID',
        'ST_CDID', 'CD_NAME', 'BIID', 'BCID', 'BTTYPE', 'CS_ACCTSTAT', 'DP_DESC',
        'DP_DAYS', 'DP_DATE', 'ST_BTDESC', 'ITEM', 'QTDE', 'PR_UNIT', 'PR_TOTAL',
        'IRRF', 'COFINS_PIS_CSLL', 'ISS', 'FAT_MIN', 'CUST_MIN', 'PROCESSADO_PRECO',
        'PRECO_EXISTENTE', 'CA_CNTRCT', 'STATUS_LANCTO', 'FATURADO', 'NUMERO DA NOTA'
    ]

    # Apenas mantém as colunas que existem no dataframe
    colunas_existentes = [col for col in colunas_desejadas if col in df_concatenado.columns]
    df_final = df_concatenado[colunas_existentes]

    # Salva o arquivo final
    caminho_saida = os.path.join(caminho_arquivos, 'arquivo_final_teste.xlsx')
    df_final.to_excel(caminho_saida, index=False)

    print("Arquivo consolidado com sucesso!")