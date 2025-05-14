import pandas as pd
import os
from tqdm import tqdm

def consolida_rem():
    print('Consolidação arquivo de REM ...')
    print()
    # Solicita o caminho
    caminho_arquivos = input('Digite o caminho dos arqui    vos para consolidação: ')
    caminho_saida = input('Digite onde será salvo o arquivo final de REM: ')

    # Cria Df vazio para armazenagem
    df_concatenado = pd.DataFrame()

    # Faz uma varredura na pasta de tudo que contém informe de emissões e armazena no df_concatenado
    print('Juntando arquivos...')
    for nome_arquivo in tqdm(os.listdir(caminho_arquivos), desc='Consolidação de arquivos', unit='arquivo'):
            if 'Informe de Emissões' in nome_arquivo:
                try:
                    caminho_completo = os.path.join(caminho_arquivos,nome_arquivo)
                    df = pd.read_excel(caminho_completo,sheet_name="Auditoria GERAL")
                    df_concatenado = pd.concat([df_concatenado, df], ignore_index=True)
                except Exception as e:
                    print(f'Erro {e}, O Arquivo {nome_arquivo} ignorado pois não possui a guia Auditoria GERAL')

    # Monta o caminho completo do arquivo de saída
    caminho_arquivo_saida = os.path.join(caminho_saida, 'informe de emissões.xlsx')

    # Salva o DataFrame em Excel
    df_concatenado.to_excel(caminho_arquivo_saida, sheet_name='informe de emissões', index=False)

    print('Salvo com sucesso')
