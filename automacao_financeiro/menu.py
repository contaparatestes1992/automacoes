from consolidar_prov import juntar_arquivos_prov
from consolidar_rem import consolida_rem
from consolidar_arquivo_final import consolida_arquivos_finais

def exibir_menu():
    print("Escolha uma opção:")
    print("1 - Iniciar")
    print("2 - Sair")

def sub_menu():
    print("1 - Consolidar arquivos prov")
    print("2 - Consolidar REM")
    print("3 - Gerar relatório mensal")
    print("4 - Voltar ao menu principal \n")

while True:
    exibir_menu()
    opcao = input('Digite uma opção: ')

    match opcao:
        case '1':
            while True:
                sub_menu()
                sub_opcao = input('Digite uma opção do submenu: \n')
                match sub_opcao:
                    case '1':
                        juntar_arquivos_prov()
                    case '2':
                        consolida_rem()
                    case '3':
                        consolida_arquivos_finais()
                    case '4':
                        break
                    case _:
                        print('Opção inválida no submenu.')
        case '2':
            break
        case _:
            print('Opção inválida, por favor escolha novamente.')
