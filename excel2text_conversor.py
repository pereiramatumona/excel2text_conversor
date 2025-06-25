import pandas as pd
import time
import os
from colorama import Fore, Back, Style, init

from pathlib import Path

def read_excel_file(file, ano):
    try:
        excel_data_df = pd.read_excel(file, sheet_name=ano, skiprows=11 )

    except FileNotFoundError as e:
        display(ficheiro='None', aba='None', acao='Ficheiro não encontrado...')
        print('ERRO:', e)
    except FileNotFoundError as e:
        display(ficheiro='None', aba='None', acao='Aba não encontrada no ficheiro...')
    else:
        csv_format = excel_data_df.to_csv('000100001100012.bin', sep=';', index=False)
    
        return csv_format

def create_csv_file():

    with open('000100001100012.bin', 'r', encoding="utf-8") as novo_csv:
        for linhas in novo_csv:
            if linhas.split(';')[0] != '':
                inss = f'1{linhas.split(';')[0]:0>10}'
                nome = linhas.split(';')[1]
                salarios = f'00000{linhas.split(';')[2]:0>14}{linhas.split(';')[3]:0>14}00000000'
                salario_info = ''.join(salarios.split('.'))

                linha_formatada = f'{inss:10}{' ':20}{nome:70}{salario_info:41}{' ':38}'
                salvar_formatado(linha_formatada)
            else:
                break


def salvar_formatado(linha_formatada):
    with open('Final_file.txt', 'a') as novo_csv:
        novo_csv.write(linha_formatada)
    
    #apagar_ficcheiro()

def display(ficheiro='None', aba='None', acao='None'):
    os.system('cls')
    traco = '-'
    barra = '|'
    print(f'{traco*81}')
    print( f'| {Style.BRIGHT}{"INFORMACOES DE FICHEIRO DE CONVERSÃO":^77}{Style.RESET_ALL} |')
    print(f'{traco*81}')
    print(f'|{"":^79}|')
    print(f'|{" FICHEIRO(xxx.xlsx):":<20} {Fore.RED }{ficheiro:58}{Style.RESET_ALL}|') if ficheiro == 'None' else print(f'|{" FICHEIRO(XXX.XLSX):":<20} {Fore.GREEN}{Style.BRIGHT}{ficheiro:58}{Style.RESET_ALL}|')
    print(f'|{" ABA:":<20} {Fore.RED }{aba:58}{Style.RESET_ALL}|') if aba == 'None' else print(f'|{" ABA:":<20} {Fore.GREEN}{Style.BRIGHT}{aba:58}{Style.RESET_ALL}|')  
    print(f'|{" ACÇÂO:":<20} {Fore.RED }{acao:58}{Style.RESET_ALL}|') if acao == 'None' else print(f'|{" ACÇÂO:":<20} {Fore.GREEN}{Style.BRIGHT}{acao:58}{Style.RESET_ALL}|')
    print(f'|{"":^79}|')
    print(f'{traco*81}')
    print(f'|{"INFO: O ficheiro Excel deve estar no mesmo diretorio com o conversor executavel":79}|')
    print(f'|{"INFO: Respeitar o maiúsculo/minúsculo para que a conversão Funcione… ":79}|')
    print(f'{traco*81}')
    print('')
    time.sleep(1)


#Pagar o ficheiro Temporario
def apagar_ficcheiro():
    # Usando os.remove()
    try:
        os.remove("new_csv_file.txt")
        print("Ficheiro apagado com sucesso.")
    except OSError as e:
        print(f"Erro ao apagar ficheiro: {e}")

    # Usando pathlib.Path.unlink()
    caminho_ficheiro = Path("new_csv_file.txt")
    try:
        caminho_ficheiro.unlink()
        print("Ficheiro apagado com sucesso.")
    except FileNotFoundError:
        print("Ficheiro não encontrado.")
    except OSError as e:
        print(f"Erro ao apagar ficheiro: {e}")

    # Usando pathlib.Path.unlink() com missing_ok=True
    caminho_ficheiro = Path("new_csv_file.txt")
    caminho_ficheiro.unlink(missing_ok=True)
    print("Ficheiro apagado com sucesso (se existir).")


def interacao_user():
    
    ficheiro = 'None'
    aba = 'None'
    acao = 'None'

    display() 
    
    ficheiro = input('Digite o nome do Ficheiro: ')
    display(ficheiro)

    aba = input('Digite a ABA no Ficheiro: ')
    display(ficheiro,aba)    
    
    opcao = input('Desejas converter? Y(Yes)/N(No): ') 
    match opcao:
        case 'Y':
            display(ficheiro,aba,acao='Ficheiro convertido com sucesso...')
            read_excel_file(ficheiro,aba)
            create_csv_file()
        case 'y': 
            display(ficheiro,aba,acao='Ficheiro convertido com sucesso...')
            read_excel_file(ficheiro,aba)
            create_csv_file()
        case 'Yes': 
            display(ficheiro,aba,acao='Ficheiro convertido com sucesso...')
            read_excel_file(ficheiro,aba)
            create_csv_file()
        case 'yes': 
            display(ficheiro,aba,acao='Ficheiro convertido com sucesso...')
            read_excel_file(ficheiro,aba)
            create_csv_file()
        case 'YES': 
            display(ficheiro,aba,acao='Ficheiro convertido com sucesso...')
            read_excel_file(ficheiro,aba)
            create_csv_file()
        case 'N':print('Sair...')
        case 'n': print('Sair...')
        case 'No': print('Sair...')
        case 'no': print('Sair...')
        case 'NO': print('Sair...')
        case _: interacao_user()
    


interacao_user()



#display(ficheiro='Ficheiro_de_ano_2025.xlsx', aba='Janeiro', acao='None')
#read_excel_file('dados.xlsx', 'Janeiro')
#create_csv_file()