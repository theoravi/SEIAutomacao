
import pandas as pd
from openpyxl import load_workbook
# from tabulate import tabulate
# import numpy as np
# import funcoes as fc
# from Levenshtein import distance as lv

# # # Criação do DataFrame
# preenchido = {'Modelo': ["uv k5", "dm 701", "uv-k5"], 'Nome Comercial': ["MINI 3", "C5", ' '], 'Número de Série (incluindo rádio controle e óculos)': ["1581FFGDF564GFDRA45", "5HAZ54GDGDG6", " "]}
# df = pd.DataFrame(preenchido)
# modelos = df['Modelo'].dropna().reset_index(drop=True)

# print(df)
# print(modelos)
# print("\n")

# # Exibe a tabela formatada
# df = df.replace(' ', np.nan).dropna(how='all')
# data = df.values.tolist()
# headers = df.columns.tolist()
# print(tabulate(data, headers=headers, tablefmt='pretty'))
# print('\n')

# planilhaDrones = input("Insira o caminho da planilha/lista de rádios conformes: ").replace('"' , '')

# # Tenta carregar a planilha de drones conformes
# try:
#     tabela = fc.corrige_planilha(planilhaDrones)
#     tabela_modelos = tabela.astype(str)  # Garante que os modelos estão como strings
# except Exception as e:
#     print(f"Erro ao carregar a planilha: {e}")
#     exit()
    
# # Verificação de conformidade
# checkexcel = [False] * len(modelos)  # Inicializa a lista com False
# lista_parecidos = []
# # Comparação de modelos
# for modelo_solicitante in range(len(modelos)):
#     for j in range(len(tabela_modelos)):
#         try:
#             if modelos[modelo_solicitante].lower().replace(' ','').strip('\n').replace('-', '') == tabela_modelos[j].lower().replace(' ','').replace('-', ''):
#                 checkexcel[modelo_solicitante] = True
#                 print('O modelo', modelos[modelo_solicitante], 'bate com o modelo', tabela_modelos[j], 'da tabela de rádios conformes')
#                 break
#             elif lv(modelos[modelo_solicitante].lower().replace(' ','').strip('\n').replace('-', ''), tabela_modelos[j].lower().replace(' ','').replace('-', '')) < 2:
#                 lista_parecidos.append(tabela_modelos[j])
#         except KeyError as e:
#             print(f"Erro ao acessar os índices: {e}")
#             continue
# for i in range(len(checkexcel)):
#     if not checkexcel[i]:
#         print(f"O modelo {modelos[i]} não se encontra na lista de rádios conformes\n", "Lista de modelos parecidos na tabela:", lista_parecidos)
# print('\n')

# usuarios = {"anasantos.estagio": "Ana Karolina Fernandes dos Santos", "andrej.estagio": "Andr\u00e9 Jacinto Rodrigues", "italoc.estagio": "\u00cdtalo Costa Cavalcante", "jhessica.estagio": "Jhessica Isabel Coelho Souza", "lucast.estagio": "Lucas Oliveira Torres Machado", "lucca.estagio": "Lucca Lopes de Medeiros", "lukasa.estagio": "Lukas Ara\u00fajo da Silva", "victorc.estagio": "Victor Andrade Cavalcante"}
# import json

# #ABRE DICIONARIO
# with open('usuarios/usuarios.json', 'r') as arquivo:
#     usuarios = json.load(arquivo)
#     print(usuarios)

# # Salvando o dicionário em um arquivo JSON
# with open('usuarios/usuarios.json', 'w') as arquivo:
#     json.dump(usuarios, arquivo)

# try: 
#     nomeEstag = usuarios[user_name]
#     print(nomeEstag)
# except:
#     nomeEstag = str(input("Como é seu primeiro acesso, digite seu nome completo: "))
#     usuarios[user_name] = nomeEstag
#     with open('usuarios/usuarios.json', 'w') as arquivo:
#         json.dump(usuarios, arquivo)
# finally:
#     nomeEstag_sem_acento = unidecode.unidecode(nomeEstag)

# wb = load_workbook(filename = "C:\\Users\\andrej.estagio\\ANATEL\\ORCN - Rádios\\Lista Radiamador.xlsx")
# sheet_ranges = pd.DataFrame(wb['CONFORMES'].values)
# print(sheet_ranges)

# Abre um arquivo Excel e carrega uma página específica como DataFrame
file_path = "C:\\Users\\andrej.estagio\\ANATEL\\ORCN - Rádios\\Lista Radiamador.xlsx"

# Para carregar uma página específica
df = pd.read_excel(file_path, sheet_name='CONFORMES', engine='openpyxl')

# Mostra as primeiras linhas do DataFrame
print(df.head())
