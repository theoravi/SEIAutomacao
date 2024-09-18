
import pandas as pd
from tabulate import tabulate
import numpy as np
import funcoes as fc
import pyautogui
import pyperclip
import time

# Criação do DataFrame
preenchido = {'Modelo': [np.nan, "RM 330"], 'Nome Comercial': ["MINI 3", "C5"], 'Número de Série (incluindo rádio controle e óculos)': ["1581FFGDF564GFDRA45", "5HAZ54GDGDG6"]}
df = pd.DataFrame(preenchido)
modelos = df['Modelo'].dropna().reset_index(drop=True)

print(df)
print(modelos)
print("\n")

# Exibe a tabela formatada
df = df.replace(' ', np.nan).dropna(how='all')
data = df.values.tolist()
headers = df.columns.tolist()
print(tabulate(data, headers=headers, tablefmt='pretty'))
print('\n')

planilhaDrones = input("Insira o caminho da planilha/lista de drones conformes: ").replace('"' , '')

# Tenta carregar a planilha de drones conformes
try:
    tabela = fc.corrige_planilha(planilhaDrones)
    tabela_modelos = tabela['MODELO'].astype(str)  # Garante que os modelos estão como strings
except Exception as e:
    print(f"Erro ao carregar a planilha: {e}")
    exit()

# checkexcel = modelos.isin(tabela['MODELO'])
# checkexcel = tabela['MODELO'].isin(modelos)
    
# Verificação de conformidade
checkexcel = [False] * len(modelos)  # Inicializa a lista com False

# for modelo_solicitante in range(len(modelos)):
#     j = 0
#     while j < len(tabela):
#         if modelos[modelo_solicitante].lower() == str(tabela[j]).lower():
#             checkexcel[modelo_solicitante] = True
#             break
#         else:
#             checkexcel[modelo_solicitante] = False
#             j+=1

# Comparação de modelos
for modelo_solicitante in range(len(modelos)):
    for j in range(len(tabela_modelos)):
        try:
            if modelos[modelo_solicitante].lower().replace(' ','').strip('\n') == tabela_modelos[j].lower().replace(' ',''):
                checkexcel[modelo_solicitante] = True
                break
        except KeyError as e:
            print(f"Erro ao acessar os índices: {e}")
            continue

# Exibe o resultado
print("Resultados da verificação de conformidade:", checkexcel)
            
print(checkexcel)
input()
for i in range(len(checkexcel)):
    if checkexcel[i]:
        print(f"O modelo {modelos[i]} está na lista de drones conformes.")
    else:
        print(f"O modelo {modelos[i]} não se encontra na lista de drones conformes")
print('\n')

