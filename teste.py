
import pandas as pd
from tabulate import tabulate
import numpy as np
import funcoes as fc

preenchido = {'Modelo': [np.nan, "RM330"], 'Nome Comercial': ["MINI 3", "C5"], 'Número de Série (incluindo rádio controle e óculos)': ["1581FFGDF564GFDRA45", "5HAZ54GDGDG6"] }
df = pd.DataFrame(preenchido)
modelos = df['Modelo']
modelos = modelos.reset_index(drop=True)
modelos = modelos.dropna()
print(df)
print(modelos)
print("\n")
df = df.replace(' ', np.nan)
df = df.dropna(how='all')
data = df.values.tolist()
headers = df.columns.tolist()
print(tabulate(data, headers=headers, tablefmt='pretty'))
print('\n')
planilhaDrones = input("Insira o caminho da planilha/lista de drones conformes: ")
#RETIRA ASPAS CASO EXISTA NO CAMINHO DA PLANILHA
planilhaDrones = planilhaDrones.replace('"' , '')
tabela = fc.corrige_planilha(planilhaDrones)
tabela = tabela['MODELO']
# checkexcel = modelos.isin(tabela['MODELO'])
# checkexcel = tabela['MODELO'].isin(modelos)
    
checkexcel = []
for i in modelos:
    checkexcel.append(0)
for modelo_solicitante in range(len(modelos)):
    j = 0
    while j < len(tabela):
        if modelos[modelo_solicitante].lower() == str(tabela[j]).lower():
            checkexcel[modelo_solicitante] = True
            break
        else:
            checkexcel[modelo_solicitante] = False
            j+=1

# for i in tabela:
#     for j in range(len(modelos)):
#         if modelos[j].lower() == str(i).lower():
#             checkexcel[j] = True
#             print("modelo é igual")
#         else:
#             if not checkexcel[j]:
#                 checkexcel[j] = False
# for i in range(len(tabela)):
#     tabela[i] = str(tabela[i])
#     tabela[i] = tabela[i].lower()
    

# for modelo in modelos:
#     if modelo.lower() in tabela:
#         print(f"O modelo {modelo} está na lista de drones conformes.")
#     else:
#         print(f"O modelo {modelo} não se encontra na lista de drones conformes")


            
print(checkexcel)
input()
for i in range(len(checkexcel)):
    if checkexcel[i]:
        print(f"O modelo {modelos[i]} está na lista de drones conformes.")
    else:
        print(f"O modelo {modelos[i]} não se encontra na lista de drones conformes")
print('\n')