
import pandas as pd
from tabulate import tabulate
import numpy as np

preenchido = {'Modelo': ["MT3PD\n", "RM 330"], 'Nome Comercial': ["MINI 3", "C5"], 'Número de Série (incluindo rádio controle e óculos)': ["1581FFGDF564GFDRA45", "5HAZ54GDGDG6"] }
df = pd.DataFrame(preenchido)
modelos = df['Modelo']
modelos = modelos.reset_index(drop=True)
print(df)
print(modelos)
print("\n")
df = df.replace(' ', np.nan)
df = df.dropna(how='all')
data = df.values.tolist()
headers = df.columns.tolist()
print(tabulate(data, headers=headers, tablefmt='pretty'))
print('\n')
tabela = pd.read_excel("C:\\Users\\andrej.estagio\\ANATEL\\ORCN - Drones\\Lista de Drones Anatel_Corrigida.xlsx")
tabela.columns = tabela.iloc[1]
tabela = tabela.iloc[2:]
tabela = tabela.reset_index(drop=True)
tabela = tabela['MODELO']
print(modelos[1])
print(type(modelos[1]))
print(modelos[1].lower())
# checkexcel = modelos.isin(tabela['MODELO'])
# checkexcel = tabela['MODELO'].isin(modelos)
input("Pressione qualquer tecla para continuar")
checkexcel = []
for i in tabela:
    for j in range(len(modelos)):
        if modelos[j].lower() == i.lower():
            checkexcel[j] = True
print(checkexcel)
input()
for i in range(len(checkexcel)):
    if checkexcel[i] == True:
        print(f"O modelo {modelos[i]} está na lista de drones conformes.")
    else:
        print(f"O modelo {modelos[i]} não se encontra na lista de drones conformes")
print('\n')