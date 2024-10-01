
import pandas as pd
from tabulate import tabulate
import numpy as np
import funcoes as fc
from Levenshtein import distance as lv

# # Criação do DataFrame
preenchido = {'Modelo': ["uv k5", "dm 701", "uv-k5"], 'Nome Comercial': ["MINI 3", "C5", ' '], 'Número de Série (incluindo rádio controle e óculos)': ["1581FFGDF564GFDRA45", "5HAZ54GDGDG6", " "]}
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

planilhaDrones = input("Insira o caminho da planilha/lista de rádios conformes: ").replace('"' , '')

# Tenta carregar a planilha de drones conformes
try:
    tabela = fc.corrige_planilha(planilhaDrones)
    tabela_modelos = tabela.astype(str)  # Garante que os modelos estão como strings
except Exception as e:
    print(f"Erro ao carregar a planilha: {e}")
    exit()
    
# Verificação de conformidade
checkexcel = [False] * len(modelos)  # Inicializa a lista com False
lista_parecidos = []
# Comparação de modelos
for modelo_solicitante in range(len(modelos)):
    for j in range(len(tabela_modelos)):
        try:
            if modelos[modelo_solicitante].lower().replace(' ','').strip('\n').replace('-', '') == tabela_modelos[j].lower().replace(' ','').replace('-', ''):
                checkexcel[modelo_solicitante] = True
                print('O modelo', modelos[modelo_solicitante], 'bate com o modelo', tabela_modelos[j], 'da tabela de rádios conformes')
                break
            elif lv(modelos[modelo_solicitante].lower().replace(' ','').strip('\n').replace('-', ''), tabela_modelos[j].lower().replace(' ','').replace('-', '')) < 2:
                lista_parecidos.append(tabela_modelos[j])
        except KeyError as e:
            print(f"Erro ao acessar os índices: {e}")
            continue
for i in range(len(checkexcel)):
    if not checkexcel[i]:
        print(f"O modelo {modelos[i]} não se encontra na lista de rádios conformes\n", "Lista de modelos parecidos na tabela:", lista_parecidos)
print('\n')