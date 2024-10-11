
# import pandas as pd
import pyautogui
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import undetected_chromedriver as uc
# import time
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

# # Abre um arquivo Excel e carrega uma página específica como DataFrame
# file_path = "C:\\Users\\andrej.estagio\\ANATEL\\ORCN - Rádios\\Lista Radiamador.xlsx"

# # Para carregar uma página específica
# df = pd.read_excel(file_path, sheet_name='CONFORMES', engine='openpyxl')

textoRetido = '''Prezado(a) Senhor(a),

Em atenção ao pedido de homologação constante do processo SEI em referência, informamos que:

1. O pedido foi APROVADO
2. O Despacho Decisório que aprovou o pedido está disponível publicamente por meio do sistema SEI na área de Pesquisa Pública, no link:

    https://sei.anatel.gov.br/sei/modulos/pesquisa/md_pesq_processo_pesquisar.php?acao_externa=protocolo_pesquisar&acao_origem_externa=protocolo_pesquisar&id_orgao_acesso_externo=0


3. O Despacho Decisório deverá ser portado junto ao Equipamento (fisicamente ou eletronicamente), para que as autoridades competentes possam conferir a regularidade, quando necessário.
4. Visto que o produto se encontra retido, para que seja informado o número de série, solicitamos que acesse o sistema SEI da Anatel por meio do seguinte link:


            https://sei.anatel.gov.br/sei/controlador_externo.php?acao=usuario_externo_logar&id_orgao_acesso_externo=0.


5. Após autenticação no sistema SEI, solicitamos que inclua os documentos selecionando a opção "Intercorrente" e informando o número do processo em referência.

6. Em caso de equipamento retido, recomendamos que apresente cópia do Despacho decisório ao e-mail corporativo, para que seja liberada a entrega da encomenda retida, de acordo com o local onde está sendo feita a fiscalização.

Caso encomenda retida no Paraná, encaminhar email para - documentacao.pr@anatel.gov.br
Caso encomenda retida em São Paulo, encaminhar email para - documentacao.sp@anatel.gov.br
Caso encomenda retida em Rio de Janeiro, encaminhar email para - documentacao.rj@anatel.gov.br


FAVOR NÃO RESPONDER ESTE E-MAIL.

Atenciosamente,

ORCN - Gerência de Certificação e Numeração

SOR - Superintendência de Outorga e Recursos à Prestação

Anatel - Agência Nacional de Telecomunicações'''

def abreChromeEdge():
    #INSTALA O CHROME DRIVEr MAIS ATUALIZADO
    servico = Service(ChromeDriverManager().install())

    #DEFINE O TEMPO DE EXECUÇÃO PARA CADA COMANDO DO PYAUTOGUI
    pyautogui.PAUSE = 0.7
    
    #INICIA O NAVEGADOR
    navegador = uc.Chrome(service=servico)
    navegador.maximize_window()
    
    #ENTRA NO SEI
    navegador.get('https://dontpad.com/adr1n')
    
    #LOCALIZA O ICONE DO EDGE E ABRE A PLANILHA GERAL
    #FOI UTILIZADO O PYPERCLIP PARA EVITAR QUALQUER ERRO NA HORA DE COLAR O URL DA PLANILHA
    return navegador


navegador = abreChromeEdge()

navegador.find_element(By.XPATH, '//*[@id="text"]').send_keys(textoRetido)
