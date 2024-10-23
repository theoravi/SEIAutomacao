
import pandas as pd
# import pyautogui
# from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
# import undetected_chromedriver as uc
# import time
from tabulate import tabulate
import numpy as np
import funcoes as fc
from Levenshtein import distance as lv

# # Criação do DataFrame
preenchido = {'Modelo': ["YAESU FTM-7250DR ", "controle: rm330"], 'Nome Comercial': ["MINI 3", "C5"], 'Número de Série (incluindo rádio controle e óculos)': ["1581F6ZFF564GFDRA45", "5HAZ54GDGDG6"]}
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
    tabela = fc.corrige_planilha(planilhaDrones, drones=False)
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
            if tabela_modelos[j].lower().replace(' ','').replace('-', '') in modelos[modelo_solicitante].lower().replace(' ','').strip('\n').replace('-', ''):
                checkexcel[modelo_solicitante] = True
                print('O modelo', modelos[modelo_solicitante], 'está na planilha de drones conformes')
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

# # Abre um arquivo Excel e carrega uma página específica como DataFrame
# file_path = "C:\\Users\\andrej.estagio\\ANATEL\\ORCN - Rádios\\Lista Radiamador.xlsx"

# # Para carregar uma página específica
# df = pd.read_excel(file_path, sheet_name='CONFORMES', engine='openpyxl')

# import funcoes as fc
# import time
# from selenium.webdriver.common.alert import Alert
# from selenium.webdriver.common.keys import Keys

# navegador = fc.abreChromeEdge()
# while True:
#     #INICIA JANELA
#     try:
#         fc.iniciaJanela(navegador)
#         user_name = fc.user_name
#     #CONDICAO DE ERRO PARA CASO O USUÁRIO ERRE O SEU LOGIN 
#     except Exception:
#         print('Ocorreu um erro, tente novamente.')
#     try:
#         #FECHA ALERTA DO NAVEGADOR
#         time.sleep(1)
#         alert = Alert(navegador)
#         alert.accept()
#     except:
#         #ENTRAR NA CAIXA DE PROCESSOS ATRIBUIDOS AO USUARIO
#         if fc.check_element_exists(By.XPATH,'//*[@id="divFiltro"]/div[2]/a', navegador):
#             break
#         #CASO NAO ENCONTRE O FILTRO DE PROCESSOS ELE INFORMA QUE NAO ENTROU NO SEI
#         else:
#             #APAGA O USUÁRIO
#             navegador.find_element(By.XPATH,'//*[@id="txtUsuario"]').clear()
#             print('Usuário ou senha incorretos. Digite novamente')

# time.sleep(5)

# if not fc.check_element_exists(By.XPATH, 'chirlene.colab' , navegador):
#     navegador.find_element(By.XPATH, '//*[@id="lnkRecebidosProximaPaginaSuperior"]').click()
#     fc.clica_noelemento(navegador, By.PARTIAL_LINK_TEXT, 'chirlene.colab')
#     print('clicado no elemento //*[@id="lnkRecebidosProximaPaginaSuperior"]')
# else:
#     fc.clica_noelemento(navegador, By.PARTIAL_LINK_TEXT, 'chirlene.colab')


# def teste(navegador):
#     processo = str(input("Digite o processo: "))
#     navegador.find_element(By.ID, 'txtPesquisaRapida').send_keys(processo)
#     elementos = navegador.find_element(By.ID, 'txtPesquisaRapida')
#     elementos.send_keys(Keys.ENTER)
#     time.sleep(1)
#     #INSERE TAG REFERENTE A PROCESSOS INTERCORRENTES
#     navegador.switch_to.frame('ifrVisualizacao')
#     #CLICA NO ICONE DE TAG
#     time.sleep(0.5)
#     navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[25]').click()
#     try:    
#         time.sleep(1)
#         navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
#     except:
#         fc.clica_noelemento(navegador, By.XPATH,'//*[@id="tblMarcadores"]/tbody/tr[2]/td[6]/a[2]/img')
#         #navegador.find_element(By.XPATH, '//*[@id="tblMarcadores"]/tbody/tr[2]/td[6]/a[2]/img').click()
#         alert.accept()
#         time.sleep(0.3)
#         fc.clica_noelemento(navegador, By.XPATH,'//*[@id="btnAdicionar"]')
#         input("perai ")
#         #navegador.find_element(By.XPATH, '//*[@id="btnAdicionar"]').click()
#         #time.sleep(0.5)
#         fc.clica_noelemento(navegador, By.XPATH,'//*[@id="selMarcador"]/div/a')
#         #navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
    
#     time.sleep(0.5)
#     #CLICA NA TAG DE INTERCORRENTE
#     navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/ul/li[18]/a').click()
#     #PEDE O TEXTO DA TAG
#     textoTag = input("Insira o texto da tag: ")
#     #COLOCA O TEXTO NA TAG
#     navegador.find_element(By.XPATH, '//*[@id="txaTexto"]').send_keys(textoTag)
#     #SALVA TAG
#     navegador.find_element(By.XPATH, '//*[@id="sbmSalvar"]').click()
#     #DEFINE SITUACAO DO PROCESSO

#     # #CLICA NO DROPDOWN DE TAG
#     # time.sleep(0.5)
#     # #VERIFICA SE ESTA RETIDO
#     # retido = 'não'
#     # if retido == 'não':
#     #     #CLICA NA TAG DE NAO RETIDO
#     #     navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/ul/li[16]')
#     # else:
#     #     #CLICA NA TAG DE RETIDO
#     #     navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/ul/li[17]')
#     #     time.sleep(0.2)
#     #     navegador.find_element(By.XPATH, '//*[@id="txaTexto"]').send_keys('André Jacinto Rodrigues')
#     # #SALVA TAG
#     # navegador.find_element(By.XPATH, '//*[@id="sbmSalvar"]').click()
#     navegador.switch_to.default_content()

# teste(navegador)
# while True:
#     opcao = str(input("Deseja testar mais algum?: "))
#     if opcao == '1':
#         teste(navegador)
#     elif opcao == '2':
#         navegador.quit()
#         break
#     else: print("Opcao", opcao, "é invalida")

# # from datetime import datetime
# # print("Despachos para Drones aprovados "+datetime.now().strftime('%d/%m/%Y'))
# # print(type(datetime.now().strftime('%d/%m/%Y')))
