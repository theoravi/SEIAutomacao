#IMPORTS NECESSÁRIOS PARA O FUNCIONAMENTO DO CÓDIGO
import numpy as np
import os
import pandas as pd
import pyautogui
import pyperclip
import time
import tkinter as tk
import undetected_chromedriver as uc
import unidecode
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from tabulate import tabulate
from tkinter import scrolledtext
from webdriver_manager.chrome import ChromeDriverManager

#FUNCAO QUE RECEBE USUARIO E SENHA NA INTERFACE GRAFICA DO TKINTER E FECHA A JANELA ABERTA
def preencher_campos():
    user = user_var.get()
    senha = senha_var.get()
    
    navegador.find_element(By.XPATH,'//*[@id="txtUsuario"]').clear()
    navegador.find_element(By.XPATH,'//*[@id="pwdSenha"]').clear()
    navegador.find_element(By.XPATH,'//*[@id="txtUsuario"]').send_keys(f"{user}")
    navegador.find_element(By.XPATH,'//*[@id="pwdSenha"]').send_keys(f"{senha}")
    root.destroy()

#FUNÇÃO PARA VERIFICAR A EXISTÊNCIA DE ELEMENTOS NA TELA UTILIZANDO O SELENIUM
def check_element_exists(by, value):
    try:
        navegador.find_element(by, value)
        return True
    except NoSuchElementException:
        return False
    
#FUNÇÃO QUE FAZ O LOG EM UM TXT COM O NOME DO USUÁRIO   
def escrever_informacoes(processos, nomeEstag):
    #OBTER DATA E HORA DO INÍCIO DA ANÁLISE DO PROCESSO
    data_hora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    
    #NOME DO ARQUIVO BASEADO NO USUÁRIO QUE ESTÁ ANALISANDO
    nome_arquivo = f"{nomeEstag}.txt"
    
    #VERIFICA SE O ARQUIVO EXISTE
    if not os.path.exists(nome_arquivo):
        #SE NAO EXISTIR ELE CRIA O ARQUIVO FORNECIDO NO INÍCIO DA ANÁLISE
        with open(nome_arquivo, 'w') as file:
            file.write(f'N do processo {processos} | Estagiario responsavel {nomeEstag} | Data e hora da analise {data_hora}\n')
    else:
        #CASO O ARQUIVO EXISTA ELE APENAS ADICIONA AS NOVAS INFORMACOES
        with open(nome_arquivo, 'a') as file:
            file.write(f'N do processo {processos} | Estagiario responsavel {nomeEstag} | Data e hora da analise {data_hora}\n')
 
def processa_tr(tr):
    #PEGA TODAS AS COLUNAS (TD) DA LINHA ATUAL (TR)
    tds = tr.find_elements(By.TAG_NAME, 'td')
    #INICIALIZA AS VARIÁVEIS QUE SERÃO USADAS PARA ARMAZENAR INFORMAÇÕES SOBRE O PROCESSO
    processo_texto = ""
    possui_anotacao = False
    aguardando_assinatura = False
    #ITERA SOBRE CADA COLUNA (TD) DA LINHA ATUAL (TR)
    for td in tds:
        #PEGA O TEXTO DA COLUNA ATUAL (TD)
        texto = td.text
        #VERIFICA SE O TEXTO NÃO CONTÉM A PALAVRA '.ESTAGIO'
        if '.estagio' not in texto:
            #REMOVE AS PALAVRAS 'RECEBIDOS' E ESPAÇOS EM BRANCO DO TEXTO
            texto = texto.replace('Recebidos', '').replace(' ', '')
            #ARMAZENA O TEXTO MODIFICADO NA VARIÁVEL 'PROCESSO_TEXTO'
            processo_texto = texto
        #PROCURA POR ÍCONES DE ANOTAÇÃO NA LINHA ATUAL (TR)
        anotacoes = tr.find_elements(By.XPATH, './/img[@src="svg/anotacao1.svg?11"]')
        #VERIFICA SE ENCONTROU ALGUM ÍCONE DE ANOTAÇÃO
        if anotacoes:
            #SE ENCONTROU, MARCA A VARIÁVEL 'POSSUI_ANOTACAO' COMO VERDADEIRA
            possui_anotacao = True
            #VERIFICA SE A ANOTAÇÃO CONTÉM O TEXTO 'AGUARDANDO ASSINATURA'
            if tr.find_elements(By.XPATH, ".//*[contains(translate(@onmouseover, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'aguardando assinatura')]"):
                #SE CONTÉM, MARCA A VARIÁVEL 'AGUARDANDO_ASSINATURA' COMO VERDADEIRA
                aguardando_assinatura = True
    #RETORNA AS INFORMAÇÕES DO PROCESSO COMO UMA TUPLA
    return (processo_texto.strip(), possui_anotacao, aguardando_assinatura)
  
#FUNCAO PARA EXIGENCIA DE INSERIR ANEXOS
def insira_anexo(processos):
    def get_text():
        #ENVIA TEXTO DO INPUT
        nonlocal texto_exigencia
        texto_exigencia = text_area.get("1.0", tk.END).strip()
        #FECHA A JANELA DE INPUT
        root.destroy()

    #INICIALIZA A VARIAVEL LOCAL QUE RECEBERA O TEXTO DE INPUT
    texto_exigencia = ""

    #CRIA A JANELA DA INTERFACE PARA CAIXA DE INPUT
    root = tk.Tk()
    root.title("Inserir Texto de Exigência")

    #INSTRUCAO DO QUE O USUARIO DEVE FAZER
    label_instruction = tk.Label(root, text="Insira apenas a exigência:")
    label_instruction.pack(padx=10, pady=5)

    #CRIA AREA DE TEXTO COM ROLAGEM
    text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=20)
    text_area.pack(padx=10, pady=10)

    #CRIA BOTAO DE RECEBER O TEXTO
    btn_get_text = tk.Button(root, text="Obter Texto", command=get_text)
    btn_get_text.pack(pady=5)

    #INICIA O LOOP PRINCIPAL QUE MANTEM A JANELA ATIVA
    root.mainloop()
    
    #PREENCHE O ASSUNTO DO EMAIL
    navegador.find_element(By.ID, 'txtAssunto').send_keys(f"Processo SEI nº {processos} - Exigência")
    #PREENCHE O CAMPO DE EMAIL COM A VARIAVEL CRIADA ACIMA
    navegador.find_element(By.ID, 'txaMensagem').send_keys(f'''Prezado(a) Senhor(a),

    Em atenção à demanda registrada no processo em referência, apresentamos as seguintes pendências observadas durante a análise da solicitação:

    {texto_exigencia}

    Favor seguir as seguintes orientações:
    
    1. Solicitamos acessar o sistema Sei da Anatel: https://sei.anatel.gov.br/sei/controlador_externo.php?acao=usuario_externo_logar&id_orgao_acesso_externo=0    

    2. Selecionar o processo e clicar em Peticionamento Intercorrente;

    3. Selecionar os arquivos que deseja incluir para atender à(s) exigência(s), selecionar o Tipo do Documento, Nível de Acesso e Formato. Após isso clique em Adicionar.

    4. Clique em Peticionar.

    5. Indicar o Cargo/Função e senha de acesso ao SEI. Clique em Assinar.

    FAVOR NÃO RESPONDER ESTE E-MAIL.

    Atenciosamente,

    ORCN - Gerência de Certificação e Numeração

    SOR - Superintendência de Outorga e Recursos à Prestação

    Anatel - Agência Nacional de Telecomunicações''')

#FUNCAO PARA ERROS NA DECLARACAO DE CONFORMIDADE PEDINDO PARA ABRIR NOVO PROCESSO SEI
def erro_declaracao(processos, impProp):
    #ENVIA TEXTO DE INPUT PARA A VARIAVEL
    def get_text():
        nonlocal exig2
        exig2 = text_area.get("1.0", tk.END).strip()
        root.destroy()

    #INICIALIZA A VARIAVEL LOCAL QUE RECEBERA O TEXTO DE INPUT
    exig2 = ""

    #CRIA A JANELA DA INTERFACE PARA CAIXA DE INPUT
    root = tk.Tk()
    root.title("Inserir Texto de Exigência")

    #INSTRUCAO DO QUE O USUARIO DEVE FAZER
    label_instruction = tk.Label(root, text="Insira apenas a exigência:")
    label_instruction.pack(padx=10, pady=5)

    #CRIA AREA DE TEXTO COM ROLAGEM
    text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=20)
    text_area.pack(padx=10, pady=10)

    #CRIA BOTAO DE RECEBER O TEXTO
    btn_get_text = tk.Button(root, text="Obter Texto", command=get_text)
    btn_get_text.pack(pady=5)

    #INICIA O LOOP PRINCIPAL QUE MANTEM A JANELA ATIVA
    root.mainloop()
    #VERIFICA SE É IMPORTADO PRA USO PRÓPRIO
    #CASO SEJA IMPORTADO PRA USO PROPRIO E SEJA UM RADIO ELE PASSARÁ O MANUAL PARA RADIO E CORRELATOS
    #SE FOR DRONE PRA USO PROPRIO ELE PASSARA O MANUAL PARA DRONE
    
    if impProp == 'Sim':
        while True:
            try:
                radioamador = int(input('O produto é um rádio ou correlatos? Se sim digite [1], se for um Drone para uso próprio digite [2]: '))
                
                if radioamador == 1:
                    link = 'https://docs.google.com/document/d/1Y-8wGVnDMuku_Dmd5uT6hHCuj3ruXbuW3o-JFDWxog0/edit'
                    break
                elif radioamador == 2:
                    link = 'https://docs.google.com/document/d/1YsrMwMxyVysCJ9VhtGOOU1S1FlbmkRFaYSjojLxXVsQ/edit'
                    break
                else:
                    print("Opção inválida. Por favor, digite [1] para rádio ou correlatos, ou [2] para Drone para uso próprio.")
            except ValueError:
                print("Entrada inválida. Por favor, digite um número [1] para rádio ou correlatos, ou [2] para Drone para uso próprio.")
    
    #CASO SEJA UMA DECLARACAO APENAS PARA DRONE ELE IRÁ PASSAR DIRETAMENTE O MANUAL PARA DRONE
    else:
        link = 'https://docs.google.com/document/d/1YsrMwMxyVysCJ9VhtGOOU1S1FlbmkRFaYSjojLxXVsQ/edit'
    
    #INSERE ASSUNTO DO PROCESO
    navegador.find_element(By.ID, 'txtAssunto').send_keys(f"Processo SEI nº {processos} - Indeferido")
    #INSERE EMAIL COM CORPO DO EMAIL
    navegador.find_element(By.ID, 'txaMensagem').send_keys(f'''Prezado(a) Senhor(a),

    Em atenção à demanda registrada no processo em referência, apresentamos as seguintes pendências observadas no documento de Declaração de Conformidade:

    {exig2}

    Favor seguir as seguintes orientações:
    
    1. Realizar um novo processo SEI, cumprindo com a(s) exigência(s) apontada(s) anteriormente. Seguindo o manual: {link}

    2. Para criar um novo processo SEI, solicitamos acessar o sistema Sei da Anatel: https://sei.anatel.gov.br/sei/controlador_externo.php?acao=usuario_externo_logar&id_orgao_acesso_externo=0

    Dessa maneira, devido a presença de exigência(s), informamos que o atual processo SEI nº {processos} será arquivado.

    FAVOR NÃO RESPONDER ESTE E-MAIL.



    Atenciosamente,

    ORCN - Gerência de Certificação e Numeração

    SOR - Superintendência de Outorga e Recursos à Prestação

    Anatel - Agência Nacional de Telecomunicações''')

#FUNCAO QUE PERMITE O USUARIO INSERIR CORPO DO EMAIL
def outro_erro():
    #INSERE SOMENTE ASSUNTO DO EMAIL
    assunto  = input("Insira o assunto do email: ")
    navegador.find_element(By.ID, 'txtAssunto').send_keys(assunto)

#FUNCAO QUE ANALISA PROCESSO E CRIA DESPACHO DECISORIO OU GERA EXIGENCIA
def analisaProcesso():
    #COLETA NOME DO ESTAGIARIO
    #COLETA CAMINHO DA PLANILHA DE DRONES CONFORMES
    planilhaDrones = input("Insira o caminho da planilha/lista de drones conformes: ")
    #RETIRA ASPAS CASO EXISTA NO CAMINHO DA PLANILHA
    planilhaDrones = planilhaDrones.replace('"' , '')
    #LE A PLANILHA DE DRONES PARA USAR COMO CONSULTA NOS PROCESSOS
    #TRATA A TABELA FORNECIDA PELO PANDAS
    tabela = pd.read_excel(planilhaDrones)
    tabela.columns = tabela.iloc[1]
    tabela = tabela.iloc[2:]
    tabela = tabela.reset_index(drop=True)

    #VARIAVEL UTILIZADA PARA O SELENIUM RETORNAR PARA A JANELA PRINCIPAL
    janela_principal = navegador.current_window_handle
    
    #ITERA PELOS PROCESSOS CONTIDOS NA LISTA DE PROCESSOS PARA ANALISE
    for processos in lista_processos[:]:
        #GERA UMA CONDICAO DE SAIDA AO USUARIO PARA CASO QUEIRA PARA DE ANALISAR
        while True:
            try:
                analisar = int(input("Deseja analisar o próximo processo? Se sim digite [1], se não digite [2]: "))
                #SAI DA ANALISE CASO SEJA DIGITADO O 2
                if analisar == 2:
                    return
                #INICIA A ANALISE CASO SEJA DIGITADO 1
                elif analisar==1:
                    try:
                        #FAZ UM LOG DO PROCESSO
                        escrever_informacoes(processos, nomeEstag)
                        #INICIA A VARIAVEL ZERADA
                        impProp=''
                        situacao=''
                        #PESQUISA NUMERO DO PROCESSO NA CAIXA DE PESQUISA DO SEI PARA ACESSAR O PROCESSO
                        navegador.switch_to.default_content()
                        navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processos) 
                        elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
                        elementos.send_keys(Keys.ENTER)    
                        #TEMPO DEFINIDO PARA QUE A PAGINA CARREGUE
                        time.sleep(1)
                        #ENTRA NO FRAME QUE CONTÉM OS DOCUMENTOS DOS PROCESSOS
                        #A PAGINA DOS PROCESSOS SAO DIVIDIDOS EM FRAMES
                        navegador.switch_to.frame('ifrArvore')
                        #VERIFICA SE O PROCESSO JA FOI DESPACHADO, CASO TENHA SIDO ELE PULA O PROCESSO
                        if check_element_exists(By.PARTIAL_LINK_TEXT, 'Despacho Decisório'):
                            print(f"O processo {processos} já foi despachado!")
                        else:
                            #VERIFICA SE HÁ PASTAS NO PROCESSO, SE HOUVER ELE ABRE A PRIMEIRA PARA ENCONTRAR A DECLARACAO DE CONFORMIDADE
                            if check_element_exists(By.XPATH, '//*[@id="spanPASTA1"]'):
                                navegador.find_element(By.XPATH, '//*[@id="spanPASTA1"]').click()
                                time.sleep(1)

                            #ENCONTRAR DECLARAÇÃO DE CONFORMIDADE NO PROCESSO DRONE
                            if check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Drone"):
                                #ACESSA O RECIBO PARA PEGAR NOME DO SOLICITANTE
                                navegador.find_element(By.PARTIAL_LINK_TEXT, 'Recibo Eletrônico').click()
                                navegador.switch_to.default_content()
                                #ENTRA NOS FRAMES QUE POSSUE O NOME DO SOLICITANTE
                                navegador.switch_to.frame('ifrVisualizacao')
                                navegador.switch_to.frame('ifrArvoreHtml')
                                #ARMAZENA NOME DO SOLICITANTE NA VARIAVEL
                                nomeSol = navegador.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[1]/td[2]').text
                                #RETORNA AO FRAME DEFAULT
                                navegador.switch_to.default_content()
                                #COLETAR DADOS DA DECLARAÇÃO DE CONFORMIDADE
                                #RETORNA AO FRAME QUE CONTÉM OS DOCUMENTOS
                                navegador.switch_to.frame('ifrArvore')
                                #CLICA NA DECLARACAO DE CONFORMIDADE
                                navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Drone").click()
                                time.sleep(1)
                                #RETORNA AO FRAME DEFAULT E ENTRA NOS FRAMES QUE CONTÉM A VISUALIZACAO DA DECLARACAO
                                navegador.switch_to.default_content()
                                navegador.switch_to.frame('ifrVisualizacao')
                                navegador.switch_to.frame('ifrArvoreHtml')
                                #EXIBE INFORMACOES DO PROCESSO
                                print("--------------------------------------------------------------------------")
                                print('--------------------------------------------------------------------------')
                                print(f"Processo: {processos}")

                                #COLETAR TEXTO NOME DO INTERESSADO
                                nome_interessado = navegador.find_element(By.XPATH, '/html/body/table[1]/tbody/tr/td').text
                                #VERIFICA SE O CAMPO ESTÁ VAZIO, CASO ESTEJA DEFINIRA CM NAO INFORMADO
                                if not nome_interessado.strip():
                                    nome_interessado = 'Não informado'
                                    print("NOME DO INTERESSADO NÃO INFORMADO")
                                #EXIBE NOMES DO INTERESSADO E SOLICITANTE
                                print("Interessado: ",nome_interessado)
                                print("Solicitante: ",nomeSol)
                                print('\n')

                                #CHECKBOX DE QUEM ESTÁ FAZENDO O PETICIONAMENTO
                                checkbox=navegador.find_element(By.XPATH,'/html/body/table[2]/tbody/tr[1]/td[1]').text
                                checkbox2=navegador.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[2]/td[1]').text
                                #VERIFICA CAMPO DA CHECKBOX
                                #SE OS DOIS CAMPOS ESTIVEREM VAZIOS AVISA QUE NAO FOI INFORMADO QUEM ESTÁ FAZENDO O PETICIONAMENTO
                                if not checkbox.strip() and not checkbox2.strip():
                                    print("O usuário não informou quem está fazendo o peticionamento, confira os nomes do documento.")
                                elif not checkbox.strip():
                                    procuracao_eletronica = navegador.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[2]/td[3]').text
                                    print("Peticionamento representando terceiro, Pessoa Jurídica ou Pessoa Física.", f"\nProcuração eletrônica: {procuracao_eletronica}.")
                                else:
                                    print("Peticionamento em interesse próprio como Pessoa Física.")
                                print('\n')

                                #CONTA QUANTIDADE DE LINHAS DA TABELA SOBRE O PRODUTO
                                quantidade_linhas = len(navegador.find_elements(By.XPATH, '/html/body/table[3]/tbody/tr'))

                                #VERIFICA CAMPOS QUE CONTÉM AS INFORMAÇÕES SOBRE O PRODUTO
                                #ARMAZENA OS MODELOS INFORMADOS PARA VERIFICAR SE ESTÃO NA PLANILHA DE DRONES CONFORMES
                                modelos = []
                                for i in range(1, quantidade_linhas, 1):
                                    linha = navegador.find_element('xpath', f'/html/body/table[3]/tbody/tr[{i}]')
                                    celulas = linha.find_elements(By.TAG_NAME,'td')
                                    modelos_produto = [celula.text for celula in celulas]
                                    modelos.append(modelos_produto)

                                #CRIA DATAFRAME PARA PRINTAR AS INFORMAÇÕES DO PRODUTO JÁ TRATADAS
                                df = pd.DataFrame(modelos)
                                df.columns = ['Modelo', 'Nome Comercial', 'Número de Série (incluindo rádio controle e óculos)']
                                df = df[1:]
                                df = df.replace(' ', np.nan)
                                df = df.dropna(how='all')
                                data = df.values.tolist()
                                headers = df.columns.tolist()
                                print(tabulate(data, headers=headers, tablefmt='pretty'))
                                print('\n')

                                #VERIFICA SE O MODELO DO DRONE E RADIO CONTROLE ESTA NA PLANILHA DE DRONES CONFORMES
                                modelos = df['Modelo']
                                modelos = modelos.reset_index(drop=True)
                                checkexcel = modelos.isin(tabela['MODELO'])
                                for i in range(len(checkexcel)):
                                    if checkexcel[i] == True:
                                        print(f"O modelo {modelos[i]} está na lista de drones conformes.")
                                    else:
                                        print(f"O modelo {modelos[i]} não se encontra na lista de drones conformes")
                                print('\n')
                                #EXIBE CÓDIGO DE RASTREIO
                                #LINHA ESPECIFICA ESCREVE NA PLANILHA SE O PRODUTO ESTA RETIDO E, CASO ESTEJA, ESCREVE O CODIGO DE RASTREIO
                                codigo_rastreio = navegador.find_element(By.XPATH,'/html/body/table[4]/tbody/tr/td[2]').text
                                #VERIFICA SE HÁ ALGUM TEXTO NA TABELA COM O CÓDIGO DE RASTREIO
                                #SE HOUVER ALGUM TEXTO ELE DA COMO RETIDO, CASO NAO TENHA TEXTO ELE DA COMO NAO RETIDO
                                if not codigo_rastreio.strip():
                                    print("Produto não retido ou código de rastreio não informado.")
                                    retido='Não'
                                else:
                                    print(f'O código de rastreio é: {codigo_rastreio}.')
                                    retido='Sim' 
                                    
                                print('\n')
                                print('Confira o relatório fotográfico e veja se os documentos estão conformes!\n')
                                print('--------------------------------------------------------------------------')
                                print('--------------------------------------------------------------------------')
                                
                            #ENCONTRAR SE HÁ ALGUMA PASTA DE DOCUMENTOS OU DECLARAÇÃO DE CONFORMIDADE NO PROCESSO
                            elif check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Importado Uso Próprio"):
                                #PEGAR NOME DO SOLICITANTE NO RECIBO ELETRONICO
                                navegador.find_element(By.PARTIAL_LINK_TEXT, 'Recibo Eletrônico').click()
                                navegador.switch_to.default_content()
                                time.sleep(0.5)
                                navegador.switch_to.frame('ifrVisualizacao')
                                navegador.switch_to.frame('ifrArvoreHtml')
                                nomeSol = navegador.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[1]/td[2]').text
                                navegador.switch_to.default_content()
                                #CLICA NA DECLARACAO DE CONFORMIDADE
                                #COLETAR DADOS DA DECLARAÇÃO DE CONFORMIDADE
                                navegador.switch_to.frame('ifrArvore')
                                navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Importado Uso Próprio").click()
                                time.sleep(0.5)
                                navegador.switch_to.default_content()
                                navegador.switch_to.frame('ifrVisualizacao')
                                navegador.switch_to.frame('ifrArvoreHtml')

                                #EXIBE INFORMACOES DO PROCESSO
                                print("--------------------------------------------------------------------------")
                                print('--------------------------------------------------------------------------')
                                print(f"Processo: {processos}")

                                #COLETAR TEXTO NOME DO INTERESSADO
                                nome_interessado=navegador.find_element(By.XPATH, '/html/body/table[1]/tbody/tr/td/p').text
                                if not nome_interessado.strip():
                                    nome_interessado = 'Não informado'
                                    print("NOME DO INTERESSADO NÃO INFORMADO")
                                print(nome_interessado)
                                print(nomeSol)
                                print('\n')
                                #CHECKBOX DE QUEM ESTÁ FAZENDO O PETICIONAMENTO
                                checkbox=navegador.find_element('xpath','/html/body/table[2]/tbody/tr[1]/td[1]/p').text
                                checkbox2=navegador.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[2]/td[1]').text

                                #CHECA CAMPO DA CHECKBOX
                                if not checkbox.strip() and not checkbox2.strip():
                                    print("O usuário não informou quem está fazendo o peticionamento, confira os nomes do documento.")
                                elif not checkbox.strip():
                                    procuracao_eletronica = navegador.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[2]/td[3]').text
                                    print("Peticionamento representando terceiro, Pessoa Jurídica ou Pessoa Física.", f"\nProcuração eletrônica: {procuracao_eletronica}.")
                                else:
                                    print("Peticionamento em interesse próprio como Pessoa Física.")
                                print('\n')

                                #INFORMA O TIPO DE PRODUTO QUE ESTÁ SENDO TRATADO
                                print("\nTipo de produto:")
                                quantidade_linhasPro = len(navegador.find_elements(By.XPATH, '/html/body/table[3]/tbody/tr'))
                                tipoPro = []
                                for i in range(1, quantidade_linhasPro, 1):
                                    linhas = navegador.find_element(By.XPATH, f'/html/body/table[3]/tbody/tr[{i}]')
                                    celulas2 = linhas.find_elements(By.TAG_NAME,'td')
                                    produtos = [celula2.text for celula2 in celulas2]
                                    tipoPro.append(produtos)
                                dfPro = pd.DataFrame(tipoPro)
                                dfPro.columns = ['Tipo de Equipamento', 'Assinale com um X', 'Observação:']
                                dfPro = dfPro.drop('Observação:', axis=1)
                                dfPro = dfPro[1:]

                                #EXIBE A TABELA DO TIPO DE PRODUTO
                                data = dfPro.values.tolist()
                                headers = dfPro.columns.tolist()
                                print(tabulate(data, headers=headers, tablefmt='pretty'))
                                print('\n')

                                #CONTA QUANTIDADE DE LINHAS DA TABELA SOBRE O PRODUTO
                                quantidade_linhas = len(navegador.find_elements(By.XPATH, '/html/body/table[4]/tbody/tr'))

                                #VERIFICA CAMPOS QUE CONTÉM AS INFORMAÇÕES SOBRE O PRODUTO
                                modelos = []
                                for i in range(1, quantidade_linhas, 1):
                                    linha = navegador.find_element('xpath', f'/html/body/table[4]/tbody/tr[{i}]')
                                    celulas = linha.find_elements(By.TAG_NAME,'td')
                                    modelos_produto = [celula.text for celula in celulas]
                                    modelos.append(modelos_produto)
                                
                                #CRIA DATAFRAME E VERIFICA SE OS MODELOS FORNECIDOS ESTÃO NA PLANILHA DE DRONES CONFORMES
                                df = pd.DataFrame(modelos)
                                df.columns = ['Modelo', 'Nome Comercial', 'Número de Série']
                                df = df[1:]
                                df = df.replace(' ', np.nan)
                                df = df.dropna(how='all')
                                modelos = df['Modelo']
                                modelos = modelos.reset_index(drop=True)
                                checkexcel = modelos.isin(tabela['MODELO'])                        
                                print("Modelos:")
                                #CRIA DATAFRAME PARA PRINTAR AS INFORMAÇÕES DO PRODUTO
                                data2 = df.values.tolist()
                                headers2 = df.columns.tolist()
                                print(tabulate(data2, headers=headers2, tablefmt='pretty'))
                                print('\n')

                                for i in range(len(checkexcel)):
                                    if checkexcel[i] == True:
                                        print(f"O modelo {modelos[i]} está na lista de drones conformes.")
                                    else:
                                        print(f"O modelo {modelos[i]} não se encontra na lista de drones conformes")
                                print('\n')
                                #EXIBE CÓDIGO DE RASTREIO
                                #ESCREVE NA TABELA EXCEL SE ESTA RETIDO, CASO ESTEJA INSERE CODIGO DE RASTREIO
                                codigo_rastreio = navegador.find_element(By.XPATH,'/html/body/table[5]/tbody/tr/td[2]').text
                                if not codigo_rastreio.strip():
                                    print("Produto não retido ou código de rastreio não informado.")
                                    retido='Não'
                                else:
                                    print(f'O código de rastreio é: {codigo_rastreio}.')
                                    retido='Sim'
                                print('\n')
                                print('Confira o relatório fotográfico e veja se os documentos estão conformes!\n')
                                print('--------------------------------------------------------------------------')
                                print('--------------------------------------------------------------------------')
                                
                            time.sleep(1)
                            navegador.switch_to.default_content()

                            #CRIA DESPACHO DECISORIO
                            #PEDE PRA USUARIO VERIFICAR SE ESTÁ TUDO CERTO
                            #VOLTA PARA A PAGINA DE INICIO DO PROCESSO
                            while True:
                                try:
                                    confirmacao = int(input('Caso esteja tudo certo e NÃO HAJA UM DESPACHO JÁ CRIADO digite [1] para continuar, senão, digite [2]: '))
                                    if confirmacao == 1:
                                        break
                                    elif confirmacao == 2:
                                        break
                                    elif nome_interessado == 'Não informado' and confirmacao == 1:
                                        print("Nome do interessado não informado, não é possível criar despacho, selecione outra opção!")
                                    else:
                                        print("Opção inválida, tente novamente!")
                                except ValueError:
                                    print("Opção inválida, tente novamente!")
                            if confirmacao == 1:
                                navegador.switch_to.frame('ifrArvore')
                                arvoredocs = navegador.find_elements(By.CLASS_NAME, 'infraArvoreNo')

                                #VERIFICA SE O DOCUMENTO ESTÁ RESTRITO, SE ESTIVER ELE IRÁ DEIXAR COMO PUBLICO UTILIZANDO O SIMBOLO DE RESTRITO COMO REFERENCIA
                                elementos_com_src = navegador.find_elements(By.CSS_SELECTOR, "[src='svg/processo_restrito.svg?11']")
                                #CRIA UMA LISTA PARA OS DOCUMENTOS QUE DEVERÃO SER DEIXADOS COMO PUBLICO
                                elementos_para_clicar = []
                                #ITERA SOBRE OS DOCUMENTOS 
                                for doc in arvoredocs:
                                    #ITERA SOBRE OS ELEMENTOS COM O SIMBOLO DE RESTRITO
                                    for elemento in elementos_com_src:
                                        #VERIFICA SE O DOCUMENTO ESTÁ AO LADO DO SIMBOLO DE RESTRITO
                                        if elemento in doc.find_elements(By.XPATH, 'following-sibling::*[2]/*'):
                                            #SE O SIMBOLO DE RESTRITO ESTIVER NO DOCUMENTO ELE INCLUIRÁ NA LISTA DE DOCUMENTOS A SEREM ACESSADOS
                                            elementos_para_clicar.append(doc.get_attribute("id"))
                                            break

                                #AGORA COM O ID DO DOCUMENTO ELE CLICARÁ EM CADA DOCUMENTO RESTRITO PARA DEIXAR PUBLICO
                                for id in elementos_para_clicar:
                                    #ENCONTRA O ELEMENTO PARA CLICAR
                                    doc = navegador.find_element(By.ID, id)
                                    doc.click()
                                    time.sleep(0.7)
                                    navegador.switch_to.default_content()
                                    navegador.switch_to.frame('ifrVisualizacao')
                                    #CLICA NO SIMBOLO DE ALTERAR DOCUMENTO
                                    navegador.find_element(By.XPATH,'//*[@id="divArvoreAcoes"]/a[2]').click()
                                    #SELECIONA A OPCAO DE PUBLICO NO DOCUMENTO
                                    navegador.find_element(By.XPATH,'//*[@id="divOptPublico"]/div/label').click()
                                    #SALVA AS MUDANCAS
                                    navegador.find_element(By.ID,'btnSalvar').click()
                                    navegador.switch_to.default_content()
                                    navegador.switch_to.frame('ifrArvore')
                                navegador.switch_to.default_content()
                                navegador.switch_to.frame('ifrArvore')
                                #ENTRA NA PAGINA INICIAL DO PROCESSO
                                #VERIFICA SE E UM PROCESSO DE DRONE OU IMPORTADO PARA USO PROPRIO
                                if check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Drone"):
                                    navegador.switch_to.default_content()
                                    navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processos)
                                    elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
                                    elementos.send_keys(Keys.ENTER)
                                    #CRIA DESPACHO
                                    navegador.switch_to.frame('ifrVisualizacao')
                                    time.sleep(1)
                                    #CLICA NO INCONE DE INCLUIR DOCUMENTO
                                    navegador.find_element(By.XPATH,'//*[@id="divArvoreAcoes"]/a[1]').click()
                                    time.sleep(1)
                                    #CLICA NA OPCAO DE DESPACHO DECISORIO
                                    navegador.find_element(By.XPATH,'//*[@id="tblSeries"]/tbody/tr[16]/td/a[2]').click()
                                    #SELECIONA TEXTO PADRAO
                                    navegador.find_element(By.XPATH,'//*[@id="divOptTextoPadrao"]/div').click()
                                    #ENVIA QUAL DESPACHO DECISORIO DEVE SER CRIADO
                                    navegador.find_element(By.XPATH,'//*[@id="txtTextoPadrao"]').send_keys('Despacho Decisório de Homologação Drones')
                                #VERIFICA SE E UM PROCESSO DE DRONE OU IMPORTADO PARA USO PROPRIO
                                elif check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Importado Uso Próprio"):
                                    impProp='Sim'
                                    navegador.switch_to.default_content()
                                    navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processos)
                                    elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
                                    elementos.send_keys(Keys.ENTER)
                                    #CRIA DESPACHO
                                    navegador.switch_to.frame('ifrVisualizacao')
                                    time.sleep(1)
                                    #CLICA NO INCONE DE INCLUIR DOCUMENTO
                                    navegador.find_element(By.XPATH,'//*[@id="divArvoreAcoes"]/a[1]').click()
                                    time.sleep(1)
                                    #CLICA NA OPCAO DE DESPACHO DECISORIO
                                    navegador.find_element(By.XPATH,'//*[@id="tblSeries"]/tbody/tr[16]/td/a[2]').click()
                                    #SELECIONA TEXTO PADRAO
                                    navegador.find_element(By.XPATH,'//*[@id="divOptTextoPadrao"]/div').click()
                                    #ENVIA QUAL DESPACHO DECISORIO DEVE SER CRIADO
                                    navegador.find_element(By.XPATH,'//*[@id="txtTextoPadrao"]').send_keys('Despacho Decisório de Homologação não licenciados')
                                time.sleep(2)
                                #CLICA NA PRIMEIRA OPCAO
                                navegador.find_element(By.XPATH,'//*[@id="divInfraAjaxtxtTextoPadrao"]/ul/li/a').click()
                                #COLOCA O DESPACHO COMO PUBLICO
                                navegador.find_element(By.XPATH,'//*[@id="divOptPublico"]/div').click()
                                #SALVA DESPACHO
                                navegador.find_element(By.ID,'btnSalvar').click()
                                time.sleep(0.7)
                                #FECHA JANELA COM A VISUALIZACAO DO DESPACHO CRIADO
                                try:
                                    navegador.switch_to.window(navegador.window_handles[-1])
                                    navegador.close()
                                    time.sleep(0.7)
                                    #MUDA PARA JANELA PRINCIAPL DO PROGRAMA
                                    navegador.switch_to.window(janela_principal)
                                except:
                                    #MUDA PARA JANELA PRINCIAPL DO PROGRAMA
                                    navegador.switch_to.window(janela_principal)
                                #INCLUI DESPACHO NO BLOCO
                                navegador.switch_to.default_content()
                                navegador.switch_to.frame('ifrArvore')
                                #CLICA NO DESPACHO DECISORIO
                                navegador.find_element(By.PARTIAL_LINK_TEXT, "Despacho Decisório").click()
                                navegador.switch_to.default_content()
                                navegador.switch_to.frame('ifrVisualizacao')
                                time.sleep(1)
                                #CLICA NO ICONE DE LEGO
                                navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[8]').click()
                                #SELECIONA BLOCO (SELECIONA O PRIMEIRO DESPACHO PARA DRONES APROVADOS QUE LER)
                                select_element = navegador.find_element(By.ID, 'selBloco')
                                select = Select(select_element)
                                #PROCURA O PRIMEIRO BLOCO QUE TENHA O TEXTO "Despachos para Drones aprovados"
                                for opcao in select.options:
                                    if "Despachos para Drones aprovados" in opcao.text:
                                        select.select_by_visible_text(opcao.text)
                                        break
                                #CLICA NO BOTAO DE INCLUIR NO BLOCO
                                navegador.find_element(By.XPATH, '//*[@id="sbmIncluir"]').click()
                                #VOLTA PARA PAGINA INICIAL DO PROCESSO
                                navegador.switch_to.default_content()
                                navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processos)
                                elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
                                elementos.send_keys(Keys.ENTER)
                                navegador.switch_to.frame('ifrVisualizacao')
                                time.sleep(1)
                                #ADICIONA NOTA PARA AGUARDAR ASSINATURA
                                #CLICA NO ICONE DE ANOTACAO
                                navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[17]').click()
                                #INSERE O TEXTO DA ANOTACAO
                                navegador.find_element(By.ID, 'txaDescricao').send_keys('Aguardando assinatura.')
                                #SALVA O TEXTO
                                navegador.find_element(By.NAME, 'sbmRegistrarAnotacao').click()
                                #DA COMO APROVADO E ANOTA NO TXT DE PROCESSOS CONFORMES
                                situacao = 'Aprovado'
                                print("Próximo processo...")

                            #CONDICAO DE PROCESSO NAO CONFORME
                            else:
                                print(f'Tome as medidas necessárias para o processo {processos}.')
                                #MOSTRA OPCOES DE EXIGENCIA
                                while True:
                                    try:
                                        exigencia = int(input('Se houver alguma exigência digite [1], senão, digite [2].'))
                                        if exigencia == 1:
                                            break
                                        elif exigencia == 2:
                                            break
                                        else:
                                            print("Opção inválida, tente novamente!")
                                    except ValueError:
                                        print("Opção inválida, tente novamente!")
                                #CONDICAO DE EXIGENCIA
                                if exigencia == 1:
                                    navegador.switch_to.default_content()
                                    navegador.switch_to.frame('ifrArvore')
                                    #CLICA NA DECLARACAO DE CONFORMIDADE
                                    navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade").click()
                                    time.sleep(1)
                                    navegador.switch_to.default_content()
                                    navegador.switch_to.frame('ifrVisualizacao')
                                    navegador.switch_to.frame('ifrArvoreHtml')
                                    #ARMAZENA O EMAIL DO SOLICITANTE
                                    emailSol = navegador.find_element(By.XPATH, '/html/body/div[3]/a[1]').text
                                    navegador.switch_to.default_content()
                                    #RETORNA PARA A PAGINA INICIAL DO PROCESSO
                                    navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processos)
                                    elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
                                    elementos.send_keys(Keys.ENTER)
                                    time.sleep(1)
                                    navegador.switch_to.frame('ifrVisualizacao')
                                    #CLICA NO ICONE DE ENVIO DE EMAIL
                                    navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[11]').click()
                                    time.sleep(1)
                                    #VAI PARA A JANELA MAIS RECENTE ABERTA
                                    navegador.switch_to.window(navegador.window_handles[-1])
                                    time.sleep(1)
                                    #ABRE O DROPDOWN COM AS OPCOES DE EMAIL DA ANATEL
                                    select_element = navegador.find_element(By.XPATH, '//*[@id="selDe"]')
                                    select = Select(select_element)
                                    #SELECIONA O EMAIL DA ANATEL
                                    select.select_by_visible_text('ANATEL/E-mail de replicação <nao-responda@anatel.gov.br>')
                                    #ESCREVE O EMAIL DO SOLICITANTE
                                    navegador.find_element(By.XPATH, '//*[@id="s2id_autogen1"]').send_keys(emailSol)
                                    time.sleep(1)
                                    #CLICA NO EMAIL DO SOLICITANTE
                                    navegador.find_element(By.CLASS_NAME, 'select2-result-label').click()
                                    #MOSTRA AS OPCOES DE EXIGENCIA
                                    #EXECUTA A CONDICAO PARA CADA EXIGENCIA
                                    while True:
                                        try:
                                            print("Selecione o tipo de exigência abaixo:\n"
                                                " [1] Falta algum anexo no processo.\n",
                                                "[2] Erro na declaração de conformidade ou alguma exigência que é necessário abrir novo processo.\n",
                                                "[3] Outros.")
                                            time.sleep(0.5)
                                            exig = int(input("Opção: "))
                                        except ValueError:
                                            print("Opção inválida. Digite um número inteiro entre 1 e 3.")
                                            continue
                                        if exig in range(1, 4):
                                            if exig == 1:
                                                #PUXA FUNCAO INSIRA_ANEXO CONTENDO O CORPO DE EMAIL PEDINDO ANEXOS
                                                insira_anexo(processos)
                                                #PEDE PARA USUARIO CONFERIR SE ESTÁ TUDO CERTO COM O EMAIL
                                                while True:
                                                    try:
                                                        confere = int(input('Caso esteja tudo certo digite [1], caso tenha algum erro digite [2]: '))
                                                        if confere == 1:
                                                            break
                                                        elif confere == 2:
                                                            break
                                                        else:
                                                            print("Opção inválida, tente novamente!")
                                                    except ValueError:
                                                        print("Opção inválida, tente novamente!")
                                                if confere == 1:
                                                    #ENVIA EMAIL COM A EXIGENCIA
                                                    navegador.find_element(By.XPATH, '//*[@id="divInfraBarraComandosInferior"]/button[1]').click()
                                                    time.sleep(0.5)
                                                    #FECHA ALERTA DO NAVEGADOR
                                                    alert = Alert(navegador)
                                                    alert.accept()
                                                    time.sleep(0.5)
                                                    #VOLTA O FOCO DO SELENIUM PARA A PAGINA PRINCIPAL NOVAMENTE
                                                    navegador.switch_to.window(janela_principal)
                                                    navegador.switch_to.default_content()
                                                    #VOLTA PARA A PAGINA INICIAL DO PROCESSO
                                                    navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processos) 
                                                    elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
                                                    elementos.send_keys(Keys.ENTER)    
                                                    time.sleep(1)
                                                    #INSERE TAG REFERENTE A PROCESSOS INTERCORRENTES
                                                    navegador.switch_to.frame('ifrVisualizacao')
                                                    #CLICA NO ICONE DE TAG
                                                    navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[25]').click()
                                                    try:    
                                                        navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
                                                    except:
                                                        navegador.find_element(By.XPATH, '//*[@id="tblMarcadores"]/tbody/tr[2]/td[6]/a[2]/img').click()
                                                        alert.accept()
                                                        navegador.find_element(By.XPATH, '//*[@id="btnAdicionar"]').click()
                                                        time.sleep(0.5)
                                                        navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
                                                    
                                                    time.sleep(0.5)
                                                    #CLICA NA TAG DE INTERCORRENTE
                                                    navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/ul/li[18]/a').click()
                                                    #PEDE O TEXTO DA TAG
                                                    textoTag = input("Insira o texto da tag: ")
                                                    #COLOCA O TEXTO NA TAG
                                                    navegador.find_element(By.XPATH, '//*[@id="txaTexto"]').send_keys(textoTag)
                                                    #SALVA TAG
                                                    navegador.find_element(By.XPATH, '//*[@id="sbmSalvar"]').click()
                                                    #DEFINE SITUACAO DO PROCESSO
                                                    situacao = 'Intercorrente'
                                                else:
                                                    #CANCELA ENVIO DO EMAIL SE HOUVER ALGUM ERRO
                                                    print("Faça as alterações necessárias mais tarde.")
                                                    navegador.find_element(By.XPATH, '//*[@id="btnCancelar"]').click()
                                                    navegador.switch_to.window(janela_principal)
                                                    return
                                            elif exig == 2:
                                                #PUXA FUNCAO ERRO_DECLARACAO CONTENDO O CORPO DE EMAIL PEDINDO PARA ABRIR NOVO PROCESSO SEI
                                                erro_declaracao(processos, impProp)
                                                #PEDE PARA USUARIO CONFERIR O EMAIL
                                                while True:
                                                    try:
                                                        confere = int(input('Caso esteja tudo certo digite [1], caso tenha algum erro digite [2]: '))
                                                        if confere == 1:
                                                            break
                                                        elif confere == 2:
                                                            break
                                                        else:
                                                            print("Opção inválida, tente novamente!")
                                                    except ValueError:
                                                        print("Opção inválida, tente novamente!")
                                                if confere == 1:
                                                    #ENVIA O EMAIL COM EXIGENCIA
                                                    navegador.find_element(By.XPATH, '//*[@id="divInfraBarraComandosInferior"]/button[1]').click()
                                                    time.sleep(0.5)
                                                    #FECHA ALERTA DO NAVEGADOR
                                                    alert = Alert(navegador)
                                                    alert.accept()
                                                    time.sleep(0.5)
                                                    #RETORNA O FOCO PARA A JANELA PRINCIPAL
                                                    navegador.switch_to.window(janela_principal)
                                                    navegador.switch_to.default_content()
                                                    #VOLTA PARA PAGINA INICIAL DO PROCESSO
                                                    navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processos) 
                                                    elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
                                                    elementos.send_keys(Keys.ENTER)    
                                                    time.sleep(1)
                                                    #INSERE TAG REFERENTE A PROCESSOS INTERCORRENTES 
                                                    navegador.switch_to.frame('ifrVisualizacao')
                                                    #CLICA NO ICONE DE TAG
                                                    navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[25]').click()
                                                    try:    
                                                        navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
                                                    except:
                                                        navegador.find_element(By.XPATH, '//*[@id="tblMarcadores"]/tbody/tr[2]/td[6]/a[2]/img').click()
                                                        alert.accept()
                                                        navegador.find_element(By.XPATH, '//*[@id="btnAdicionar"]').click()
                                                        time.sleep(0.5)
                                                        navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
                                                    
                                                    #CLICA NO DROPDOWN COM OS TIPOS DE TAG
                                                    time.sleep(0.5)
                                                    #CLICA NA TAG DE PENDENCIA
                                                    navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/ul/li[19]/a').click()
                                                    #PEDE O TEXTO DA TAG
                                                    textoTag = input("Insira o texto da tag: ")
                                                    navegador.find_element(By.XPATH, '//*[@id="txaTexto"]').send_keys(textoTag)
                                                    #SALVA A TAG
                                                    navegador.find_element(By.XPATH, '//*[@id="sbmSalvar"]').click()
                                                    #DEFINE SITUACAO DO PROCESSO
                                                    situacao = 'Exigência'
                                                else:
                                                    #CANCELA O ENVIO CASO TENHA ALGUM ERRO
                                                    print("Faça as alterações necessárias mais tarde.")
                                                    navegador.find_element(By.XPATH, '//*[@id="btnCancelar"]').click()
                                                    navegador.switch_to.window(janela_principal)
                                                    return
                                            elif exig == 3:
                                                #PUXA FUNCAO OUTRO_ERRO EM QUE USUARIO INSERE O EMAIL COMO QUISER
                                                outro_erro()
                                                #PEDE PARA USUARIO ESCREVER EMAIL E DIGITAR 1 PARA CONTINUAR
                                                while True:
                                                    try:
                                                        confere = int(input('Escreva o email da forma que deseja e logo em seguida digite [1], caso queira cancelar digite [2]: '))
                                                        if confere == 1:
                                                            break
                                                        elif confere == 2:
                                                            break
                                                        else:
                                                            print("Opção inválida, tente novamente!")
                                                    except ValueError:
                                                        print("Opção inválida, tente novamente!")
                                                if confere == 1:
                                                    #ENVIA EMAIL
                                                    navegador.find_element(By.XPATH, '//*[@id="divInfraBarraComandosInferior"]/button[1]').click()
                                                    time.sleep(0.5)
                                                    #FECHA ALERTA DO NAVEGADOR
                                                    alert = Alert(navegador)
                                                    alert.accept()
                                                    time.sleep(0.5)
                                                    #RETORNA O FOCO PARA A JANELA PRINCIPAL
                                                    navegador.switch_to.window(janela_principal)
                                                    navegador.switch_to.default_content()
                                                    #VAI PARA A JANELA INICIAL DO PROCESSO
                                                    navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processos) 
                                                    elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
                                                    elementos.send_keys(Keys.ENTER)    
                                                    time.sleep(1)
                                                    navegador.switch_to.frame('ifrVisualizacao')
                                                    #CLICA NO ICONE DE TAG
                                                    navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[25]').click()
                                                    try:    
                                                        navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
                                                    except:
                                                        navegador.find_element(By.XPATH, '//*[@id="tblMarcadores"]/tbody/tr[2]/td[6]/a[2]/img').click()
                                                        alert.accept()
                                                        navegador.find_element(By.XPATH, '//*[@id="btnAdicionar"]').click()
                                                        time.sleep(0.5)
                                                        navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
                                                    
                                                    time.sleep(0.5)
                                                    #INSERE TAG DE 'PENDENCIAS'
                                                    navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/ul/li[19]/a').click()
                                                    #PEDE TEXTO DA TAG
                                                    textoTag = input("Insira o texto da tag: ")
                                                    #INSERE O TEXTO DA TAG
                                                    navegador.find_element(By.XPATH, '//*[@id="txaTexto"]').send_keys(textoTag)
                                                    #SALVA TAG
                                                    navegador.find_element(By.XPATH, '//*[@id="sbmSalvar"]').click()
                                                    #DEFINE SITUACAO DO PROCESSO
                                                    situacao = 'Exigência'
                                                else:
                                                    #CANCELA ENVIO DO PROCESSO
                                                    print("Faça as alterações necessárias mais tarde.")
                                                    navegador.find_element(By.XPATH, '//*[@id="btnCancelar"]').click()
                                                    navegador.switch_to.window(janela_principal)
                                                    return
                                            #VOLTA PARA A PAGINA INICIAL DO PROCESSO
                                            navegador.switch_to.default_content()
                                            navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processos) 
                                            elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
                                            elementos.send_keys(Keys.ENTER)
                                            time.sleep(1)
                                            navegador.switch_to.frame('ifrVisualizacao')
                                            #CLICA NO ICONE DE CONCLUIR PROCESSO
                                            navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[20]').click()
                                            print("Próximo processo...")
                                            break
                                        #CONDICAO PARA OPCAO INEXISTENTE
                                        else:
                                            print("Opção inválida. Tente novamente.")
                                #CONDICAO CASO PRECISE PULAR O PROCESSO
                                else:
                                    print("Pulando processo.")
                            #VERIFICA SE O PROCESSO CONTEM ALGUMA SITUACAO ANTES DE ESCREVER SOBRE
                            #SE HOUVER ALGUM ERRO OU PULAR PROCESSO NAO ESCREVERA NA PLANILHA
                            if situacao == 'Aprovado' or situacao == 'Exigência' or situacao == 'Intercorrente':
                                #ENCONTRA O ICONE DO EDGE E ABRE O NAVEGADOR DA PLANILHA
                                pyautogui.PAUSE = 0.7
                                edge=pyautogui.locateOnScreen('imagensAut/edge.png', confidence=0.7)
                                #COPIA NUMERO DO PROCESSO
                                pyperclip.copy(processos)
                                #CLICA NO NAVEGADOR
                                pyautogui.click(edge)
                                time.sleep(0.5)
                                #PESQUISA PROCESSO NA PLANILHA
                                pyautogui.hotkey('ctrl', 'l')
                                time.sleep(0.3)
                                #COLA NUMERO DO PROCESSO E APERTA ENTER PARA PESQUISAR O PROCESSO
                                pyautogui.hotkey('ctrl', 'v')
                                time.sleep(0.3)
                                pyautogui.press('enter')
                                time.sleep(0.3)
                                pyautogui.press('esc')
                                #NAVEGA ENTRE AS CELULAS E INSERE OS DADOS SOBRE O PROCESSO
                                pyautogui.press('right')
                                pyautogui.PAUSE = 0.3
                                pyperclip.copy(nomeEstag)
                                pyautogui.hotkey('ctrl', 'v')
                                pyautogui.press('right')
                                pyperclip.copy(retido)
                                pyautogui.hotkey('ctrl', 'v')
                                pyautogui.press('right')
                                #VERIFICA SE ESTÁ RETIDO, SE NAO ESTIVER NAO ESCREVE CODIGO DE RASTREIO
                                if retido == 'Sim':
                                    pyperclip.copy(codigo_rastreio)
                                    pyautogui.hotkey('ctrl', 'v')
                                    pyautogui.press('right')
                                else:
                                    pyautogui.press('right')
                                pyautogui.press('right')
                                pyperclip.copy(nome_interessado)
                                pyautogui.hotkey('ctrl', 'v')
                                pyautogui.press('right')
                                pyperclip.copy(situacao)
                                pyautogui.hotkey('ctrl', 'v')
                                #VOLTA PARA A JANELA DO CHROME
                                chrome=pyautogui.locateOnScreen('imagensAut/chrome.png', confidence=0.7)
                                pyautogui.click(chrome)
                            else:
                                print('--------------------------------------------------------------------------')
                                print('--------------------------------------------------------------------------')
                    except Exception as error:
                        print(f"Ocorreu um erro: {error}")
                    #REMOVE O PROCESSO ANALISADO DA LISTA
                    lista_processos.remove(processos)
                    break
                else:
                    print("Opção inválida, tente novamente!")
            except ValueError:
                print("Opção inválida, tente novamente!")
            pyautogui.PAUSE = 0.7

#FUNCAO PARA CONCLUIR PROCESSO
def concluiProcesso():
    #DEFINE EMAIL PARA PROCESSO RETIDO
    textoRetido = '''Prezado(a) Senhor(a),

    Em atenção ao pedido de homologação constante do processo SEI em referência, informamos que:

    1. O pedido foi APROVADO
    2. O Despacho Decisório que aprovou o pedido está disponível publicamente por meio do sistema SEI na área de Pesquisa Pública, no link:

            https://sei.anatel.gov.br/sei/modulos/pesquisa/md_pesq_processo_pesquisar.php?acao_externa=protocolo_pesquisar&acao_origem_externa=protocolo_pesquisar&id_orgao_acesso_externo=0


    3. O Despacho Decisório deverá ser portado junto ao Equipamento (fisicamente ou eletronicamente), para que as autoridades competentes possam conferir a regularidade, quando necessário.
    4. Visto que o produto se encontra retido, para que seja informado o número de série, solicitamos que acesse o sistema SEI da Anatel por meio do seguinte link:


            https://sei.anatel.gov.br/sei/controlador_externo.php?acao=usuario_externo_logar&id_orgao_acesso_externo=0.


    5. Após autenticação no sistema SEI, solicitamos que inclua os documentos selecionando a opção "Intercorrente" e informando o número do processo em referência.



    FAVOR NÃO RESPONDER ESTE E-MAIL.

    Atenciosamente,

    ORCN - Gerência de Certificação e Numeração

    SOR - Superintendência de Outorga e Recursos à Prestação

    Anatel - Agência Nacional de Telecomunicações'''

    #DEFINE EMAIL PARA PROCESSO NAO RETIDO
    textoNaoRetido = '''Prezado(a) Senhor(a),

    Em atenção ao pedido de homologação constante do processo SEI em referência, informamos que:

    1. O pedido foi APROVADO
    2. O Despacho Decisório que aprovou o pedido está disponível publicamente por meio do sistema SEI na área de Pesquisa Pública, no link:

            https://sei.anatel.gov.br/sei/modulos/pesquisa/md_pesq_processo_pesquisar.php?acao_externa=protocolo_pesquisar&acao_origem_externa=protocolo_pesquisar&id_orgao_acesso_externo=0


    3. O Despacho Decisório deverá ser portado junto ao Equipamento (fisicamente ou eletronicamente), para que as autoridades competentes possam conferir a regularidade, quando necessário.

    FAVOR NÃO RESPONDER ESTE E-MAIL.

    Atenciosamente,

    ORCN - Gerência de Certificação e Numeração

    SOR - Superintendência de Outorga e Recursos à Prestação

    Anatel - Agência Nacional de Telecomunicações'''

    #VARIAVEL PARA VOLTAR À JANELA INICIAL
    janela_principal = navegador.current_window_handle

    #ITERA SOBRE OS PROCESSOS DA LISTA DE PROCESSOS CONFORMES
    for processosAssinados in lista_procConformes[:]:
        while True:
            #CRIA CONDICAO PARA PARAR DE CONCLUIR PROCESSO
            try:
                print('--------------------------------------------------------------------------')
                concluir = int(input("Deseja concluir o próximo processo? Se sim digite [1], se não digite [2]: "))
                print('--------------------------------------------------------------------------')
                if concluir == 2:
                    return
                else:
                    try:
                        #ENTRA NA PAGINA DO PROCESSO
                        navegador.switch_to.default_content()
                        navegador.find_element(By.ID, 'txtPesquisaRapida').clear()
                        navegador.find_element(By.ID, 'txtPesquisaRapida').send_keys(processosAssinados)
                        elementos = navegador.find_element(By.ID, 'txtPesquisaRapida')
                        elementos.send_keys(Keys.ENTER)
                        time.sleep(1)
                        #PEGA DADOS DA DECLARACAO DE CONFORMIDADE
                        navegador.switch_to.frame('ifrArvore')
                        #CONFERE SE EXISTE A DECLARACAO DE CONFORMIDADE OU UMA PASTA
                        if check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade") or check_element_exists(By.XPATH, '//*[@id="spanPASTA1"]'):
                            #CASO HAJA PASTA ELE CLICARA NA PRIMEIRA
                            if check_element_exists(By.XPATH, '//*[@id="spanPASTA1"]'):
                                navegador.find_element(By.XPATH, '//*[@id="spanPASTA1"]').click()
                                time.sleep(1)
                        #CLICA NA DECLARACAO DE CONFORMIDADE
                        navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade").click()
                        #CONFERE SE HÁ CODIGO DE RASTREIO
                        navegador.switch_to.default_content()
                        navegador.switch_to.frame('ifrVisualizacao')
                        navegador.switch_to.frame('ifrArvoreHtml')
                        tabela_cod = navegador.find_element(By.XPATH, '/html/body/table[4]/tbody/tr/td[1]').text
                        #FLUXO DE DRONES PRE APROVADOS
                        #CONFERE SE HÁ CÓDIGO DE RASTREIO E SE É PRE APROVADO PELO INDICE DAS TABELAS E O SEU TEXTO
                        if check_element_exists(By.XPATH, '/html/body/table[4]/tbody/tr/td[2]') and tabela_cod == 'Código de rastreio':
                            #ARMAZENA CODIGO DE RASTREIO
                            codigo_rastreioConforme = navegador.find_element(By.XPATH, '/html/body/table[4]/tbody/tr/td[2]').text
                            #ARMAZENA EMAIL DO SOLICITANTE
                            emailSol = navegador.find_element(By.XPATH, '/html/body/div[3]/a[1]').text
                            time.sleep(0.2)
                            navegador.switch_to.default_content()
                            navegador.switch_to.frame('ifrArvore')
                            #VERIFICA SE EXISTE O DESPACHO DECISORIO
                            if check_element_exists(By.PARTIAL_LINK_TEXT, "Despacho Decisório"):
                                #CLICA NO DESPACHO DECISORIO
                                navegador.find_element(By.PARTIAL_LINK_TEXT, "Despacho Decisório").click()
                                navegador.switch_to.default_content()
                                navegador.switch_to.frame('ifrVisualizacao')
                                navegador.switch_to.frame('ifrArvoreHtml')
                                #VERIFICA SE EXISTE ASSINATURA DO GERENTE
                                if check_element_exists(By.XPATH, "/html/body/div[1]/table[1]"):
                                    navegador.switch_to.default_content()
                                    navegador.switch_to.frame('ifrVisualizacao')
                                    #CLICA NO ICONE DE EMAIL
                                    navegador.find_element(By.XPATH, "//img[@title='Enviar Documento por Correio Eletrônico']").click()
                                    time.sleep(0.5)
                                    #MUDA PARA A JANELA MAIS RECENTE
                                    navegador.switch_to.window(navegador.window_handles[-1])
                                    time.sleep(1)
                                    #SELECIONA O DROPDOWN COM AS OPCOES DE EMAIL
                                    select_element = navegador.find_element(By.XPATH, '//*[@id="selDe"]')
                                    select = Select(select_element)
                                    #SELECIONA A OPCAO DE EMAIL ANATEL
                                    select.select_by_visible_text('ANATEL/E-mail de replicação <nao-responda@anatel.gov.br>')
                                    #INSERE EMAIL DO SOLICITANTE
                                    navegador.find_element(By.XPATH, '//*[@id="s2id_autogen1"]').send_keys(emailSol)
                                    time.sleep(1)
                                    #CLICA NO EMAIL DO SOLICITANTE
                                    navegador.find_element(By.XPATH, '//*[@id="select2-result-label-2"]').click()
                                    #INSERE ASSUNTO DO EMAIL
                                    navegador.find_element(By.ID, 'txtAssunto').send_keys(f'Processo SEI nº {processosAssinados} - Aprovado')
                                    #VERIFICA SE HA CODIGO DE RASTREIO
                                    #SE NAO HA INSERE TEXTO DE NAO RETIDO
                                    if not codigo_rastreioConforme.strip():
                                        navegador.find_element(By.ID, 'txaMensagem').send_keys(textoNaoRetido)
                                        prod = 'naoRetido'
                                    #SE HA INSERE TEXTO DE RETIDO
                                    else:
                                        navegador.find_element(By.ID, 'txaMensagem').send_keys(textoRetido)
                                        prod = 'Retido'
                                    #PEDE PARA VERIFICAR SE O EMAIL ESTA CORRETO
                                    while True:
                                        try:
                                            verifica = int(input('Verifique se está tudo nos conformes, se sim, digite [1], caso não, digite [2]: '))
                                            if verifica == 1:
                                                break
                                            elif verifica == 2:
                                                break
                                            else:
                                                print("Opção inválida, tente novamente!")
                                        except ValueError:
                                            print("Opção inválida, tente novamente!")
                                    if verifica == 1:
                                        #ENVIA EMAIL
                                        navegador.find_element(By.XPATH, '//*[@id="divInfraBarraComandosInferior"]/button[1]').click()
                                        #FECHA ALERTA DO NAVEGADOR
                                        alert = Alert(navegador)
                                        alert.accept()
                                        #RETORNA PARA A JANELA PRINCIPAL
                                        navegador.switch_to.window(janela_principal)
                                        navegador.switch_to.default_content()
                                        #VOLTA PARA A JANELA INICIAL DO PROCESSO
                                        navegador.find_element(By.ID, 'txtPesquisaRapida').send_keys(processosAssinados)
                                        elementos = navegador.find_element(By.ID, 'txtPesquisaRapida')
                                        elementos.send_keys(Keys.ENTER)
                                        time.sleep(1)
                                        #CLICA NO ICONE DE ANOTACAO
                                        navegador.switch_to.frame('ifrVisualizacao')
                                        navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[17]').click()
                                        #LIMPA O TEXTO DA ANOTACAO
                                        navegador.find_element(By.XPATH, '//*[@id="txaDescricao"]').clear()
                                        #SALVA ANOTACAO
                                        navegador.find_element(By.XPATH, '//*[@id="divInfraBarraComandosSuperior"]/button').click()
                                        #CLICA NO ICONE DE TAG
                                        navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[25]').click()
                                        try:
                                            navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click() 
                                        except:
                                            navegador.find_element(By.XPATH, '//*[@id="tblMarcadores"]/tbody/tr[2]/td[6]/a[2]/img').click()
                                            alert.accept()
                                            navegador.find_element(By.XPATH, '//*[@id="btnAdicionar"]').click()
                                            time.sleep(0.5)
                                            navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
                                        #CLICA NO DROPDOWN DE TAG
                                        time.sleep(0.5)
                                        #VERIFICA SE ESTA RETIDO
                                        if prod == 'naoRetido':
                                            #CLICA NA TAG DE NAO RETIDO
                                            navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/ul/li[16]').click()
                                        else:
                                            #CLICA NA TAG DE RETIDO
                                            navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/ul/li[17]').click()
                                        #SALVA TAG
                                        navegador.find_element(By.XPATH, '//*[@id="sbmSalvar"]').click()
                                        navegador.switch_to.default_content()
                                        #ENTRA NA PAGINA INICIAL DO PROCESSO
                                        navegador.find_element(By.ID, 'txtPesquisaRapida').send_keys(processosAssinados)
                                        elementos = navegador.find_element(By.ID, 'txtPesquisaRapida')
                                        elementos.send_keys(Keys.ENTER)
                                        time.sleep(1)
                                        #CONCLUI PROCESSO
                                        navegador.switch_to.frame('ifrVisualizacao')
                                        navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[20]').click()
                                        time.sleep(0.3)
                                        navegador.switch_to.default_content()
                                        
                                    else:
                                        #CANCELA ENVIO DO EMAIL
                                        navegador.find_element(By.XPATH, '//*[@id="btnCancelar"]')
                                        print(f"Confira o processo {processosAssinados} mais tarde.")
                                        navegador.switch_to.window(janela_principal)
                                        navegador.switch_to.default_content()
                                else:
                                    #NAO ENVIA EMAIL CASO NAO ESTEJA COM DESPACHO
                                    print(f"O despacho do processo {processosAssinados} ainda não foi assinado, execute o script novamente após conter a assinatura do gerente.")
                                    time.sleep(0.5)
                        #FLUXO DE DRONES IMPORTADOS PRA USO PROPRIO           
                        elif check_element_exists(By.XPATH, '/html/body/table[5]/tbody/tr/td[2]') and tabela_cod != 'Código de rastreio':
                            #CONFERE SE ESTA RETIDO PELO INDICE E TEXTO DA TABELA DE CODIGO DE RASTREIO
                            codigo_rastreioConforme = navegador.find_element(By.XPATH, '/html/body/table[5]/tbody/tr/td[2]').text
                            #ARMAZENA EMAIL DO SOLICITANTE
                            emailSol = navegador.find_element(By.XPATH, '/html/body/div[3]/a[1]').text
                            navegador.switch_to.default_content()
                            navegador.switch_to.frame('ifrArvore')
                            #VERIFICA SE EXISTE DESPACHO
                            if check_element_exists(By.PARTIAL_LINK_TEXT, "Despacho Decisório"):
                                #CLICA NO DESPACHO
                                navegador.find_element(By.PARTIAL_LINK_TEXT, "Despacho Decisório").click()
                                navegador.switch_to.default_content()
                                navegador.switch_to.frame('ifrVisualizacao')
                                navegador.switch_to.frame('ifrArvoreHtml')
                                #VERIFICA ASSINATURA DO DESPACHO DECISORIO  
                                if check_element_exists(By.XPATH, "/html/body/div[1]/table[1]"):
                                    #ENVIA EMAIL DE APROVACAO DO PROCESSO
                                    navegador.switch_to.default_content()
                                    navegador.switch_to.frame('ifrVisualizacao')
                                    #CLICA NO ICONE DE ENVIAR EMAIL
                                    navegador.find_element(By.XPATH, "//img[@title='Enviar Documento por Correio Eletrônico']").click()
                                    time.sleep(0.5)
                                    #MUDA PARA JANELA MAIS RECENTE
                                    navegador.switch_to.window(navegador.window_handles[-1])
                                    time.sleep(1)
                                    #SELECIONA EMAIL DA ANATEL
                                    select_element = navegador.find_element(By.ID, 'selDe')
                                    select = Select(select_element)
                                    select.select_by_visible_text('ANATEL/E-mail de replicação <nao-responda@anatel.gov.br>')
                                    #ESCREVE EMAIL DO SOLICITANTE
                                    navegador.find_element(By.XPATH, '//*[@id="s2id_autogen1"]').send_keys(emailSol)
                                    time.sleep(1)
                                    #CLICA NO EMAIL DO SOLICITANTE
                                    navegador.find_element(By.XPATH, '//*[@id="select2-result-label-2"]').click()
                                    #ESCREVE ASSUNTO DO EMAIL
                                    navegador.find_element(By.ID, 'txtAssunto').send_keys(f'Processo SEI nº {processosAssinados} - Aprovado')
                                    #DIFERE O EMAIL PARA PRODUTO RETIDO E NAO RETIDO
                                    if not codigo_rastreioConforme.strip():
                                        #EMAIL PARA NAO RETIDO
                                        navegador.find_element(By.ID, 'txaMensagem').send_keys(textoNaoRetido)
                                        prod = 'naoRetido'
                                    else:
                                        #EMAIL PARA RETIDO
                                        navegador.find_element(By.ID, 'txaMensagem').send_keys(textoRetido)
                                        prod = 'Retido'
                                    #PEDE PARA O USUARIO CONFIRMAR AS INFORMACOES
                                    while True:
                                        try:
                                            verifica = int(input('Verifique se está tudo nos conformes, se sim, digite [1], caso não, digite [2]: '))
                                            if verifica == 1:
                                                break
                                            elif verifica == 2:
                                                break
                                            else:
                                                print("Opção inválida, tente novamente!")
                                        except ValueError:
                                            print("Opção inválida, tente novamente!")
                                    if verifica == 1:
                                        #ENVIA EMAIL
                                        navegador.find_element(By.XPATH, '//*[@id="divInfraBarraComandosSuperior"]/button[1]').click()
                                        #FECHA ALERTA DO NAVEGADOR
                                        alert = Alert(navegador)
                                        alert.accept()      
                                        #MUDA PARA JANELA PRINCIPAL
                                        navegador.switch_to.window(janela_principal)
                                        navegador.switch_to.default_content()
                                        #RETORNA PARA PAGINA INICIAL DO PROCESSO
                                        navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processosAssinados)
                                        elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
                                        elementos.send_keys(Keys.ENTER)
                                        time.sleep(1)
                                        navegador.switch_to.frame('ifrVisualizacao')
                                        #CLICA NO ICONE DE ANOTACAO
                                        navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[17]').click()
                                        #APAGA ANOTACAO DE AGUARDANDO ASSINATURA
                                        navegador.find_element(By.XPATH, '//*[@id="txaDescricao"]').clear()
                                        #SALVA MUDANCAS
                                        navegador.find_element(By.XPATH, '//*[@id="divInfraBarraComandosSuperior"]/button').click()
                                        #INSERE TAG NO PROCESSO
                                        navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[25]').click()
                                        try:  
                                            navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
                                        except:
                                            navegador.find_element(By.XPATH, '//*[@id="tblMarcadores"]/tbody/tr[2]/td[6]/a[2]/img').click()
                                            alert.accept()
                                            navegador.find_element(By.XPATH, '//*[@id="btnAdicionar"]').click()
                                            time.sleep(0.5)
                                            navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
                                        time.sleep(0.5)
                                        #DIFERE A TAG PARA RETIDO E NAO RETIDO
                                        if prod == 'naoRetido':
                                            #TAG DE NAO RETIDO
                                            navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/ul/li[16]').click()                            
                                        else:
                                            #TAG DE RETIDO
                                            navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/ul/li[17]').click()
                                        #SALVA TAG
                                        navegador.find_element(By.XPATH, '//*[@id="sbmSalvar"]').click()
                                        navegador.switch_to.default_content()
                                        time.sleep(0.3)
                                        #VOLTA PARA PAGINA INICIAL DO PROCESSO
                                        navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processosAssinados)
                                        elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
                                        elementos.send_keys(Keys.ENTER)
                                        time.sleep(1)
                                        navegador.switch_to.frame('ifrVisualizacao')
                                        #CONCLUI PROCESSO
                                        navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[20]').click()
                                        time.sleep(0.3)
                                        navegador.switch_to.default_content()
                                    #CASO HAJA ERRO NO EMAIL ELE CANCELA O ENVIO
                                    elif verifica == 2:
                                        navegador.find_element(By.XPATH, '//*[@id="btnCancelar"]')
                                        print(f"Confira o processo {processosAssinados} mais tarde.")
                                        navegador.switch_to.window(janela_principal)
                                        navegador.switch_to.default_content()
                                    
                                #CASO O DESPACHO NAO ESTEJA ASSINADO ELE PEDE PARA AGUARDAR
                                else:
                                    print(f"O despacho do processo {processosAssinados} ainda não foi assinado, execute o script novamente após conter a assinatura do gerente.")
                                    time.sleep(0.5)
                        lista_procConformes.remove(processosAssinados)        
                        break
                    #CONDICAO DE ERRO
                    except Exception:
                        print("Ocorreu algum erro.\nPulando processo...")
                        break
            #CONDICAO DE ERRO
            except ValueError:
                print("Opção inválida, tente novamente")


# Função para atribuir processos
def atribuir_processos(file_path, num_processos):
    # Lê a planilha do Excel
    df = pd.read_excel(file_path)

    # Filtra as linhas onde a segunda coluna (Estagiario responsável) é uma data
    df['Estagiario responsável'] = pd.to_datetime(df['Estagiario responsável'], format='%d/%m/%Y', errors='coerce')
    df_com_datas = df.dropna(subset=['Estagiario responsável'])

    # Remove componente de tempo das datas usando .loc[]
    df_com_datas.loc[:, 'Estagiario responsável'] = df_com_datas['Estagiario responsável'].dt.date

    # Pega as n datas mais antigas
    processos_mais_antigos = df_com_datas.head(num_processos)

    # Cria a lista de processos a serem atribuídos
    lista_processos_atr = processos_mais_antigos['Nº do Processo SEI'].tolist()

    return lista_processos_atr


def atribuicao():
    # Exemplo de uso
    caminho = input("Insira o caminho da planilha geral: ")
    file_path = caminho.replace('"','').replace('\\','/')
    navegador.find_element(By.ID, 'lnkControleProcessos').click()
    navegador.find_element(By.ID, 'ancLiberarMeusProcessos').click()
    num_processos = int(input('Digite a quantidade de processos que deseja atribuir: '))
    lista_processos_atr = atribuir_processos(file_path, num_processos)
    print('Lista de processos atribuídos:', lista_processos_atr)
    navegador.find_element(By.PARTIAL_LINK_TEXT,'chirlene.colab').click()
    processos_para_atr = []
    for processos_atr in lista_processos_atr:
        try:
            navegador.find_element(By.XPATH, f'//label[@title="{processos_atr}"]').click()
            processos_para_atr.append(processos_atr)
            time.sleep(0.3)
        except: 
            print('Processo não encontrado.')

    select_element = navegador.find_element(By.XPATH, '//*[@id="selAtribuicao"]')
    select = Select(select_element)
    time.sleep(0.5)
    for opcao in select.options:
        if nomeEstag_sem_acento in opcao.text:
            select.select_by_visible_text(opcao.text)
            break
    navegador.find_element(By.ID, 'btnSalvar').click()
    navegador.find_element(By.XPATH,'//*[@id="divFiltro"]/div[2]/a').click()
    pyautogui.PAUSE = 0.7
    edge=pyautogui.locateOnScreen('imagensAut/edge.png', confidence=0.7)
    #CLICA NO NAVEGADOR
    pyautogui.click(edge)
    time.sleep(0.5)
    for processos_atr2 in processos_para_atr:
        pyautogui.PAUSE = 0.7
        #COPIA NUMERO DO PROCESSO
        pyperclip.copy(processos_atr2)
        #PESQUISA PROCESSO NA PLANILHA
        pyautogui.hotkey('ctrl', 'l')
        time.sleep(0.5)
        #COLA NUMERO DO PROCESSO E APERTA ENTER PARA PESQUISAR O PROCESSO
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.3)
        pyautogui.press('enter')
        pyautogui.PAUSE = 0.4
        pyautogui.press('esc')
        pyautogui.press('right')
        pyperclip.copy(nomeEstag)
        pyautogui.hotkey('ctrl', 'v')
    
#INSTALA O CHROME DRIVEr MAIS ATUALIZADO
servico = Service(ChromeDriverManager().install())
#DEFINE O TEMPO DE EXECUÇÃO PARA CADA COMANDO DO PYAUTOGUI
pyautogui.PAUSE = 0.7
#INICIA O NAVEGADOR
navegador = uc.Chrome(service=servico)
navegador.maximize_window()
#ENTRA NO SEI
navegador.get('https://sei.anatel.gov.br/')
#LOCALIZA O ICONE DO EDGE E ABRE A PLANILHA GERAL
#FOI UTILIZADO O PYPERCLIP PARA EVITAR QUALQUER ERRO NA HORA DE COLAR O URL DA PLANILHA
time.sleep(1)
edge=pyautogui.locateOnScreen('imagensAut/edge.png', confidence=0.7)
pyautogui.click(edge)
time.sleep(3)
sitePlan = 'https://anatel365.sharepoint.com/:x:/r/sites/lista.orcn/_layouts/15/Doc.aspx?sourcedoc=%7B4130A4D6-7F00-45D4-A328-ED0866A62335%7D&file=Distribui%C3%A7%C3%A3o%20Processo%20Drone.xlsx&action=default&mobileredirect=true'
pyperclip.copy(sitePlan)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
chrome=pyautogui.locateOnScreen('imagensAut/chrome.png', confidence=0.7)
pyautogui.click(chrome)

while True:
    #INICIA JANELA
    try:
        root = tk.Tk()
        root.title("Login no SEI")
        root.attributes('-topmost', True)
        #MAXIMIZA O TAMANHO DA TELA ABERTA PEGANDO INFORMAÇÕES DA TELA EM QUE ESTÁ SENDO EXECUTADA
        largura_tela = root.winfo_screenwidth()
        root.geometry(f'+{largura_tela}+0')
        root.state('zoomed')

        #RECEBE USUARIO E SENHA DA FUNÇÃO preencher_campos()
        user_var = tk.StringVar()
        senha_var = tk.StringVar()

        frame = tk.Frame(root)
        frame.place(relx=0.5, rely=0.5, anchor='center')

        #CRIA LABEL CENTRALIZADO
        tk.Label(frame, text="LOGIN NO SEI ANATEL", font=('Arial',30)).grid(row=0, columnspan=2)
        tk.Label(frame, text="").grid(row=1, column=0)
        tk.Label(frame, text="").grid(row=2, column=0)

        #CRIA INPUT DE USUARIO
        tk.Label(frame, text="Usuário:", font=('Arial', 15)).grid(row=3, column=0, sticky='e')
        user_entry = tk.Entry(frame, textvariable=user_var, width=30)
        user_entry.grid(row=3, column=1, sticky='w')

        #CRIA INPUT COM SENHA CENSURADA
        tk.Label(frame, text="Senha:", font=('Arial', 15)).grid(row=4, column=0, sticky='e')
        senha_entry = tk.Entry(frame, textvariable=senha_var, show="*", width=30)
        senha_entry.grid(row=4, column=1, sticky='w')

        #BOTAO QUE ENVIA AS INFORMAÇÕES DA CAIXA DE USUARIO E SENHA
        tk.Label(frame, text="").grid(row=5, column=0)
        login_button = tk.Button(frame, text="Entrar", font=('Arial', 10), command=preencher_campos, width=20)
        login_button.grid(row=6, columnspan=2)

        #ASSOCIA A TECLA ENTER DO TECLADO PARA APERTA O BOTAO DE LOGIN
        root.bind('<Return>', lambda event: preencher_campos())

        #INFORMA QUE OS DADOS NAO SERAO ARMAZENADOS NO SOFTWARE
        tk.Label(frame, text="").grid(row=7, column=0)
        tk.Label(frame, text="").grid(row=8, column=0)
        tk.Label(frame, text="O usuário e senha não serão armazenados em nenhum local durante a execução do script. Ambos serão utilizados apenas para o acesso ao site.", font=('Arial',15)).grid(row=9, columnspan=2)

        #INFORMA AUTOR DO SOFTWARE
        label = tk.Label(root, text="Script made by Theo Ravi.", font=('Arial', 10))
        label.place(relx=1.0, rely=1.0, anchor='se')

        root.mainloop()
        #ACESSA O SEI
        time.sleep(1)
        navegador.find_element(By.XPATH,'//*[@id="sbmAcessar"]').click()
    
    #CONDICAO DE ERRO PARA CASO O USUÁRIO ERRE O SEU LOGIN 
    except Exception:
        print('Ocorreu um erro, tente novamente.')
    #ENTRAR NA CAIXA DE PROCESSOS ATRIBUIDOS AO USUARIO
    if check_element_exists(By.XPATH,'//*[@id="divFiltro"]/div[2]/a'):
        navegador.find_element(By.XPATH,'//*[@id="divFiltro"]/div[2]/a').click()
        break
    #CASO NAO ENCONTRE O FILTRO DE PROCESSOS ELE INFOMA QUE NAO ENTROU NO SEI
    else:
        print('Usuário ou senha incorretos. Digite novamente')
  
#PEGA O CORPO DA TABELA NO NAVEGADOR
tbody = navegador.find_element(By.XPATH, '//*[@id="tblProcessosRecebidos"]/tbody')
#PEGA TODAS AS LINHAS (TR) DO CORPO DA TABELA, EXCETO A PRIMEIRA
trs = tbody.find_elements(By.TAG_NAME, 'tr')[1:]

#CRIA UM EXECUTOR DE THREADS
executor = ThreadPoolExecutor()
#APLICA A FUNÇÃO 'PROCESSA_TR' A CADA LINHA (TR) DA TABELA USANDO MULTITHREADING
resultados = list(executor.map(processa_tr, trs))
#AGUARDA TODAS AS THREADS TERMINAREM
executor.shutdown(wait=True)

#CRIA DUAS LISTAS PARA ARMAZENAR OS PROCESSOS QUE SERÃO ANALISADOS E CONCLUÍDOS
lista_processos = [r[0] for r in resultados if not r[1]]
lista_procConformes = [r[0] for r in resultados if r[1] and r[2]]

#INVERTE A ORDEM DOS PROCESSOS NAS LISTAS
lista_processos.reverse()

#IMPRIME A QUANTIDADE DE PROCESSOS QUE SERÃO ANALISADOS E CONCLUÍDOS
print("Quantidade de processos para analisar",len(lista_processos))
print("Quantidade de processos para concluir", len(lista_procConformes))

#MOSTRA OPÇÕES DE EXECUCAO PARA O USUARIO
nomeEstag=input('Insira seu nome completo: ')
nomeEstag_sem_acento = unidecode.unidecode(nomeEstag)
while True:
    try:
        opcoes = int(input("O que deseja fazer?\nDigite [1] para analisar processo e criar despacho.\nDigite [2] para concluir processos assinados.\nDigite [3] para atribuir processos para si.\nDigite [4] para cancelar a operação. "))
        if opcoes == 1:
            #EXECUTA FUNCAO PARA ANALISAR PROCESSO
            analisaProcesso()
            print("Análise finalizada.")
        elif opcoes == 2:
            #EXECUTA FUNCAO PARA CONCLUIR PROCESSO
            concluiProcesso()
        elif opcoes == 3:
            try:
                atribuicao()
            except:
                print("Ocorreu um erro!")
        else:
            #ENCERRA O PROGRAMA
            print("Encerrando programa")
            navegador.close()
            break
    #CONDICAO DE ERRO
    except ValueError:
        print("Opção inválida, tente novamente.")
