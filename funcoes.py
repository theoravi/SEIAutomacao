import numpy as np
import os
import pandas as pd
import pyautogui
import pyperclip
import time
import tkinter as tk
import undetected_chromedriver as uc
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tabulate import tabulate
from tkinter import scrolledtext
from webdriver_manager.chrome import ChromeDriverManager
import subprocess
from pywinauto import Application
from pywinauto.findwindows import find_windows, ElementAmbiguousError, ElementNotFoundError
import traceback

def preencher_campos(user_var, senha_var, navegador, root):
    global user_name
    user = user_var.get()
    senha = senha_var.get()
    user_name = user
    navegador.find_element(By.XPATH,'//*[@id="txtUsuario"]').clear()
    navegador.find_element(By.XPATH,'//*[@id="pwdSenha"]').clear()
    navegador.find_element(By.XPATH,'//*[@id="txtUsuario"]').send_keys(f"{user}")
    navegador.find_element(By.XPATH,'//*[@id="pwdSenha"]').send_keys(f"{senha}")
    root.destroy()

# FUNÇÃO PARA VERIFICAR A EXISTÊNCIA DE ELEMENTOS NA TELA UTILIZANDO O SELENIUM
def check_element_exists(by, value, navegador):
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
    nome_arquivo = f"logs/{nomeEstag}.txt"
    
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
    pedido_cancelamento = False
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
        anotacoes = tr.find_elements(By.XPATH, './/img[@src="svg/anotacao1.svg?18"]')
        #VERIFICA SE ENCONTROU ALGUM ÍCONE DE ANOTAÇÃO
        if anotacoes:
            #SE ENCONTROU, MARCA A VARIÁVEL 'POSSUI_ANOTACAO' COMO VERDADEIRA
            possui_anotacao = True
            #VERIFICA SE A ANOTAÇÃO CONTÉM O TEXTO 'AGUARDANDO ASSINATURA'
            if tr.find_elements(By.XPATH, ".//*[contains(translate(@onmouseover, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'aguardando assinatura')]"):
                #SE CONTÉM, MARCA A VARIÁVEL 'AGUARDANDO_ASSINATURA' COMO VERDADEIRA
                aguardando_assinatura = True
            if tr.find_elements(By.XPATH, ".//*[contains(translate(@onmouseover, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'pedido de cancelamento')]"):
                pedido_cancelamento = True
    #RETORNA AS INFORMAÇÕES DO PROCESSO COMO UMA TUPLA
    return (processo_texto.strip(), possui_anotacao, aguardando_assinatura, pedido_cancelamento)
  
#FUNCAO PARA EXIGENCIA DE INSERIR ANEXOS
def insira_anexo(processos, navegador):
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

    #ASSOCIA O ATALHO CTRL + ENTER À AÇAO DE ENVIAR EXIGENCIA
    root.bind('<Control-Return>', lambda event: get_text())

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
def erro_declaracao(processos, impProp, navegador, jaHomologado=False):
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

    #ASSOCIA O ATALHO CTRL + ENTER À AÇAO DE ENVIAR EXIGENCIA
    root.bind('<Control-Return>', lambda event: get_text())

    #INICIA O LOOP PRINCIPAL QUE MANTEM A JANELA ATIVA
    root.mainloop()
    #VERIFICA SE É IMPORTADO PRA USO PRÓPRIO
    #CASO SEJA IMPORTADO PRA USO PROPRIO E SEJA UM RADIO ELE PASSARÁ O MANUAL PARA RADIO E CORRELATOS
    #SE FOR DRONE PRA USO PROPRIO ELE PASSARA O MANUAL PARA DRONE
    
    if impProp == 'Sim':
        while True:
            try:
                radioamador = str(input('O produto é um rádio ou correlatos? Se sim digite [1], se for um Drone para uso próprio digite [2]: '))
                
                if radioamador == '1':
                    link = 'https://docs.google.com/document/d/1Y-8wGVnDMuku_Dmd5uT6hHCuj3ruXbuW3o-JFDWxog0/edit'
                    break
                elif radioamador == '2':
                    link = 'https://docs.google.com/document/d/1YsrMwMxyVysCJ9VhtGOOU1S1FlbmkRFaYSjojLxXVsQ/edit'
                    break
                else:
                    print("Opção inválida. Por favor, digite [1] para rádio ou correlatos, ou [2] para Drone para uso próprio.")
            except ValueError:
                print("Entrada inválida. Por favor, digite um número [1] para rádio ou correlatos, ou [2] para Drone para uso próprio.")
    
    #CASO SEJA UMA DECLARACAO APENAS PARA DRONE ELE IRÁ PASSAR DIRETAMENTE O MANUAL PARA DRONE
    else:
        link = 'https://docs.google.com/document/d/1YsrMwMxyVysCJ9VhtGOOU1S1FlbmkRFaYSjojLxXVsQ/edit'
    
    #TESTA QUAL TIPO DE EXIGENCIA É
    if jaHomologado:
        texto_padrao = f'''Prezado(a) Senhor(a),

    O requerimento realizado nesta Declaração de Conformidade possui um produto que já possui o selo de homologação da Anatel, sendo esse produto:

    {exig2}

    Como o produto deve ser homologado apenas uma vez, solicitamos que realize um novo processo SEI contendo apenas os produtos não homologados previamente.

    Favor seguir as seguintes orientações:
    1. Realizar um novo processo SEI, cumprindo com a(s) exigência(s) apontada(s) anteriormente. Seguindo o manual: {link}

    2. Para criar um novo processo SEI, solicitamos acessar o sistema Sei da Anatel: https://sei.anatel.gov.br/sei/controlador_externo.php?acao=usuario_externo_logar&id_orgao_acesso_externo=0

    Dessa maneira, devido a presença de exigência(s), informamos que o atual processo SEI nº {processos} será arquivado.

    FAVOR NÃO RESPONDER ESTE E-MAIL.



    Atenciosamente,

    ORCN - Gerência de Certificação e Numeração

    SOR - Superintendência de Outorga e Recursos à Prestação

    Anatel - Agência Nacional de Telecomunicações'''

    else:
        texto_padrao = f'''Prezado(a) Senhor(a),

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

    Anatel - Agência Nacional de Telecomunicações'''
        
    #INSERE ASSUNTO DO PROCESO
    navegador.find_element(By.ID, 'txtAssunto').send_keys(f"Processo SEI nº {processos} - Indeferido")
    #INSERE EMAIL COM CORPO DO EMAIL
    navegador.find_element(By.ID, 'txaMensagem').send_keys(texto_padrao)

def processo_errado(processo, navegador):
    #ENVIA TEXTO DE INPUT PARA A VARIAVEL
    def get_text():
        nonlocal exig2
        exig2 = text_area.get("1.0", tk.END).strip()
        root.destroy()

    #INICIALIZA A VARIAVEL LOCAL QUE RECEBERA O TEXTO DE INPUT
    exig2 = ""

    #CRIA A JANELA DA INTERFACE PARA CAIXA DE INPUT
    root = tk.Tk()
    root.title("Inserir o tipo de processo que o usuário abriu")

    #INSTRUCAO DO QUE O USUARIO DEVE FAZER
    label_instruction = tk.Label(root, text="Copie e cole aqui o tipo de processo no SEI")
    label_instruction.pack(padx=10, pady=5)

    #CRIA AREA DE TEXTO COM ROLAGEM
    text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=20)
    text_area.pack(padx=10, pady=10)

    #CRIA BOTAO DE RECEBER O TEXTO
    btn_get_text = tk.Button(root, text="Obter Texto", command=get_text)
    btn_get_text.pack(pady=5)

    #ASSOCIA O ATALHO CTRL + ENTER À AÇAO DE ENVIAR EXIGENCIA
    root.bind('<Control-Return>', lambda event: get_text())

    #INICIA O LOOP PRINCIPAL QUE MANTEM A JANELA ATIVA
    root.mainloop()

    texto_padrao = f'''Prezado(a) Senhor(a),

    Em atenção à demanda registrada no processo em referência, apresentamos as seguintes pendências observadas durante a análise da solicitação:

    1. O processo criado foi feito com a tipologia incorreta. Ao invés do tipo "{exig2}", o referente processo deveria ter sido aberto como "Certificação de Produto: Homologação de Drone";
    2. Foi aberto requerimento no documento incorreto. O documento certo seria de "Declaração de Conformidade - Drone";
    2.1. Não podem ser adicionadas informações pessoais nesse documento, como CPF, RG ou CNH.
    3. Para demais dúvidas quanto a como prosseguir com o requerimento, consulte o manual do usuário.

    Dessa maneira, informamos que o atual processo SEI nº {processo} será arquivado. 

    FAVOR NÃO RESPONDER ESTE E-MAIL. 

     

    Atenciosamente,

    ORCN - Gerência de Certificação e Numeração

    SOR - Superintendência de Outorga e Recursos à Prestação

    Anatel - Agência Nacional de Telecomunicações
    '''
    #INSERE ASSUNTO DO PROCESO
    navegador.find_element(By.ID, 'txtAssunto').send_keys(f"Processo SEI nº {processo} - Indeferido")
    #INSERE EMAIL COM CORPO DO EMAIL
    navegador.find_element(By.ID, 'txaMensagem').send_keys(texto_padrao)

def processo_reaberto(navegador, processo, nomeEstag):
    #VAI PARA A PÁGINA INICIAL DO PROCESSO
    vai_para_processo(navegador, processo)

    #CLICA NO INCONE DE INCLUIR DOCUMENTO
    navegador.switch_to.frame('ifrConteudoVisualizacao')
    clica_noelemento(navegador, By.XPATH,'//*[@id="divArvoreAcoes"]/a[1]')

    #CLICA NA OPCAO DE DESPACHO DECISORIO
    navegador.switch_to.frame('ifrVisualizacao')
    clica_noelemento(navegador, By.PARTIAL_LINK_TEXT,'Despacho Decisório')

    #SELECIONA TEXTO PADRAO
    clica_noelemento(navegador, By.XPATH,'//*[@id="divOptTextoPadrao"]/div')

    #ENVIA QUAL DESPACHO DECISORIO DEVE SER CRIADO
    navegador.find_element(By.XPATH,'//*[@id="txtTextoPadrao"]').send_keys('Despacho Decisório - Cancelamento de Homologação')

    #CLICA NA PRIMEIRA OPCAO
    clica_noelemento(navegador, By.XPATH,'//*[@id="divInfraAjaxtxtTextoPadrao"]/ul/li/a')

    #COLOCA O DESPACHO COMO PUBLICO
    clica_noelemento(navegador, By.XPATH,'//*[@id="divOptPublico"]/div')

    #SALVA DESPACHO
    clica_noelemento(navegador, By.ID,'btnSalvar')

    while True:
        try:
            confirmacao = str(input('Edite o documento da forma que preferir, salve e feche a janela.\nPara concluir digite [1], caso haja algum problema e queira parar o código, digite [2]: '))
            if confirmacao == '1':
                break
            elif confirmacao == '2':
                return
            else:
                print("Opção inválida, tente novamente!")
        except ValueError:
            print("Opção inválida, tente novamente!")

    navegador.switch_to.default_content()
    navegador.switch_to.frame('ifrArvore')
    lista_despachos = navegador.find_elements(By.PARTIAL_LINK_TEXT, 'Despacho Decisório')
    for i in lista_despachos:
        navegador.switch_to.default_content()
        navegador.switch_to.frame('ifrArvore')
        i.click()
        navegador.switch_to.default_content()
        navegador.switch_to.frame('ifrConteudoVisualizacao')
        navegador.switch_to.frame('ifrVisualizacao')
        time.sleep(1)
        if check_element_exists(By.XPATH, "//*[contains(text(), 'Revogar o')]", navegador):
            navegador.switch_to.default_content()
            navegador.switch_to.frame('ifrConteudoVisualizacao')
            print('Despacho Decisório de Cancelamento encontrado.\n')

            time.sleep(1)
            #CLICA NO ICONE DE LEGO
            clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/bloco_incluir_protocolo.svg?18')]")
            
            #SELECIONA BLOCO (SELECIONA O PRIMEIRO DESPACHO PARA DRONES APROVADOS QUE LER)
            time.sleep(1.5)
            navegador.switch_to.frame('ifrVisualizacao')
            select_element = navegador.find_element(By.ID, 'selBloco')
            select = Select(select_element)
            #PROCURA O PRIMEIRO BLOCO QUE TENHA O TEXTO "Despachos para Drones aprovados"
            for opcao in select.options:
                if "Despachos para Drones aprovados" in opcao.text:
                    select.select_by_visible_text(opcao.text)
                    break
            time.sleep(0.5)
            #CLICA NO BOTAO DE INCLUIR NO BLOCO
            clica_noelemento(navegador, By.XPATH,'//*[@id="sbmIncluir"]')
            #navegador.find_element(By.XPATH, '//*[@id="sbmIncluir"]').click()

    #VOLTA PARA PAGINA INICIAL DO PROCESSO
    vai_para_processo(navegador, processo)
    time.sleep(1)
    #ADICIONA NOTA PARA AGUARDAR ASSINATURA
    #CLICA NO ICONE DE ANOTACAO
    navegador.switch_to.default_content()
    navegador.switch_to.frame('ifrConteudoVisualizacao')
    clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/anotacao_cadastro.svg?18')]")
    #clica_noelemento(navegador, By.XPATH,'//*[@id="divArvoreAcoes"]/a[16]')
    time.sleep(0.1)
    #INSERE O TEXTO DA ANOTACAO
    navegador.switch_to.frame('ifrVisualizacao')
    navegador.find_element(By.ID, 'txaDescricao').send_keys('Pedido de Cancelamento')
    #SALVA O TEXTO
    navegador.find_element(By.NAME, 'sbmRegistrarAnotacao').click()
    #DA COMO APROVADO E ANOTA NO TXT DE PROCESSOS CONFORMES
    preenche_planilhageral(processo, nomeEstag, situacao="Cancelado")


#FUNCAO QUE PERMITE O USUARIO INSERIR CORPO DO EMAIL
def outro_erro(navegador):
    #INSERE SOMENTE ASSUNTO DO EMAIL
    assunto  = input("Insira o assunto do email: ")
    navegador.find_element(By.ID, 'txtAssunto').send_keys(assunto)

#FUNCAO QUE VERIFICA SE OS MODELOS INFORMADOS ESTAO NA LISTA DE DRONES CONFORMES
def verifica_conformidade(modelos, tabela_modelos, num=1):
    # Inicializa a lista com False
    checkexcel = [False] * len(modelos)

    # Comparação de modelos
    for modelo_solicitante in range(len(modelos)):
        for j in range(len(tabela_modelos)):
            try:
                if num == 0:
                    if str(modelos[modelo_solicitante]).lower().replace(' ','').strip('\n') == str(tabela_modelos[j]).lower().replace(' ',''):
                        checkexcel[modelo_solicitante] = True
                        break
                elif num == 1:
                    if str(modelos[modelo_solicitante]).lower().replace(' ','').replace('-','').strip('\n') == str(tabela_modelos[j]).lower().replace(' ','').replace('-',''):
                        checkexcel[modelo_solicitante] = True
                        break
            except KeyError as e:
                print(f"Erro ao acessar os índices: {e}")
                continue 
    return checkexcel

#FUNCAO QUE FORMATA A PLANILHA DE DRONES CONFORMES
def corrige_planilha(planilha, drones: bool):
    if drones:
        # Se for pra a planilha de drones:
        tabela = pd.read_excel(planilha, usecols=[2,3])
        tabela.columns = tabela.iloc[1]
        tabela = tabela.iloc[2:]
        tabela = tabela.reset_index(drop=True)
        # Se for para a de rádios:
    else:
        tabela = pd.read_excel(planilha)
        tabela.columns = tabela.iloc[1]
        tabela = tabela.iloc[2:]
        tabela = tabela.reset_index(drop=True)
        tabela = tabela['MODELO']
    return tabela

def muda_janela(janela: str):
    try:
        # Busca todas as janelas que contenham o título especificado
        matches = find_windows(title_re=f".*{janela}.*", backend="win32", visible_only=True)
        
        if not matches:
            raise ElementNotFoundError(f"Nenhuma janela encontrada com o título '{janela}'.")

        if len(matches) > 1:
            print(f"Mais de uma janela encontrada com o título '{janela}'. Selecionando a mais recente.")
        
        # Seleciona a última janela encontrada (mais recente/ativa)
        handle = matches[-1]
        app = Application(backend="win32").connect(handle=handle)
        window = app.window(handle=handle)

        # Garante que a janela será restaurada e focada
        if not window.is_active():
            # window.restore()
            window.set_focus()
        return True

    except ElementAmbiguousError:
        print(f"Conflito: mais de uma janela corresponde ao título '{janela}'.")
        # Lista as janelas e permite ao usuário escolher
        matches = find_windows(title_re=f".*{janela}.*", backend="win32", visible_only=False)
        for i, hwnd in enumerate(matches):
            print(f"[{i}] Handle: {hwnd}")
        
        escolha = input("Digite o número da janela que deseja focar: ").strip()
        if escolha.isdigit() and int(escolha) in range(len(matches)):
            app = Application().connect(handle=matches[int(escolha)])
            window = app.window(handle=matches[int(escolha)])
            window.set_focus()
            print("Janela selecionada com sucesso.")
            return True
        else:
            print("Opção inválida. Nenhuma ação foi tomada.")

    except ElementNotFoundError:
        while True:
            print(f"Parece que a janela '{janela}' não está visível para o programa.")
            opcao = input("Reabra-a e digite [1] para continuar ou [2] para encerrar: ").strip()
            if opcao == '1':
                muda_janela(janela)
                return True
            elif opcao == '2':
                return False
            else:
                print('Opção inválida!')

    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
        return False
    
def localiza_processo(processo):
    pyautogui.PAUSE = 0.7
    pyperclip.copy(processo)
    #PESQUISA PROCESSO NA PLANILHA
    pyautogui.hotkey('ctrl', 'l')
    time.sleep(0.3)
    #COLA NUMERO DO PROCESSO E APERTA ENTER PARA PESQUISAR O PROCESSO
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(0.3)
    pyautogui.press('esc')
    #COPIA O NUMERO DO PROCESSO NA CELULA PARA SABER SE ELE FOI ENCONTRADO CORRETAMENTE
    pyperclip.copy('nan')
    pyautogui.hotkey('ctrl','c')
    time.sleep(0.3)
    celula_encontrada = pyperclip.paste()
    celula_encontrada = celula_encontrada.replace('\n', '').strip('"')
    return celula_encontrada

#FUNCAO QUE PREENCHE A PLANILHA GERAL
def preenche_planilhageral(processo, nomeEstag, retido='', situacao='', codigo_rastreio='', nome_interessado=''):
    #ENCONTRA O ICONE DO EDGE E ABRE O NAVEGADOR DA PLANILHA
    mudou_janela = muda_janela('Distribuição Processo Drone.xlsx')
    if mudou_janela:
        celula_encontrada = localiza_processo(processo)
        print(f'Célula encontrada: {celula_encontrada}')
        while celula_encontrada != processo:
            muda_janela('theomation')
            if celula_encontrada == 'nan':
                input('Parece que o Excel não estava em foco na janela do Microsoft Edge. Volte para a página do app, coloque em foco (clicando em qualquer lugar da planilha) e pressione enter.')
            if celula_encontrada == 'Recuperando dados. Aguarde alguns segundos e tente cortar ou copiar novamente.':
                opcao = input('Parece que o Excel está lento. Caso o processo tenha sido encontrado, pressione [1]. Caso contrário, pressione enter para que o programa tente localizá-lo novamente')
                if opcao == '1':
                    muda_janela('Distribuição Processo Drone.xlsx')
                    break
            muda_janela('Distribuição Processo Drone.xlsx')
            celula_encontrada = localiza_processo(processo)
        if situacao == 'Cancelado':
            pyautogui.PAUSE = 0.3
            for i in range(6):
                pyautogui.press('right')

            pyperclip.copy(situacao)
            pyautogui.hotkey('ctrl', 'v')

        else:
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
            # chrome=pyautogui.locateOnScreen('imagensAut/chrome.png', confidence=0.7)
            # pyautogui.click(chrome)
            muda_janela('Google Chrome')
            muda_janela('theomation')


#FUNCAO QUE ESPERA UM ELEMENTO CARREGAR NA TELA E CLICA NELE
def clica_noelemento(navegador, modo_procura, element_id, tempo=5):
    # Espera até que o elemento seja carregado (exemplo: elemento localizado por ID)
    element = WebDriverWait(navegador, tempo).until(
        EC.element_to_be_clickable((modo_procura, element_id))
    )
    # Clica no elemento
    element.click()

def vai_para_processo(navegador, processo):
    navegador.switch_to.default_content()
    navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processo)
    elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
    elementos.send_keys(Keys.ENTER)

def armazena_nomesolicitante(navegador):
    navegador.switch_to.default_content()
    navegador.switch_to.frame('ifrArvore')
    navegador.find_element(By.PARTIAL_LINK_TEXT, 'Recibo Eletrônico').click()
    navegador.switch_to.default_content()
    #ENTRA NOS FRAMES QUE POSSUE O NOME DO SOLICITANTE
    navegador.switch_to.frame('ifrConteudoVisualizacao')
    navegador.switch_to.frame('ifrVisualizacao')
    #ARMAZENA NOME DO SOLICITANTE NA VARIAVEL
    nomeSol = navegador.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[1]/td[2]').text
    return nomeSol

def tira_restrito(navegador):
    navegador.switch_to.default_content()
    navegador.switch_to.frame('ifrArvore')
    arvoredocs = navegador.find_elements(By.CLASS_NAME, 'infraArvoreNo')

    #VERIFICA SE O DOCUMENTO ESTÁ RESTRITO, SE ESTIVER ELE IRÁ DEIXAR COMO PUBLICO UTILIZANDO O SIMBOLO DE RESTRITO COMO REFERENCIA
    elementos_com_src = navegador.find_elements(By.CSS_SELECTOR, "[src='svg/processo_restrito.svg?18']")
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

    for id in elementos_para_clicar:
        #ENCONTRA O ELEMENTO PARA CLICAR
        doc = navegador.find_element(By.ID, id)
        doc.click()
        time.sleep(1)
        navegador.switch_to.default_content()
        navegador.switch_to.frame('ifrConteudoVisualizacao')
        #CLICA NO SIMBOLO DE ALTERAR DOCUMENTO
        clica_noelemento(navegador, By.XPATH, '//*[@id="divArvoreAcoes"]/a[2]/img')
        time.sleep(0.7)
        navegador.switch_to.frame('ifrVisualizacao')
        #SELECIONA A OPCAO DE PUBLICO NO DOCUMENTO
        clica_noelemento(navegador, By.XPATH,'//*[@id="divOptPublico"]/div/label')
        #SALVA AS MUDANCAS
        clica_noelemento(navegador, By.ID,'btnSalvar')
        navegador.switch_to.default_content()
        navegador.switch_to.frame('ifrArvore')
    navegador.switch_to.default_content()

def abre_email(navegador, processo, processo_errado=False):
    navegador.switch_to.default_content()
    navegador.switch_to.frame('ifrArvore')
    
    if processo_errado:
        try:
            #CLICA NO REQUERIMENTO DE INFORMACOES SOBRE HOMOLOGACAO
            clica_noelemento(navegador, By.PARTIAL_LINK_TEXT, "Requerimento de", 1)
            processo_sem_email = False
        except:
            processo_sem_email = True
    else:
        #CLICA NA DECLARACAO DE CONFORMIDADE
        clica_noelemento(navegador, By.PARTIAL_LINK_TEXT, "Declaração de Conformidade")
    time.sleep(1)
    navegador.switch_to.default_content()
    navegador.switch_to.frame('ifrConteudoVisualizacao')
    navegador.switch_to.frame('ifrVisualizacao')

    #ARMAZENA O EMAIL DO SOLICITANTE
    if processo_errado and processo_sem_email:
        print("Não foi possivel obter o email do solicitante!")
        while True:
            emailSol = input("Digite o email do solicitante: ")
            opcao = input("Confirma o email digitado? Digite [1] para confirmar, [2] para digitar novamente ou [3] para cancelar: ")
            if opcao == '1':
                break   
            elif opcao == '2':  
                continue            
            elif opcao == '3':
                return
            else:
                print("Opção inválida! Tente novamente.")
    else:
        try:
            emailSol = navegador.find_element(By.XPATH, '/html/body/div[3]/a[1]').text
        except:
            print("Não foi possivel obter o email do solicitante!")
            while True:
                emailSol = input("Digite o email do solicitante: ")
                opcao = input("Confirma o email digitado? Digite [1] para confirmar, [2] para digitar novamente ou [3] para cancelar: ")
                if opcao == '1':
                    break   
                elif opcao == '2':  
                    continue            
                elif opcao == '3':
                    return
                else:
                    print("Opção inválida! Tente novamente.")


    #RETORNA PARA A PAGINA INICIAL DO PROCESSO
    vai_para_processo(navegador, processo)
    time.sleep(1)
    navegador.switch_to.frame('ifrConteudoVisualizacao')
    #CLICA NO ICONE DE ENVIO DE EMAIL
    clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/email_enviar.svg?18')]")
    #clica_noelemento(navegador, By.XPATH, '//*[@id="divArvoreAcoes"]/a[10]')
    time.sleep(1)
    #VAI PARA A JANELA MAIS RECENTE ABERTA
    navegador.switch_to.window(navegador.window_handles[-1])
    #time.sleep(1)
    #ABRE O DROPDOWN COM AS OPCOES DE EMAIL DA ANATEL
    select_element = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="selDe"]')))
    #select_element = navegador.find_element(By.XPATH, '//*[@id="selDe"]')
    select = Select(select_element)
    #SELECIONA O EMAIL DA ANATEL
    select.select_by_visible_text('ANATEL/E-mail de replicação <nao-responda@anatel.gov.br>')
    #ESCREVE O EMAIL DO SOLICITANTE
    navegador.find_element(By.XPATH, '//*[@id="s2id_autogen1"]').send_keys(emailSol)
    time.sleep(1.2)
    #CLICA NO EMAIL DO SOLICITANTE
    clica_noelemento(navegador, By.CLASS_NAME, 'select2-result-label')
    #MOSTRA AS OPCOES DE EXIGENCIA
    #EXECUTA A CONDICAO PARA CADA EXIGENCIA

#FUNCAO QUE ESPERA UM ELEMENTO CARREGAR NA TELA E ENVIA TEXTO NELE
def sendkeys_elemento(navegador, modo_procura, element_id, texto):
# Espera até que o elemento seja carregado (exemplo: elemento localizado por ID)
    try:
        element = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((modo_procura, element_id))
        )
        # Envia texto no elemento
        element.send_keys(texto)
    except TimeoutException:
        print(f"Elemento {element_id} não foi carregado no tempo esperado")
        return TimeoutException

#FUNCAO PARA INSERIR O EMAIL
def endereco_email(endereco, navegador):
    navegador.find_element(By.XPATH, '//*[@id="s2id_autogen1"]').send_keys(endereco)
    time.sleep(1)
    #CLICA NO EMAIL DO SOLICITANTE
    clica_noelemento(navegador, By.XPATH, '//*[@id="select2-drop"]/ul')
    time.sleep(0.2)

def envia_email(navegador, janela_principal, processo, nomeEstag, impProp, tipoExig):
    if tipoExig == 1:
        #PUXA FUNCAO INSIRA_ANEXO CONTENDO O CORPO DE EMAIL PEDINDO ANEXOS
        insira_anexo(processo, navegador)
    elif tipoExig == 2:
        #PUXA FUNCAO ERRO_DECLARACAO CONTENDO O CORPO DE EMAIL PEDINDO PARA ABRIR NOVO PROCESSO SEI
        erro_declaracao(processo, impProp, navegador, jaHomologado=False)
    elif tipoExig == 3:
        #PUXA FUNCAO ERRO_DECLARACAO CONTENDO O CORPO DE EMAIL PEDINDO PARA ABRIR NOVO PROCESSO SEI
        erro_declaracao(processo, impProp, navegador, jaHomologado=True)
    elif tipoExig == 4:
        #PUXA FUNCAO OUTRO_ERRO EM QUE USUARIO INSERE O EMAIL COMO QUISER
        outro_erro(navegador)
    else:
        processo_errado(processo, navegador)

    #PEDE PARA USUARIO CONFERIR SE ESTÁ TUDO CERTO COM O EMAIL
    muda_janela('theomation')
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
        #VOLTA PARA A PAGINA INICIAL DO PROCESSO
        vai_para_processo(navegador, processo)
        time.sleep(1)
        navegador.switch_to.frame('ifrConteudoVisualizacao')
        #CLICA NO ICONE DE TAG
        time.sleep(0.5)
        navegador.find_element(By.XPATH, "//img[contains(@src, 'svg/marcador_gerenciar.svg?18')]").click()
        navegador.switch_to.frame('ifrVisualizacao')
        try:    
            time.sleep(1)
            navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
        except:
            time.sleep(0.3)
            clica_noelemento(navegador, By.XPATH,'//*[@id="tblMarcadores"]/tbody/tr[2]/td[6]/a[2]/img')
            #navegador.find_element(By.XPATH, '//*[@id="tblMarcadores"]/tbody/tr[2]/td[6]/a[2]/img').click()
            alert.accept()
            time.sleep(0.3)
            clica_noelemento(navegador, By.XPATH,'//*[@id="btnAdicionar"]')
            #navegador.find_element(By.XPATH, '//*[@id="btnAdicionar"]').click()
            #time.sleep(0.5)
            clica_noelemento(navegador, By.XPATH,'//*[@id="selMarcador"]/div/a')
            #navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
        
        time.sleep(0.5)
        if tipoExig == 1:
            #CLICA NA TAG DE INTERCORRENTE
            navegador.find_element(By.PARTIAL_LINK_TEXT, 'Processo SEI ORCN - Intercorrente').click()
        else:
            #CLICA NA TAG DE PENDENCIA
            navegador.find_element(By.PARTIAL_LINK_TEXT, 'Processo SEI ORCN - Pendência').click()

        #PEDE O TEXTO DA TAG
        muda_janela('theomation')
        textoTag = input("Insira o texto da tag: ")
        textoFinaltag = nomeEstag+"\n"+textoTag
        #COLOCA O TEXTO NA TAG
        navegador.find_element(By.XPATH, '//*[@id="txaTexto"]').send_keys(textoFinaltag)
        #SALVA TAG
        navegador.find_element(By.XPATH, '//*[@id="sbmSalvar"]').click()
        return True
    else:
        #CANCELA ENVIO DO EMAIL SE HOUVER ALGUM ERRO
        print("Faça as alterações necessárias mais tarde.")
        navegador.find_element(By.XPATH, '//*[@id="btnCancelar"]').click()
        navegador.switch_to.window(janela_principal)
        return False

def fazListasProcessos(navegador):
    #PEGA O CORPO DA TABELA NO NAVEGADOR
    tbody = navegador.find_element(By.XPATH, '//*[@id="tblProcessosRecebidos"]/tbody')
    #PEGA TODAS AS LINHAS (TR) DO CORPO DA TABELA, EXCETO A PRIMEIRA
    trs = tbody.find_elements(By.TAG_NAME, 'tr')[1:]

    lista_processos = []
    lista_procConformes = []
    lista_procCancelamento = []

    for tr in trs:
        processo_texto, possui_anotacao, aguardando_assinatura, pedido_cancelamento = processa_tr(tr)
        if aguardando_assinatura:
            lista_procConformes.append(processo_texto) 
        if pedido_cancelamento:
            lista_procCancelamento.append(processo_texto)
        elif not possui_anotacao:
            lista_processos.append(processo_texto) 

    #INVERTE A ORDEM DOS PROCESSOS NAS LISTAS
    lista_processos.reverse()
    return lista_procConformes, lista_procCancelamento, lista_processos

#FUNCAO QUE ABRE O CHROME E O EDGE
def abreChromeEdge():
    #INSTALA O CHROME DRIVEr MAIS ATUALIZADO
    # Desativa o bloqueio de pop-ups
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-popup-blocking") 
    #DEFINE O TEMPO DE EXECUÇÃO PARA CADA COMANDO DO PYAUTOGUI
    pyautogui.PAUSE = 0.7
    #INICIA O NAVEGADOR
    # Caminho do ChromeDriver local
    # chrome_driver_path = r"main\chromedriver-win64\chromedriver.exe"  # Altere para o caminho correto do seu ChromeDriver
    # Configura o serviço do ChromeDriver
    servico = Service(ChromeDriverManager().install())
    # Inicia o navegador Chrome
    navegador = webdriver.Chrome(service=servico, options=options)
    navegador.maximize_window()
    #ENTRA NO SEI
    navegador.get('https://sei.anatel.gov.br/')
    #LOCALIZA O ICONE DO EDGE E ABRE A PLANILHA GERAL
    #FOI UTILIZADO O PYPERCLIP PARA EVITAR QUALQUER ERRO NA HORA DE COLAR O URL DA PLANILHA
    time.sleep(1)
    # Tenta selecionar o Edge
    try:
        # Busca todas as janelas que contenham o título especificado
        matches = find_windows(title_re=f".*Distribuição Processo Drone.xlsx.*", backend="win32", visible_only=True)
        
        if not matches:
            raise ElementNotFoundError(f"Nenhuma janela encontrada com o título 'Distribuição Processo Drone.xlsx'.")

        if len(matches) > 1:
            print(f"Mais de uma janela encontrada com o título 'Distribuição Processo Drone.xlsx'. Selecionando a mais recente.")
        
        # Seleciona a última janela encontrada (mais recente/ativa)
        handle = matches[-1]
        app = Application(backend="win32").connect(handle=handle)
        window = app.window(handle=handle)

        # Garante que a janela será restaurada e focada
        if not window.is_active():
            # window.restore()
            window.set_focus()
    except ElementNotFoundError:
        # Abre o Edge
        subprocess.Popen(["C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"])
        time.sleep(2)
        sitePlan = 'https://anatel365.sharepoint.com/:x:/r/sites/lista.orcn/_layouts/15/Doc.aspx?sourcedoc=%7B4130A4D6-7F00-45D4-A328-ED0866A62335%7D&file=Distribui%C3%A7%C3%A3o%20Processo%20Drone.xlsx&action=default&mobileredirect=true'
        pyperclip.copy(sitePlan)
        pyautogui.hotkey('ctrl', 'l')
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
    # chrome=pyautogui.locateOnScreen('imagensAut/chrome.png', confidence=0.7)
    # pyautogui.click(chrome)
    muda_janela('Google Chrome')
    return navegador

#FUNCAO QUE PEDE A SENHA E O USUARIO
def iniciaJanela(navegador):
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
    login_button = tk.Button(frame, text="Entrar", font=('Arial', 10), command=lambda: preencher_campos(user_var, senha_var, navegador, root), width=20)
    login_button.grid(row=6, columnspan=2)

    #ASSOCIA A TECLA ENTER DO TECLADO PARA APERTA O BOTAO DE LOGIN
    root.bind('<Return>', lambda event: preencher_campos(user_var, senha_var, navegador, root))

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
    navegador.find_element(By.XPATH, '//*[@id="sbmAcessar"]').click()

#FUNCAO QUE ANALISA TODOS OS PROCESSOS NA CAIXA DO USUARIO
def analisaListaDeProcessos(navegador, lista_processos, nomeEstag, planilhaDrones, planilhaRadios):
    drone_modelos = corrige_planilha(drones=True, planilha=planilhaDrones)
    radio_modelos = corrige_planilha(drones=False, planilha=planilhaRadios)

    for processos in lista_processos[:]:
        try:
            analisa(navegador, processos, nomeEstag, drone_modelos, radio_modelos)
        except Exception as e:
            print("Ocorreu um erro: ", e)
            print("Traceback: ")
            traceback.print_exc()
        #REMOVE O PROCESSO ANALISADO DA LISTA
        lista_processos.remove(processos)
        while True:
            opcao = str(input("Caso deseje analisar o próximo processo digite [1], caso contrario, digite [2]: "))
            if opcao == '2':
                return
            elif opcao == '1':
                break
            else:
                print("Alternativa inválida")
    return

#FUNCAO QUE ANALISA UM PROCESSO ESPECIFICO
def analisaApenasUmProcesso(navegador, nomeEstag, planilhaDrones, planilhaRadios):
    drone_modelos = corrige_planilha(drones=True, planilha=planilhaDrones)
    radio_modelos = corrige_planilha(drones=False, planilha=planilhaRadios)

    while True:
        #PEDE O PROCESSO A SER ANALISADO
        processo = str(input("Digite o número do processo: " ))
        opcao = str(input("Caso deseje analisar este processo, digite [1], caso contrario digite [2]: "))
        if opcao == '1':
            try:
                analisa(navegador, processo, nomeEstag, drone_modelos, radio_modelos)
            except Exception as e:
                print("Ocorreu um erro: ", e)
                print("Traceback: ")
                traceback.print_exc()
        elif opcao == '2':
            return
        else:
            print("Alternativa inválida")
        while True:
            analisar = str(input("Caso deseje analisar mais um processo digite [1], caso contrario, digite [2]: "))
            if analisar == '2':
                return
            elif analisar == '1':
                break
            else:
                print("Alternativa inválida")

#FUNCAO PARA ANALISAR PROCESSO
def analisa(navegador, processo, nomeEstag, drone_modelos, radio_modelos):
    #VARIAVEL UTILIZADA PARA O SELENIUM RETORNAR PARA A JANELA PRINCIPAL
    janela_principal = navegador.current_window_handle
# try:
    #FAZ UM LOG DO PROCESSO
    escrever_informacoes(processo, nomeEstag)
    #INICIA A VARIAVEL ZERADA
    impProp=''
    situacao=''
    #PESQUISA NUMERO DO PROCESSO NA CAIXA DE PESQUISA DO SEI PARA ACESSAR O PROCESSO
    navegador.switch_to.default_content()
    navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processo)
    elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
    elementos.send_keys(Keys.ENTER)
    #TEMPO DEFINIDO PARA QUE A PAGINA CARREGUE
    time.sleep(1)
    #ENTRA NO FRAME QUE CONTÉM OS DOCUMENTOS DOS PROCESSOS
    #A PAGINA DOS PROCESSOS SAO DIVIDIDOS EM FRAMES
    navegador.switch_to.frame('ifrArvore')
    #VERIFICA SE O PROCESSO JA FOI DESPACHADO, CASO TENHA SIDO ELE PULA O PROCESSO
    if check_element_exists(By.PARTIAL_LINK_TEXT, 'Despacho Decisório', navegador) and not (check_element_exists(By.XPATH, "//img[contains(@src, 'svg/marcador_verde.svg?18')]", navegador) or check_element_exists(By.XPATH, "//img[contains(@src, 'svg/marcador_preto.svg?18')]", navegador)):
        print(f"O processo {processo} já foi despachado!")
        return
    
    elif check_element_exists(By.PARTIAL_LINK_TEXT, 'Despacho Decisório', navegador) and (check_element_exists(By.XPATH, "//img[contains(@src, 'svg/marcador_verde.svg?18')]", navegador) or check_element_exists(By.XPATH, "//img[contains(@src, 'svg/marcador_preto.svg?18')]", navegador)):
        tipo_processo = 'Reaberto'
        print(f"O processo {processo} foi reaberto!")
        while True:
            try:
                confirmacao = str(input('Caso este seja um PEDIDO DE CANCELAMENTO, digite [1] para gerar o despacho. Caso contrário digite [2]: '))
                if confirmacao == '1':
                    processo_reaberto(navegador, processo, nomeEstag)
                    return
                elif confirmacao == '2':
                    return
                else:
                    print("Opção inválida, tente novamente!")
            except ValueError:
                print("Opção inválida, tente novamente!")
    
    #CONFERE SE EXISTE A DECLARACAO DE CONFORMIDADE OU UMA PASTA
    # if check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade", navegador):
    #     navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade").click()
    #     time.sleep(0.3)
    if check_element_exists(By.XPATH, '//*[@id="spanPASTA1"]', navegador):
        navegador.find_element(By.XPATH, '//*[@id="spanPASTA1"]').click()
        time.sleep(1)
        navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade").click()
        time.sleep(0.3)

    #CONFERE QUAL É O TIPO DE PROCESSO
    if check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Drone", navegador):
        tipo_processo = "Drone"
    elif check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Importado Uso Próprio", navegador):
        tipo_processo = "Importado"
    else:
        while True:
            try:
                print("Não foi possível encontrar a Declaração de Conformidade.")
                confirmacao = str(input('Caso este seja um processo CRIADO COM O TIPO ERRADO, digite [1] para enviar exigência, senão, digite [2] para pular o processo: '))
                if confirmacao == '1':
                    abre_email(navegador, processo, processo_errado=True)
                    email_enviado = envia_email(navegador, janela_principal, processo, nomeEstag, impProp, 5)
                    vai_para_processo(navegador, processo)
                    nomeSol = armazena_nomesolicitante(navegador)
                    if email_enviado:
                        preenche_planilhageral(processo, nomeEstag, 'Não', 'Cancelado', '', nomeSol)
                    return
                elif confirmacao == '2':
                    return
                else:
                    print("Opção inválida, tente novamente!")
            except ValueError:
                print("Opção inválida, tente novamente!")
    #ESCREVER AS INFORMAÇÕES DO PROCESSO NO CONSOLE
    #ACESSA O RECIBO PARA PEGAR NOME DO SOLICITANTE
    # navegador.switch_to.frame('ifrArvore')
    # navegador.find_element(By.PARTIAL_LINK_TEXT, 'Recibo Eletrônico').click()
    # navegador.switch_to.default_content()
    # #ENTRA NOS FRAMES QUE POSSUE O NOME DO SOLICITANTE
    # navegador.switch_to.frame('ifrConteudoVisualizacao')
    # navegador.switch_to.frame('ifrVisualizacao')
    # #ARMAZENA NOME DO SOLICITANTE NA VARIAVEL
    # nomeSol = navegador.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[1]/td[2]').text
    
    nomeSol = armazena_nomesolicitante(navegador)
    
    #RETORNA AO FRAME DEFAULT
    navegador.switch_to.default_content()
    #COLETAR DADOS DA DECLARAÇÃO DE CONFORMIDADE
    #RETORNA AO FRAME QUE CONTÉM OS DOCUMENTOS
    navegador.switch_to.frame('ifrArvore')

    if tipo_processo == "Drone":
        #CLICA NA DECLARACAO DE CONFORMIDADE
        navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Drone").click()
        time.sleep(1)
    elif tipo_processo == "Importado":
        #CLICA NA DECLARACAO DE CONFORMIDADE
        navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Importado Uso Próprio").click()
        time.sleep(1)

    #RETORNA AO FRAME DEFAULT E ENTRA NOS FRAMES QUE CONTÉM A VISUALIZACAO DA DECLARACAO
    navegador.switch_to.default_content()
    navegador.switch_to.frame('ifrConteudoVisualizacao')
    navegador.switch_to.frame('ifrVisualizacao')
    #EXIBE INFORMACOES DO PROCESSO
    print("--------------------------------------------------------------------------")
    print('--------------------------------------------------------------------------')
    print(f"Processo: {processo}")

    #COLETAR TEXTO NOME DO INTERESSADO
    if tipo_processo == "Drone":
        nome_interessado = navegador.find_element(By.XPATH, '/html/body/table[1]/tbody/tr/td').text
    elif tipo_processo == "Importado":
        nome_interessado = navegador.find_element(By.XPATH, '/html/body/table[1]/tbody/tr/td/p').text

    #VERIFICA SE O CAMPO ESTÁ VAZIO, CASO ESTEJA DEFINIRA CM NAO INFORMADO
    if not nome_interessado.strip():
        nome_interessado = 'Não informado'
        print("NOME DO INTERESSADO NÃO INFORMADO")
    #EXIBE NOMES DO INTERESSADO E SOLICITANTE
    print("Interessado: ",nome_interessado)
    print("Solicitante: ",nomeSol)
    print('\n')

    #CHECKBOX DE QUEM ESTÁ FAZENDO O PETICIONAMENTO
    if tipo_processo == "Drone":
        checkbox=navegador.find_element(By.XPATH,'/html/body/table[2]/tbody/tr[1]/td[1]').text
        checkbox2=navegador.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[2]/td[1]').text
    elif tipo_processo == "Importado":
        checkbox=navegador.find_element('xpath','/html/body/table[2]/tbody/tr[1]/td[1]/p').text
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

    if tipo_processo == "Importado":
        try:
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
        except Exception as error:
            print(f"Ocorreu um erro: {error}")
            print("É possível que o solicitante tenha editado o layout da tabela. Confira manualmente os modelos.\n")

    #CONTA QUANTIDADE DE LINHAS DA TABELA SOBRE O PRODUTO
    if tipo_processo == "Drone":
        try:
            quantidade_linhas = len(navegador.find_elements(By.XPATH, '/html/body/table[3]/tbody/tr'))

            #VERIFICA CAMPOS QUE CONTÉM AS INFORMAÇÕES SOBRE O PRODUTO
            #ARMAZENA OS MODELOS INFORMADOS PARA VERIFICAR SE ESTÃO NA PLANILHA DE DRONES CONFORMES
            modelos = []
            for i in range(1, quantidade_linhas, 1):
                linha = navegador.find_element('xpath', f'/html/body/table[3]/tbody/tr[{i}]')
                celulas = linha.find_elements(By.TAG_NAME,'td')
                modelos_produto = [celula.text for celula in celulas]
                modelos.append(modelos_produto)
        except Exception as error:
            print(f"Ocorreu um erro: {error}")
            print("É possível que o solicitante tenha editado o layout da tabela. Confira manualmente os modelos.\n")
    elif tipo_processo == "Importado":
        try:
            quantidade_linhas = len(navegador.find_elements(By.XPATH, '/html/body/table[4]/tbody/tr'))

            #VERIFICA CAMPOS QUE CONTÉM AS INFORMAÇÕES SOBRE O PRODUTO
            modelos = []
            for i in range(1, quantidade_linhas, 1):
                linha = navegador.find_element('xpath', f'/html/body/table[4]/tbody/tr[{i}]')
                celulas = linha.find_elements(By.TAG_NAME,'td')
                modelos_produto = [celula.text for celula in celulas]
                modelos.append(modelos_produto)
        except Exception as error:
            print(f"Ocorreu um erro: {error}")
            print("É possível que o solicitante tenha editado o layout da tabela. Confira manualmente os modelos.\n")

    #CRIA DATAFRAME PARA PRINTAR AS INFORMAÇÕES DO PRODUTO JÁ TRATADAS
    df = pd.DataFrame(modelos)
    try:
        if tipo_processo == "Drone":
            df.columns = ['Modelo', 'Nome Comercial', 'Número de Série (incluindo rádio controle e óculos)']
        elif tipo_processo == "Importado":
            df.columns = ['Modelo', 'Nome Comercial', 'Número de Série']
        df = df[1:]
        df = df.replace(' ', np.nan)
        df = df.dropna(how='all')
        data = df.values.tolist()
        headers = df.columns.tolist()
        print(tabulate(data, headers=headers, tablefmt='pretty'))
        print('\n')

        modelos = df['Modelo']
        modelos = modelos.reset_index(drop=True)

        if tipo_processo == 'Importado':
            #VERIFICA SE O MODELO DO RADIO CONTROLE ESTA NA PLANILHA DE RADIOS CONFORMES
            checkexcel = verifica_conformidade(modelos, radio_modelos)
            for i in range(len(checkexcel)):
                if checkexcel[i] == True:
                    print(f"O modelo {modelos[i]} está na lista de rádios conformes.")
                else:
                    print(f"O modelo {modelos[i]} não se encontra na lista de rádios conformes")
            print('\n')

        elif tipo_processo == "Drone":
            #VERIFICA SE O MODELO DO DRONE E RADIO CONTROLE ESTA NA PLANILHA DE DRONES CONFORMES
            drone_modelos['MODELO_CORRIGIDO'] = drone_modelos['MODELO'].apply(lambda x: str(x).lower().replace(' ', '').replace('-', ''))
            for mod in modelos:
                try:
                    mod2 = mod.lower().replace(' ','').replace('-','').strip('\n')
                    if mod2 in drone_modelos['MODELO_CORRIGIDO'].values:    
                        linha_nome = drone_modelos['MODELO_CORRIGIDO'] == mod2
                        nome_comercial = drone_modelos.loc[linha_nome, 'NOME COMERCIAL'].values[0]
                        print(f"O modelo '{mod}' está na lista de drones conformes. Seu nome comercial é '{nome_comercial}'")
                    else:
                        print(f"O modelo '{mod}' não se encontra na lista de drones conformes")
                except:
                    print("Este modelo está vazio, não foi possível fazer a verificação de conformidade")
            print('\n')

    except Exception as error:
        print(f"Ocorreu um erro: {error}")
        # traceback.print_exc()
        # print("\n")
        print("É possível que o solicitante tenha editado o layout da tabela. Confira manualmente os modelos.\n")
    #EXIBE CÓDIGO DE RASTREIO
    #LINHA ESPECIFICA ESCREVE NA PLANILHA SE O PRODUTO ESTA RETIDO E, CASO ESTEJA, ESCREVE O CODIGO DE RASTREIO
    if tipo_processo == "Drone":
        n_serie = navegador.find_element(By.XPATH, '/html/body/table[3]/tbody/tr[2]/td[3]').text
        codigo_rastreio = navegador.find_element(By.XPATH,'/html/body/table[4]/tbody/tr/td[2]').text
    elif tipo_processo == "Importado":
        n_serie = navegador.find_element(By.XPATH, '/html/body/table[4]/tbody/tr[2]/td[3]').text
        codigo_rastreio = navegador.find_element(By.XPATH,'/html/body/table[5]/tbody/tr/td[2]').text
    #VERIFICA SE HÁ ALGUM TEXTO NA TABELA COM O CÓDIGO DE RASTREIO
    
    #RETIRA O RESTRITO DOS DOCUMENTOS
    tira_restrito(navegador)
    
    #SE HOUVER ALGUM TEXTO ELE DA COMO RETIDO, CASO NAO TENHA TEXTO ELE DA COMO NAO RETIDO
    if not codigo_rastreio.strip():
        print("Produto não retido ou código de rastreio não informado.")
        retido='Não'
    else:
        if not n_serie.strip():
            print(f'O código de rastreio é: {codigo_rastreio}.')
            retido='Sim'
        elif n_serie.strip():
            print('Confira se o produto está retido.')
            while True:
                verifica_retido = str(input('O produto está retido? Se sim digite [1], se não digite [2]: '))
                if verifica_retido == '1':
                    retido='Sim'
                    print('O código de rastreio é:', codigo_rastreio)
                    break
                elif verifica_retido == '2':
                    retido='Não'
                    break
                else:
                    print("Alternativa invalida")
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
            confirmacao = str(input('Caso esteja tudo certo e NÃO HAJA UM DESPACHO JÁ CRIADO digite [1] para continuar, senão, digite [2]: '))
            if (nome_interessado == 'Não informado' or '') and confirmacao == '1':
                print("Nome do interessado não informado, não é possível criar despacho, selecione outra opção!")
            elif confirmacao == '1':
                break
            elif confirmacao == '2':
                break
            else:
                print("Opção inválida, tente novamente!")
        except ValueError:
            print("Opção inválida, tente novamente!")
    if confirmacao == '1':
        navegador.switch_to.default_content()
        navegador.switch_to.frame('ifrArvore')
        #ENTRA NA PAGINA INICIAL DO PROCESSO
        #VERIFICA SE E UM PROCESSO DE DRONE OU IMPORTADO PARA USO PROPRIO
        if check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Drone", navegador):
            navegador.switch_to.default_content()
            navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processo)
            elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
            elementos.send_keys(Keys.ENTER)
            #CRIA DESPACHO
            navegador.switch_to.frame('ifrConteudoVisualizacao')
            #time.sleep(2)
            #CLICA NO INCONE DE INCLUIR DOCUMENTO
            clica_noelemento(navegador, By.XPATH,'//*[@id="divArvoreAcoes"]/a[1]')
            #navegador.find_element(By.XPATH,'//*[@id="divArvoreAcoes"]/a[1]').click()
            #time.sleep(2)
            #CLICA NA OPCAO DE DESPACHO DECISORIO
            navegador.switch_to.frame('ifrVisualizacao')
            clica_noelemento(navegador, By.PARTIAL_LINK_TEXT,'Despacho Decisório')
            #navegador.find_element(By.XPATH,'//*[@id="tblSeries"]/tbody/tr[16]/td/a[2]').click()
            #time.sleep(2)
            #SELECIONA TEXTO PADRAO
            clica_noelemento(navegador, By.XPATH,'//*[@id="divOptTextoPadrao"]/div')
            #navegador.find_element(By.XPATH,'//*[@id="divOptTextoPadrao"]/div').click()
            #time.sleep(2)
            #ENVIA QUAL DESPACHO DECISORIO DEVE SER CRIADO
            navegador.find_element(By.XPATH,'//*[@id="txtTextoPadrao"]').send_keys('Despacho Decisório de Homologação Drones')
            #time.sleep(2)
        #VERIFICA SE E UM PROCESSO DE DRONE OU IMPORTADO PARA USO PROPRIO
        elif check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Importado Uso Próprio", navegador):
            impProp='Sim'
            navegador.switch_to.default_content()
            navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processo)
            elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
            elementos.send_keys(Keys.ENTER)
            #CRIA DESPACHO
            navegador.switch_to.frame('ifrConteudoVisualizacao')
            #time.sleep(1)
            #CLICA NO INCONE DE INCLUIR DOCUMENTO
            clica_noelemento(navegador, By.XPATH,'//*[@id="divArvoreAcoes"]/a[1]')
            #navegador.find_element(By.XPATH,'//*[@id="divArvoreAcoes"]/a[1]').click()
            ##time.sleep(1)
            #CLICA NA OPCAO DE DESPACHO DECISORIO
            navegador.switch_to.frame('ifrVisualizacao')
            clica_noelemento(navegador, By.PARTIAL_LINK_TEXT,'Despacho Decisório')
            #navegador.find_element(By.XPATH,'//*[@id="tblSeries"]/tbody/tr[16]/td/a[2]').click()
            #SELECIONA TEXTO PADRAO
            clica_noelemento(navegador, By.XPATH,'//*[@id="divOptTextoPadrao"]/div')
            #navegador.find_element(By.XPATH,'//*[@id="divOptTextoPadrao"]/div').click()
            #ENVIA QUAL DESPACHO DECISORIO DEVE SER CRIADO
            navegador.find_element(By.XPATH,'//*[@id="txtTextoPadrao"]').send_keys('Despacho Decisório de Homologação não licenciados')
        #time.sleep(1)
        #CLICA NA PRIMEIRA OPCAO
        clica_noelemento(navegador, By.XPATH,'//*[@id="divInfraAjaxtxtTextoPadrao"]/ul/li/a')
        #navegador.find_element(By.XPATH,'//*[@id="divInfraAjaxtxtTextoPadrao"]/ul/li/a').click()
        #COLOCA O DESPACHO COMO PUBLICO
        clica_noelemento(navegador, By.XPATH,'//*[@id="divOptPublico"]/div')
        #navegador.find_element(By.XPATH,'//*[@id="divOptPublico"]/div').click()
        #SALVA DESPACHO
        clica_noelemento(navegador, By.ID,'btnSalvar')
        #navegador.find_element(By.ID,'btnSalvar').click()
        time.sleep(2)
        navegador.switch_to.window(navegador.window_handles[-1])
        navegador.close()
        time.sleep(0.7)
        #MUDA PARA JANELA PRINCIAPL DO PROGRAMA
        navegador.switch_to.window(janela_principal)
        #INCLUI DESPACHO NO BLOCO
        navegador.switch_to.default_content()
        navegador.switch_to.frame('ifrArvore')
        #CLICA NO DESPACHO DECISORIO
        clica_noelemento(navegador, By.PARTIAL_LINK_TEXT,"Despacho Decisório")
        #navegador.find_element(By.PARTIAL_LINK_TEXT, "Despacho Decisório").click()
        navegador.switch_to.default_content()
        navegador.switch_to.frame('ifrConteudoVisualizacao')
        time.sleep(1)
        #CLICA NO ICONE DE LEGO
        clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/bloco_incluir_protocolo.svg?18')]")
        #navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[8]').click()
        #SELECIONA BLOCO (SELECIONA O PRIMEIRO DESPACHO PARA DRONES APROVADOS QUE LER)
        time.sleep(1.5)
        navegador.switch_to.frame('ifrVisualizacao')
        select_element = navegador.find_element(By.ID, 'selBloco')
        select = Select(select_element)
        #PROCURA O PRIMEIRO BLOCO QUE TENHA O TEXTO "Despachos para Drones aprovados"
        for opcao in select.options:
            if "Despachos para Drones aprovados" in opcao.text:
                select.select_by_visible_text(opcao.text)
                break
        time.sleep(0.5)
        #CLICA NO BOTAO DE INCLUIR NO BLOCO
        clica_noelemento(navegador, By.XPATH,'//*[@id="sbmIncluir"]')
        #navegador.find_element(By.XPATH, '//*[@id="sbmIncluir"]').click()
        #VOLTA PARA PAGINA INICIAL DO PROCESSO
        navegador.switch_to.default_content()
        navegador.find_element(By.ID,'txtPesquisaRapida').send_keys(processo)
        elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
        elementos.send_keys(Keys.ENTER)
        navegador.switch_to.frame('ifrConteudoVisualizacao')
        time.sleep(1)
        #ADICIONA NOTA PARA AGUARDAR ASSINATURA
        #CLICA NO ICONE DE ANOTACAO
        clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/anotacao_cadastro.svg?18')]")
        #clica_noelemento(navegador, By.XPATH,'//*[@id="divArvoreAcoes"]/a[16]')
        time.sleep(0.1)
        #INSERE O TEXTO DA ANOTACAO
        navegador.switch_to.frame('ifrVisualizacao')
        navegador.find_element(By.ID, 'txaDescricao').send_keys('Aguardando assinatura.')
        #SALVA O TEXTO
        navegador.find_element(By.NAME, 'sbmRegistrarAnotacao').click()
        #DA COMO APROVADO E ANOTA NO TXT DE PROCESSOS CONFORMES
        situacao = 'Aprovado'
        print("Próximo processo...")

    #CONDICAO DE PROCESSO NAO CONFORME
    else:
        print(f'Tome as medidas necessárias para o processo {processo}.')
        #MOSTRA OPCOES DE EXIGENCIA
        while True:
            try:
                exigencia = str(input('Se houver alguma exigência digite [1], senão, digite [2]: '))
                if exigencia == '1':
                    break
                elif exigencia == '2':
                    break
                else:
                    print("Opção inválida, tente novamente!")
            except ValueError:
                print("Opção inválida, tente novamente!")
        #CONDICAO DE EXIGENCIA
        if exigencia == '1':
            abre_email(navegador, processo)
            muda_janela('theomation')
            while True:
                try:
                    print("Selecione o tipo de exigência abaixo:\n",
                        "[1] Falta algum anexo no processo.\n",
                        "[2] Erro na declaração de conformidade ou alguma exigência que é necessário abrir novo processo.\n",
                        "[3] Produto já homologado.\n",
                        "[4] Outros.")
                    time.sleep(0.5)
                    exig = int(input("Opção: "))
                except ValueError:
                    print("Opção inválida. Digite um número inteiro entre 1 e 3.")
                    continue
                if exig in range(1, 5):
                    if exig == 1:
                        #DEFINE SITUACAO DO PROCESSO
                        situacao = 'Intercorrente'
                    elif exig == 3:
                        #DEFINE SITUACAO DO PROCESSO
                        situacao = 'Cancelado'
                    else:
                        situacao = 'Exigência'
                    email_enviado = envia_email(navegador, janela_principal, processo, nomeEstag, impProp, exig)
                    if email_enviado:
                        #VOLTA PARA A PAGINA INICIAL DO PROCESSO
                        vai_para_processo(navegador, processo)
                        time.sleep(1)
                        navegador.switch_to.frame('ifrConteudoVisualizacao')
                        #CLICA NO ICONE DE CONCLUIR PROCESSO
                        clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/processo_concluir.svg?18')]")
                        #clica_noelemento(navegador, By.XPATH,'//*[@id="divArvoreAcoes"]/a[19]')
                        navegador.switch_to.frame('ifrVisualizacao')
                        clica_noelemento(navegador, By.XPATH, '//*[@id="sbmSalvar"]')
                    else:
                        print('Cancelado envio de email')
                        return
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
    if situacao == 'Aprovado' or situacao == 'Exigência' or situacao == 'Intercorrente' or situacao == 'Cancelado':
        preenche_planilhageral(processo, nomeEstag, retido, situacao, codigo_rastreio, nome_interessado)
    else:
        print('--------------------------------------------------------------------------')
        print('--------------------------------------------------------------------------')
    # except Exception as error:
    #     print(f"Ocorreu um erro: {error}")
    pyautogui.PAUSE = 0.7

#FUNCAO PARA CONCLUIR PROCESSO
def concluiProcesso(navegador, lista_procConformes, lista_procCancelamento, nomeEstag, planilhaGeral):
    #VARIAVEL PARA VOLTAR À JANELA INICIAL
    janela_principal = navegador.current_window_handle

    # Carrega a planilha geral de processos com apenas a primeira coluna de processos (índice 0) 
    # e a terceira coluna que informa se está retido ou não (índice 2)
    try:
        df = pd.read_excel(planilhaGeral, usecols=[0,2,3])
    except Exception as e:
        print("Ocorreu um erro ao buscar o caminho da pasta geral")
        print(e)
        planilhaGeral = input("Digite o caminho da planilha geral manualmente: ").strip('"')
        planilhaGeral = planilhaGeral.replace("\\", "\\\\")
        df = pd.read_excel(planilhaGeral, usecols=[0,2,3])

    #ITERA SOBRE OS processo DA LISTA DE PROCESSOS CONFORMES
    print('Concluindo processos conformes\n')
    for processosAssinados in lista_procConformes[:]:
        print('\n')
        print(f'Concluindo processo nº {processosAssinados}')

        # Verificar se o processo está na planilha
        resultado = df[df['Nº do Processo SEI'] == processosAssinados]

        if not resultado.empty:
            # Se o processo for encontrado, pegar o valor da coluna 'Retido'
            try:
                retido = resultado.iloc[0]['Retido'].lower()
            except Exception as e:
                print(f"Processo {processosAssinados} não possui informação quanto à situação 'retido' na planilha.")
                continue
            codRastreio = resultado.iloc[0]['NumerodoRastreio']

            # Diz se o processo está retido ou não
            if retido == 'sim':
                print(f"O processo {processosAssinados} está retido com código de rastreio {codRastreio}")
            else:
                print(f"O processo {processosAssinados} não está retido")
        else:
            print(f"Processo {processosAssinados} não encontrado na planilha.")
            continue  # Pula para o próximo processo 

        try:
            # Interagir com o navegador para acessar a página do processo
            navegador.switch_to.default_content()
            navegador.find_element(By.ID, 'txtPesquisaRapida').clear()
            navegador.find_element(By.ID, 'txtPesquisaRapida').send_keys(processosAssinados)
            navegador.find_element(By.ID, 'txtPesquisaRapida').send_keys(Keys.ENTER)
            time.sleep(1.5)
            
            #PEGA DADOS DA DECLARACAO DE CONFORMIDADE
            navegador.switch_to.frame('ifrArvore')

            #CONFERE SE EXISTE A DECLARACAO DE CONFORMIDADE OU UMA PASTA
            # if check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade", navegador):
            #     navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade").click()
            #     time.sleep(0.3)
            # elif check_element_exists(By.XPATH, '//*[@id="spanPASTA1"]', navegador):
            #     navegador.find_element(By.XPATH, '//*[@id="spanPASTA1"]').click()
            #     time.sleep(1)
            #     navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade").click()
            #     time.sleep(0.3)
            #CONFERE SE EXISTE A DECLARACAO DE CONFORMIDADE OU UMA PASTA
            if check_element_exists(By.XPATH, '//*[@id="spanPASTA1"]', navegador):
                navegador.find_element(By.XPATH, '//*[@id="spanPASTA1"]').click()
                time.sleep(1)
                navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade").click()
                time.sleep(0.3)

            elif check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade", navegador):
                navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade").click()
                time.sleep(0.3)
            navegador.switch_to.default_content()
            navegador.switch_to.frame('ifrConteudoVisualizacao')
            navegador.switch_to.frame('ifrVisualizacao')

            #ARMAZENA EMAIL DO SOLICITANTE
            emailSol = navegador.find_element(By.XPATH, '/html/body/div[3]/a[1]').text

            navegador.switch_to.default_content()
            navegador.switch_to.frame('ifrArvore')

            #VERIFICA SE EXISTE O DESPACHO DECISORIO
            if check_element_exists(By.PARTIAL_LINK_TEXT, "Despacho Decisório", navegador):
                #CLICA NO DESPACHO DECISORIO
                navegador.find_element(By.PARTIAL_LINK_TEXT, "Despacho Decisório").click()
                # navegador.find_element(By.PARTIAL_LINK_TEXT, "Despacho Decisório").click()
                navegador.switch_to.default_content()
                navegador.switch_to.frame('ifrConteudoVisualizacao')
                navegador.switch_to.frame('ifrVisualizacao')
                #VERIFICA SE EXISTE ASSINATURA DO GERENTE
                if check_element_exists(By.XPATH, "/html/body/div[1]/table[1]", navegador):
                    navegador.switch_to.default_content()
                    navegador.switch_to.frame('ifrConteudoVisualizacao')
                    #CLICA NO ICONE DE EMAIL
                    clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/email_enviar.svg?18')]")
                    #clica_noelemento(navegador, By.XPATH, '//*[@id="divArvoreAcoes"]/a[8]')
                    # navegador.find_element(By.XPATH, "//img[@title='Enviar Documento por Correio Eletrônico']").click()
                    time.sleep(0.7)
                    #MUDA PARA A JANELA MAIS RECENTE
                    navegador.switch_to.window(navegador.window_handles[-1])
                    #time.sleep(2)
                    #SELECIONA O DROPDOWN COM AS OPCOES DE EMAIL
                    select_element = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="selDe"]')))
                    select = Select(select_element)
                    #SELECIONA A OPCAO DE EMAIL ANATEL
                    select.select_by_visible_text('ANATEL/E-mail de replicação <nao-responda@anatel.gov.br>')
                    #INSERE EMAIL DO SOLICITANTE E CLICA NO EMAIL DO SOLICITANTE
                    navegador.find_element(By.XPATH, '//*[@id="s2id_autogen1"]').send_keys(emailSol)
                    time.sleep(1)
                    #CLICA NO EMAIL DO SOLICITANTE
                    clica_noelemento(navegador, By.XPATH,'//*[@id="select2-result-label-2"]')
                    # navegador.find_element(By.XPATH, '//*[@id="select2-result-label-2"]').click()
                    #INSERE ASSUNTO DO EMAIL E TEXTO DO EMAIL
                    if retido == "sim":
                        endereco_email("notificacaosei.sp@anatel.gov.br", navegador)
                        endereco_email("notificacaosei.rj@anatel.gov.br", navegador)
                        endereco_email("notificacaosei.pr@anatel.gov.br", navegador)
                        navegador.find_element(By.ID, 'txtAssunto').send_keys(f'Processo SEI nº {processosAssinados} - Aprovado ({codRastreio})')
                        select_element = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="selTextoPadrao"]')))
                        select = Select(select_element)
                        #SELECIONA O TEXTO PADRAO DE DRONES RETIDOS APROVADOS
                        select.select_by_visible_text('Processo SEI ORCN - APROVAÇÃO - DRONE RETIDO')
                    else:
                        navegador.find_element(By.ID, 'txtAssunto').send_keys(f'Processo SEI nº {processosAssinados} - Aprovado')
                        select_element = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="selTextoPadrao"]')))
                        select = Select(select_element)
                        #SELECIONA O TEXTO PADRAO DE DRONES APROVADOS
                        select.select_by_visible_text('Processo SEI ORCN - APROVAÇÃO - DRONE')

                    #ENVIA EMAIL
                    clica_noelemento(navegador, By.XPATH,'//*[@id="divInfraBarraComandosInferior"]/button[1]')
                    # navegador.find_element(By.XPATH, '//*[@id="divInfraBarraComandosInferior"]/button[1]').click()
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
                    navegador.switch_to.frame('ifrConteudoVisualizacao')
                    #time.sleep(1)
                    clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/anotacao_cadastro.svg?18')]")
                    #navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[17]').click()
                    #LIMPA O TEXTO DA ANOTACAO
                    time.sleep(0.3)
                    navegador.switch_to.frame('ifrVisualizacao')
                    navegador.find_element(By.XPATH, '//*[@id="txaDescricao"]').clear()
                    #SALVA ANOTACAO
                    clica_noelemento(navegador, By.XPATH,'//*[@id="divInfraBarraComandosSuperior"]/button')
                    # navegador.find_element(By.XPATH, '//*[@id="divInfraBarraComandosSuperior"]/button').click()
                    #CLICA NO ICONE DE TAG
                    #time.sleep(0.5)
                    navegador.switch_to.default_content()
                    navegador.switch_to.frame('ifrConteudoVisualizacao')
                    clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/marcador_gerenciar.svg?18')]")
                    #navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[24]').click()
                    navegador.switch_to.frame('ifrVisualizacao')
                    try:
                        time.sleep(1)
                        navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click() 
                    except:
                        time.sleep(0.3)
                        clica_noelemento(navegador, By.XPATH,'//*[@id="tblMarcadores"]/tbody/tr[2]/td[6]/a[2]/img')
                        #navegador.find_element(By.XPATH, '//*[@id="tblMarcadores"]/tbody/tr[2]/td[6]/a[2]/img').click()
                        alert.accept()
                        time.sleep(0.3)
                        clica_noelemento(navegador, By.XPATH,'//*[@id="btnAdicionar"]')
                        #navegador.find_element(By.XPATH, '//*[@id="btnAdicionar"]').click()
                        #time.sleep(0.5)
                        clica_noelemento(navegador, By.XPATH,'//*[@id="selMarcador"]/div/a')
                        #navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/div/a').click()
                    #CLICA NO DROPDOWN DE TAG
                    #time.sleep(0.5)
                    #VERIFICA SE ESTA RETIDO
                    if retido == 'não':
                        #CLICA NA TAG DE NAO RETIDO
                        clica_noelemento(navegador, By.PARTIAL_LINK_TEXT, 'Processo SEI ORCN - Concluido')
                        # navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/ul/li[16]').click()
                    else:
                        #CLICA NA TAG DE RETIDO
                        clica_noelemento(navegador, By.PARTIAL_LINK_TEXT, 'Processo SEI ORCN - Concluido - Retidos')
                        # navegador.find_element(By.XPATH, '//*[@id="selMarcador"]/ul/li[17]').click()
                    time.sleep(0.2)
                    #COLOCA O TEXTO NA TAG
                    navegador.find_element(By.XPATH, '//*[@id="txaTexto"]').send_keys(nomeEstag)
                    #SALVA TAG
                    navegador.find_element(By.XPATH, '//*[@id="sbmSalvar"]').click()
                    
                    # navegador.find_element(By.XPATH, '//*[@id="sbmSalvar"]').click()
                    navegador.switch_to.default_content()
                    #ENTRA NA PAGINA INICIAL DO PROCESSO
                    navegador.find_element(By.ID, 'txtPesquisaRapida').send_keys(processosAssinados)
                    elementos = navegador.find_element(By.ID, 'txtPesquisaRapida')
                    elementos.send_keys(Keys.ENTER)
                    time.sleep(1)
                    #CONCLUI PROCESSO
                    navegador.switch_to.frame('ifrConteudoVisualizacao')
                    clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/processo_concluir.svg?18')]")
                    #clica_noelemento(navegador, By.XPATH,'//*[@id="divArvoreAcoes"]/a[19]')
                    navegador.switch_to.frame('ifrVisualizacao')
                    clica_noelemento(navegador, By.XPATH, '//*[@id="sbmSalvar"]')
                    # navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[20]').click()
                    time.sleep(0.3)
                    navegador.switch_to.default_content()
                else:
                    #NAO ENVIA EMAIL CASO NAO ESTEJA COM DESPACHO
                    print(f"O despacho do processo {processosAssinados} ainda não foi assinado, execute o script novamente após conter a assinatura do gerente.")
                    time.sleep(0.5)
        #CONDICAO DE ERRO
        except Exception as e:
            print("Ops! Algo deu errado.")
            print(e)
            print("Ocorreu um erro ao concluir um dos processos. Tire print da mensagem de erro e comunique ao desenvolvedor.")
            input("Pressione enter para continuar...")
            continue
            # continue  # Não interrompe, mas pula para o próximo processo
        
        finally:
            # Remover o processo da lista fora da iteração original
            lista_procConformes.remove(processosAssinados)
    
    if len(lista_procCancelamento) > 0:
        print("Concluindo pedidos de cancelamento")

    textoCancelamento = f"""Prezado(a) Senhor(a),

        Informamos que a solicitação de cancelamento de homologação foi deferida.


        FAVOR NÃO RESPONDER ESTE E-MAIL. 

        

        Atenciosamente,

        ORCN - Gerência de Certificação e Numeração

        SOR - Superintendência de Outorga e Recursos à Prestação

        Anatel - Agência Nacional de Telecomunicações"""

    for processosAssinados in lista_procCancelamento[:]:
        print('\n')
        print(f'Concluindo processo nº {processosAssinados}')
        vai_para_processo(navegador, processosAssinados)
        time.sleep(1)
        #PEGA DADOS DA DECLARACAO DE CONFORMIDADE
        navegador.switch_to.frame('ifrArvore')

        #CONFERE SE EXISTE A DECLARACAO DE CONFORMIDADE OU UMA PASTA
        if check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade", navegador):
            navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade").click()
            time.sleep(0.3)
        elif check_element_exists(By.XPATH, '//*[@id="spanPASTA1"]', navegador):
            navegador.find_element(By.XPATH, '//*[@id="spanPASTA1"]').click()
            time.sleep(1)
            navegador.find_element(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade").click()
            time.sleep(0.3)

        navegador.switch_to.default_content()
        navegador.switch_to.frame('ifrConteudoVisualizacao')
        navegador.switch_to.frame('ifrVisualizacao')

        try:
            #ARMAZENA EMAIL DO SOLICITANTE
            emailSol = navegador.find_element(By.XPATH, '/html/body/div[3]/a[1]').text
        except Exception as e:
            print("Não foi possivel obter o email do solicitante!")
            while True:
                emailSol = input("Digite o email do solicitante: ")
                opcao = input("Confirma o email digitado? Digite [1] para confirmar, [2] para digitar novamente ou [3] para cancelar: ")
                if opcao == '1':
                    break   
                elif opcao == '2':  
                    continue            
                elif opcao == '3':
                    return
                else:
                    print("Opção inválida! Tente novamente.")

        navegador.switch_to.default_content()
        navegador.switch_to.frame('ifrArvore')
        lista_despachos = navegador.find_elements(By.PARTIAL_LINK_TEXT, 'Despacho Decisório')
        for i in lista_despachos:
            navegador.switch_to.default_content()
            navegador.switch_to.frame('ifrArvore')
            i.click()
            navegador.switch_to.default_content()
            navegador.switch_to.frame('ifrConteudoVisualizacao')
            navegador.switch_to.frame('ifrVisualizacao')
            time.sleep(1)
            if check_element_exists(By.XPATH, "//*[contains(text(), 'Revogar o')]", navegador):
                print('Encontrado despacho de cancelamento')
                #VERIFICA SE EXISTE ASSINATURA DO GERENTE
                navegador.switch_to.default_content()
                navegador.switch_to.frame('ifrConteudoVisualizacao')
                navegador.switch_to.frame('ifrVisualizacao')
                if check_element_exists(By.XPATH, "/html/body/div[1]/table[1]", navegador):
                    navegador.switch_to.default_content()
                    navegador.switch_to.frame('ifrConteudoVisualizacao')
                    #CLICA NO ICONE DE ENVIO DE EMAIL
                    clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/email_enviar.svg?18')]")
                    #clica_noelemento(navegador, By.XPATH, '//*[@id="divArvoreAcoes"]/a[10]')
                    time.sleep(1)
                    #VAI PARA A JANELA MAIS RECENTE ABERTA
                    navegador.switch_to.window(navegador.window_handles[-1])
                    #time.sleep(1)
                    #ABRE O DROPDOWN COM AS OPCOES DE EMAIL DA ANATEL
                    select_element = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="selDe"]')))
                    select = Select(select_element)
                    #SELECIONA O EMAIL DA ANATEL
                    select.select_by_visible_text('ANATEL/E-mail de replicação <nao-responda@anatel.gov.br>')
                    #ESCREVE O EMAIL DO SOLICITANTE
                    navegador.find_element(By.XPATH, '//*[@id="s2id_autogen1"]').send_keys(emailSol)
                    time.sleep(1.2)
                    #CLICA NO EMAIL DO SOLICITANTE
                    clica_noelemento(navegador, By.CLASS_NAME, 'select2-result-label')
                    #INSERE ASSUNTO DO PROCESO
                    navegador.find_element(By.ID, 'txtAssunto').send_keys(f"Processo SEI nº {processosAssinados}")
                    #INSERE EMAIL COM CORPO DO EMAIL
                    navegador.find_element(By.ID, 'txaMensagem').send_keys(textoCancelamento)
                    #ENVIA EMAIL
                    clica_noelemento(navegador, By.XPATH,'//*[@id="divInfraBarraComandosInferior"]/button[1]')
                    # navegador.find_element(By.XPATH, '//*[@id="divInfraBarraComandosInferior"]/button[1]').click()
                    #FECHA ALERTA DO NAVEGADOR
                    alert = Alert(navegador)
                    alert.accept()

                    #RETORNA PARA A JANELA PRINCIPAL
                    navegador.switch_to.window(janela_principal)
                    vai_para_processo(navegador, processosAssinados)
                    time.sleep(1)
                    #CLICA NO ICONE DE ANOTACAO
                    navegador.switch_to.frame('ifrConteudoVisualizacao')
                    #time.sleep(1)
                    clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/anotacao_cadastro.svg?18')]")
                    #navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[17]').click()
                    #LIMPA O TEXTO DA ANOTACAO
                    time.sleep(0.3)
                    navegador.switch_to.frame('ifrVisualizacao')
                    navegador.find_element(By.XPATH, '//*[@id="txaDescricao"]').clear()
                    #SALVA ANOTACAO
                    clica_noelemento(navegador, By.XPATH,'//*[@id="divInfraBarraComandosSuperior"]/button')
                    #CONCLUI PROCESSO

                    vai_para_processo(navegador, processosAssinados)
                    navegador.switch_to.frame('ifrConteudoVisualizacao')
                    clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/processo_concluir.svg?18')]")
                    #clica_noelemento(navegador, By.XPATH,'//*[@id="divArvoreAcoes"]/a[19]')
                    navegador.switch_to.frame('ifrVisualizacao')
                    clica_noelemento(navegador, By.XPATH, '//*[@id="sbmSalvar"]')
                    # navegador.find_element(By.XPATH, '//*[@id="divArvoreAcoes"]/a[20]').click()
                    print(f"Processo nº {processosAssinados}")
                    time.sleep(0.3)
                    navegador.switch_to.default_content()
                else:
                    print(f"O despacho de cancelamento do processo {processosAssinados} ainda não foi assinado, execute o script novamente após conter a assinatura do gerente.")



def atribuir_processos(planilhaGeral, num_processos):

    # Lê a planilha do Excel
    try:
        df = pd.read_excel(planilhaGeral, usecols=[0,1])
    except Exception as e:
        print("Ocorreu um erro ao buscar o caminho da pasta geral")
        print(e)
        planilhaGeral = input("Digite o caminho da planilha geral manualmente: ").strip('"')
        planilhaGeral = planilhaGeral.replace("\\", "\\\\")
        df = pd.read_excel(planilhaGeral, usecols=[0,1])
        
    # Filtra as linhas onde a segunda coluna (Estagiário responsável) é uma data
    df['Estagiario responsável'] = pd.to_datetime(df['Estagiario responsável'], format='%d/%m/%Y', errors='coerce')
    df_com_datas = df.dropna(subset=['Estagiario responsável'])

    # Remove componente de tempo das datas usando .loc[]
    df_com_datas.loc[:, 'Estagiario responsável'] = df_com_datas['Estagiario responsável'].dt.date

    # Pega as n datas mais antigas
    processos_mais_antigos = df_com_datas.head(num_processos)

    # Cria a lista de processos a serem atribuídos
    lista_processos_atr = processos_mais_antigos['Nº do Processo SEI'].tolist()

    # Remove os \n e as aspas
    lista_processos_atr = [elemento.replace('\n', '') for elemento in lista_processos_atr]

    return lista_processos_atr, df_com_datas


def atribuicao(navegador, nomeEstag_sem_acento, nomeEstag, planilhaGeral):
    # Exemplo de uso
    navegador.switch_to.default_content()
    try:    
        clica_noelemento(navegador, By.ID, 'lnkControleProcessos').click()
    except Exception as e:
        pass

    try:
        navegador.find_element(By.ID, 'ancLiberarMeusProcessos').click()
    except Exception as e:
        print(e)
    num_processos = int(input('Digite a quantidade de processos que deseja atribuir: '))
    lista_processos_atr, df_com_datas = atribuir_processos(planilhaGeral, num_processos)
    print('Lista de processos a serem atribuídos:', lista_processos_atr)

    # Procura a caixa da Chirlene
    if not check_element_exists(By.PARTIAL_LINK_TEXT, 'chirlene.colab' , navegador):
        print("Não há processos da caixa da Chirlene na página inicial")
        print("Passando para a próxima página")
        navegador.find_element(By.XPATH, '//*[@id="lnkRecebidosProximaPaginaSuperior"]').click()
    clica_noelemento(navegador, By.PARTIAL_LINK_TEXT, 'chirlene.colab')

    processos_para_atr = []
    processos_nao_encontrados = []
    processos_restantes = list(df_com_datas['Nº do Processo SEI'])
    processos_restantes = [elemento.replace('\n', '') for elemento in processos_restantes]

    for processos_atr in lista_processos_atr:
        try:
            clica_noelemento(navegador, By.XPATH, f'//label[@title="{processos_atr}"]', 1)
            processos_para_atr.append(processos_atr)
            processos_restantes.remove(processos_atr)
            time.sleep(0.3)
        except:
            print(f'Processo {processos_atr} não encontrado.')
            processos_restantes.remove(processos_atr)
            processos_nao_encontrados.append(processos_atr)

    # Caso não encontre todos os processos, pegar os próximos na lista
    if processos_nao_encontrados:
        qtd_nao_encontrados = len(processos_nao_encontrados)

        # Seleciona os próximos processos a partir da lista de processos restantes
        novos_processos = processos_restantes[:qtd_nao_encontrados]
        print(f"Processos a serem atribuídos no lugar dos não encontrados: {novos_processos}.")

        # Tenta atribuir esses novos processos
        for processos_atr in novos_processos:
            try:
                navegador.find_element(By.XPATH, f'//label[@title="{processos_atr}"]').click()
                processos_para_atr.append(processos_atr)
                time.sleep(0.3)
            except:
                print(f'Processo {processos_atr} não encontrado nos adicionais.')

    select_element = navegador.find_element(By.XPATH, '//*[@id="selAtribuicao"]')
    select = Select(select_element)
    time.sleep(0.5)
    for opcao in select.options:
        if nomeEstag_sem_acento in opcao.text:
            select.select_by_visible_text(opcao.text)
            break
    navegador.switch_to.default_content()
    navegador.find_element(By.ID, 'btnSalvar').click()
    navegador.find_element(By.XPATH, '//*[@id="divFiltro"]/div[2]/a').click()
    pyautogui.PAUSE = 0.7
    # edge = pyautogui.locateOnScreen('imagensAut/edge.png', confidence=0.7)
    # CLICA NO NAVEGADOR
    muda_janela('Distribuição Processo Drone.xlsx')
    
    time.sleep(0.5)
    for processos_atr2 in processos_para_atr:
        # pyautogui.PAUSE = 0.7
        # # COPIA NUMERO DO PROCESSO
        # pyperclip.copy(processos_atr2)
        # # PESQUISA PROCESSO NA PLANILHA
        # pyautogui.hotkey('ctrl', 'l')
        # time.sleep(0.5)
        # # COLA NUMERO DO PROCESSO E APERTA ENTER PARA PESQUISAR O PROCESSO
        # pyautogui.hotkey('ctrl', 'v')
        # time.sleep(0.3)
        # pyautogui.press('enter')
        # pyautogui.PAUSE = 0.4
        # pyautogui.press('esc')
        celula_encontrada = localiza_processo(processos_atr2)
        #COPIA O NUMERO DO PROCESSO NA CELULA PARA SABER SE ELE FOI ENCONTRADO CORRETAMENTE
        # pyperclip.copy('nan')
        # pyautogui.hotkey('ctrl','c')
        # time.sleep(0.3)
        # celula_encontrada = pyperclip.paste()
        # celula_encontrada = celula_encontrada.replace('\n', '').strip('"')
        print(f'Célula encontrada: {celula_encontrada}')
        while celula_encontrada != processos_atr2:
            if celula_encontrada == 'nan':
                muda_janela('theomation')
                input('Parece que o Excel não estava em foco na janela do Microsoft Edge. Volte para a página do app, coloque em foco (clicando em qualquer lugar da planilha) e pressione enter.')
                muda_janela('Distribuição Processo Drone.xlsx')
            celula_encontrada = localiza_processo(processos_atr2)
        pyautogui.press('right')
        pyperclip.copy(nomeEstag)
        pyautogui.hotkey('ctrl', 'v')