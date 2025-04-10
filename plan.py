import pyautogui
import pyperclip
import time
import funcoes as fc
import tkinter as tk
# import undetected_chromedriver as uc
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import subprocess
from pywinauto import Application
from pywinauto.findwindows import find_windows, ElementAmbiguousError, ElementNotFoundError
from selenium.webdriver import Edge

def preencher_campos():
    user = user_var.get()
    senha = senha_var.get()

    navegador.find_element(By.XPATH,'//*[@id="txtUsuario"]').send_keys(f"{user}")
    navegador.find_element(By.XPATH,'//*[@id="pwdSenha"]').send_keys(f"{senha}")
    root.destroy()


def check_element_exists(by, value):
    try:
        navegador.find_element(by, value)
        return True
    except NoSuchElementException:
        return False


def preenche_plan(nomeSol, nomeInt, data, retido, codigo_rastreio, n_serie, n_serie2):
    fc.muda_janela('Distribuição Processo Drone_2025.xlsx')
    time.sleep(0.4)
    pyautogui.PAUSE = 0.2
    pyautogui.press('right')
    pyautogui.press('right')
    pyperclip.copy(retido)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('right')
    pyperclip.copy(codigo_rastreio)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('right')
    pyperclip.copy(nomeSol)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('right')
    pyperclip.copy(nomeInt)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('right')
    pyautogui.press('right')
    data_somente=data.split(' ')[0]
    pyperclip.copy(data_somente)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('right')
    if not n_serie.strip():
        pyautogui.press('right')
    else:
        pyperclip.copy(n_serie)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('right')
    if not n_serie2.strip():
        pyautogui.press('right')
    else:
        pyperclip.copy(n_serie2)
        pyautogui.hotkey('ctrl', 'v')
    pyautogui.PAUSE = 0.1
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.PAUSE = 0.7
    pyautogui.press('down')


def preenche_plan2(nomeSol, nomeInt, data):
    fc.muda_janela('Distribuição Processo Drone_2025.xlsx')
    time.sleep(0.4)
    pyautogui.PAUSE = 0.2
    pyautogui.press('right')
    pyautogui.press('right')
    pyautogui.press('right')
    pyautogui.press('right')
    pyperclip.copy(nomeSol)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('right')
    pyperclip.copy(nomeInt)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('right')
    pyautogui.press('right')
    data_somente=data.split(' ')[0]
    pyperclip.copy(data_somente)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.PAUSE = 0.1
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.PAUSE = 0.7
    pyautogui.press('down')


def endereco_email(endereco, navegador):
    navegador.find_element(By.XPATH, '//*[@id="s2id_autogen1"]').send_keys(endereco)
    time.sleep(1)
    #CLICA NO EMAIL DO SOLICITANTE
    navegador.find_element(By.XPATH, '//*[@id="select2-drop"]/ul').click()
    time.sleep(0.2)


def manda_email(n_processo, codigo_rastreio):
    navegador.switch_to.default_content()
    elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
    elementos.send_keys(n_processo)
    elementos.send_keys(Keys.ENTER)
    time.sleep(1.5)
    navegador.switch_to.frame('ifrConteudoVisualizacao')
    clica_noelemento(By.XPATH, "//img[contains(@src, 'svg/email_enviar.svg?18')]")
    navegador.switch_to.window(navegador.window_handles[-1])
    select_element = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="selDe"]')))
    select = Select(select_element)
    select.select_by_visible_text('ANATEL/E-mail de replicação <nao-responda@anatel.gov.br>')
    endereco_email("notificacaosei.sp@anatel.gov.br", navegador)
    endereco_email("notificacaosei.rj@anatel.gov.br", navegador)
    endereco_email("notificacaosei.pr@anatel.gov.br", navegador)
    navegador.find_element(By.ID, 'txtAssunto').send_keys(f'Processo SEI nº {n_processo} - Código de rastreio: {codigo_rastreio.strip()} - Aberto')
    navegador.find_element(By.ID, 'txaMensagem').send_keys(f'Processo SEI nº {n_processo} - Código de rastreio: {codigo_rastreio.strip()} - Aberto')
    #ENVIA EMAIL
    navegador.find_element(By.XPATH, '//*[@id="divInfraBarraComandosInferior"]/button[1]').click()
    #FECHA ALERTA DO NAVEGADOR
    time.sleep(0.5)
    alert = Alert(navegador)
    alert.accept()
    time.sleep(2)
    # chrome=pyautogui.locateOnScreen('imagensAut/chrome.png', confidence=0.7)
    # pyautogui.click(chrome)
    navegador.switch_to.window(janela_principal)


def clica_noelemento(modo_procura, element_id):
    # Espera até que o elemento seja carregado (exemplo: elemento localizado por ID)
    try:
        element = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((modo_procura, element_id))
        )
        # Clica no elemento
        element.click()
    except TimeoutException:
        print(f"Elemento {element_id} não foi carregado no tempo esperado")

def reabre_processo():
    navegador.switch_to.default_content()
    navegador.switch_to.frame('ifrConteudoVisualizacao')
    try:
        fc.clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/processo_reabrir.svg?18')]", 2)
        print('Processo foi aberto novamente')
        navegador.switch_to.default_content()
        return True
    except:
        navegador.switch_to.default_content()
        return False
    
def fecha_processo():
    navegador.switch_to.default_content()
    navegador.switch_to.frame('ifrConteudoVisualizacao')
    try:
        fc.clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/processo_concluir.svg?18')]", 2)
        navegador.switch_to.frame('ifrVisualizacao')
        fc.clica_noelemento(navegador, By.XPATH, '//*[@id="sbmSalvar"]')
    except Exception as e:
        print(e)

def tira_restrito():
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
        clica_noelemento(By.XPATH, '//*[@id="divArvoreAcoes"]/a[2]/img')
        time.sleep(0.7)
        navegador.switch_to.frame('ifrVisualizacao')
        #SELECIONA A OPCAO DE PUBLICO NO DOCUMENTO
        clica_noelemento(By.XPATH,'//*[@id="divOptPublico"]/div/label')
        #SALVA AS MUDANCAS
        clica_noelemento(By.ID,'btnSalvar')
        navegador.switch_to.default_content()
        navegador.switch_to.frame('ifrArvore')
    navegador.switch_to.default_content()

#chrome_driver_path = "chromedriver-win64\chromedriver.exe" 
#servico = Service(chrome_driver_path)
# servico = Service(ChromeDriverManager().install())
# pyautogui.PAUSE = 0.7
# #INICIA O NAVEGADOR
# navegador = webdriver.Chrome(service=servico)
navegador = Edge()
navegador.maximize_window()
navegador.get('https://sei.anatel.gov.br/')
janela_principal = navegador.current_window_handle
# edge=pyautogui.locateOnScreen('imagensAut/edge.png', confidence=0.7)
# pyautogui.click(edge)
try:
    # Busca todas as janelas que contenham o título especificado
    matches = find_windows(title_re=f".*Distribuição Processo Drone_2025.xlsx.*", backend="win32", visible_only=True)
    
    if not matches:
        raise ElementNotFoundError(f"Nenhuma janela encontrada com o título 'Distribuição Processo Drone_2025.xlsx'.")

    if len(matches) > 1:
        print(f"Mais de uma janela encontrada com o título 'Distribuição Processo Drone_2025.xlsx'. Selecionando a mais recente.")
    
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
    time.sleep(3)
    sitePlan = 'https://anatel365.sharepoint.com/:x:/r/sites/lista.orcn/_layouts/15/doc.aspx?sourcedoc=%7B458e5539-83f9-4bd7-91a8-21f8cd0ea009%7D&action=edit'
    pyperclip.copy(sitePlan)
    pyautogui.hotkey('ctrl', 'l')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')

# time.sleep(3)
# sitePlan = 'https://anatel365.sharepoint.com/:x:/r/sites/lista.orcn/_layouts/15/Doc.aspx?sourcedoc=%7B4130A4D6-7F00-45D4-A328-ED0866A62335%7D&file=Distribui%C3%A7%C3%A3o%20Processo%20Drone.xlsx&action=default&mobileredirect=true'
# pyperclip.copy(sitePlan)
# pyautogui.hotkey('ctrl', 'l')
# pyautogui.hotkey('ctrl', 'v')
# pyautogui.press('enter')
# chrome=pyautogui.locateOnScreen('imagensAut/chrome.png', confidence=0.7)
# pyautogui.click(chrome)
fc.muda_janela('SEI / ANATEL')

while True:
    #INICIA JANELA
    try:
        root = tk.Tk()
        root.title("Login de usuário")
        root.attributes('-topmost', True)
        largura_tela = root.winfo_screenwidth()
        root.geometry(f'+{largura_tela}+0')
        root.state('zoomed')

        #RECEBE USUARIO E SENHA
        user_var = tk.StringVar()
        senha_var = tk.StringVar()

        frame = tk.Frame(root)
        frame.place(relx=0.5, rely=0.5, anchor='center')

        # CRIA LABEL CENTRALIZADO
        tk.Label(frame, text="LOGIN NO SEI ANATEL", font=('Arial',30)).grid(row=0, columnspan=2)
        tk.Label(frame, text="").grid(row=1, column=0)
        tk.Label(frame, text="").grid(row=2, column=0)

        # Cria input usuário
        tk.Label(frame, text="Usuário:", font=('Arial', 15)).grid(row=3, column=0, sticky='e')
        user_entry = tk.Entry(frame, textvariable=user_var, width=30)
        user_entry.grid(row=3, column=1, sticky='w')

        # Cria input senha censurada
        tk.Label(frame, text="Senha:", font=('Arial', 15)).grid(row=4, column=0, sticky='e')
        senha_entry = tk.Entry(frame, textvariable=senha_var, show="*", width=30)
        senha_entry.grid(row=4, column=1, sticky='w')

        # Botão que envia usuário e senha
        tk.Label(frame, text="").grid(row=5, column=0)
        login_button = tk.Button(frame, text="Entrar", font=('Arial', 10), command=preencher_campos, width=20)
        login_button.grid(row=6, columnspan=2)
        
        # Associa a tecla Enter ao botão de login
        root.bind('<Return>', lambda event: preencher_campos())

        #INFORMA QUE OS DADOS NAO SERAO ARMAZENADOS NO SOFTWARE
        tk.Label(frame, text="").grid(row=7, column=0)
        tk.Label(frame, text="").grid(row=8, column=0)
        tk.Label(frame, text="O usuário e senha não serão armazenados em nenhum local durante a execução do script. Ambos serão utilizados apenas para o acesso ao site.", font=('Arial',15)).grid(row=9, columnspan=2)

        label = tk.Label(root, text="Script made by Theo Ravi.", font=('Arial', 10))
        label.place(relx=1.0, rely=1.0, anchor='se')

        root.mainloop()
        #ACESSA O SEI
        time.sleep(1)

        navegador.find_element(By.XPATH,'//*[@id="sbmAcessar"]').click()
    except Exception:
        print('Ocorreu um erro, tente novamente.')
# ENTRAR NO PROCESSO
    if check_element_exists(By.XPATH,'//*[@id="divFiltro"]/div[2]/a'):
        navegador.find_element(By.XPATH,'//*[@id="divFiltro"]/div[2]/a').click()
        break
    else:
        print('Usuário ou senha incorretos. Digite novamente')

verifica=input('Aperte enter após filtrar a planilha geral.')

# edge=pyautogui.locateOnScreen('imagensAut/edge.png', confidence=0.7)
# pyautogui.click(edge)
fc.muda_janela('Distribuição Processo Drone_2025.xlsx')
pyautogui.click(x=160, y=375)
counter = 0
while True:
    try:
        retido=""
        pyautogui.PAUSE = 0.7
        # Pressiona a tecla Ctrl
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.2)
        n_processo = pyperclip.paste()
        print("Processo: " + n_processo)
        if n_processo == '':
            counter += 1
        if counter == 5:
            print('Parece que todos os processos foram preenchidos. Finalizando programa...')
            navegador.close()
            break
        # chrome=pyautogui.locateOnScreen('imagensAut/chrome.png', confidence=0.7)
        # pyautogui.click(chrome)
        fc.muda_janela('SEI -')
        navegador.switch_to.default_content()
        navegador.find_element(By.ID,'txtPesquisaRapida').click()
        pyautogui.hotkey('ctrl', 'v')
        elementos = navegador.find_element(By.ID,'txtPesquisaRapida')
        elementos.send_keys(Keys.ENTER)    
        time.sleep(0.5)
        processo_reaberto = reabre_processo()
        navegador.switch_to.frame('ifrArvore')
        if check_element_exists(By.PARTIAL_LINK_TEXT, 'Recibo Eletrônico'):
            if check_element_exists(By.XPATH, '//*[@id="spanPASTA1"]'):
                navegador.find_element(By.XPATH, '//*[@id="spanPASTA1"]').click()
            time.sleep(1)
            clica_noelemento(By.PARTIAL_LINK_TEXT, 'Recibo Eletrônico')
            navegador.switch_to.default_content()
            navegador.switch_to.frame('ifrConteudoVisualizacao')
            navegador.switch_to.frame('ifrVisualizacao')
            time.sleep(1)
            nomeSol = navegador.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[1]/td[2]').text
            nomeInt = navegador.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[6]/td').text.strip()
            data = navegador.find_element(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr[2]/td[2]').text
            navegador.switch_to.default_content()
            #COLETAR DADOS DA DECLARAÇÃO DE CONFORMIDADE
            navegador.switch_to.frame('ifrArvore')
            time.sleep(1)
            if check_element_exists(By.PARTIAL_LINK_TEXT, 'Declaração de Conformidade - Drone'):
                clica_noelemento(By.PARTIAL_LINK_TEXT, 'Declaração de Conformidade - Drone')
                navegador.switch_to.default_content()
                navegador.switch_to.frame('ifrConteudoVisualizacao')
                navegador.switch_to.frame('ifrVisualizacao')
                time.sleep(0.7)
                codigo_rastreio = navegador.find_element(By.XPATH,'/html/body/table[4]/tbody/tr/td[2]').text
                codigo_rastreio = codigo_rastreio.replace('-','').replace('.','')
                try:
                    n_serie = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[2]/td[3]').text
                    n_serie2 = navegador.find_element(By.XPATH,'/html/body/table[3]/tbody/tr[3]/td[3]').text
                except:
                    n_serie = ''
                    n_serie2 = ''
                if not codigo_rastreio.strip():
                    retido='Não'
                else:
                    manda_email(n_processo, codigo_rastreio)
                    retido='Sim'
                    time.sleep(0.2)
                tira_restrito()
                preenche_plan(nomeSol, nomeInt, data, retido, codigo_rastreio, n_serie, n_serie2)
            elif check_element_exists(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Importado Uso Próprio"):
                print('Encontrada declaracao de conformidade drone')
                clica_noelemento(By.PARTIAL_LINK_TEXT, "Declaração de Conformidade - Importado Uso Próprio")
                navegador.switch_to.default_content()
                navegador.switch_to.frame('ifrConteudoVisualizacao')
                navegador.switch_to.frame('ifrVisualizacao')
                time.sleep(0.7)
                try:
                    codigo_rastreio = navegador.find_element(By.XPATH,'/html/body/table[5]/tbody/tr/td[2]').text
                except Exception as e:
                    print(e)   
                codigo_rastreio = codigo_rastreio.replace('-','').replace('.','').replace(' ','').strip()

                n_serie = navegador.find_element(By.XPATH,'/html/body/table[4]/tbody/tr[2]/td[3]').text
                try:
                    n_serie2 = navegador.find_element(By.XPATH,'/html/body/table[4]/tbody/tr[3]/td[3]').text
                except:
                    n_serie2 = ''
                if not codigo_rastreio.strip():
                    retido='Não'
                else:
                    manda_email(n_processo, codigo_rastreio)
                    retido='Sim'
                    time.sleep(0.2)
                if not processo_reaberto:
                    tira_restrito()
                preenche_plan(nomeSol, nomeInt, data, retido, codigo_rastreio, n_serie, n_serie2)
            else:
                print("Declaração de conformidade não encontrada.")
                preenche_plan2(nomeSol, nomeInt, data)
        else:
            print("Processo não contém recibo.")
            print("Pulando processo...")
            # edge=pyautogui.locateOnScreen('imagensAut/edge.png', confidence=0.7)
            # edge=pyautogui.locateOnScreen('imagensAut/edge.png', confidence=0.7)
            # pyautogui.click(edge)
            fc.muda_janela('Distribuição Processo Drone_2025.xlsx')
            pyautogui.press('down')
    
        if processo_reaberto:
            fecha_processo()
    except Exception as e:
        print(e)
        # pyautogui.click(edge)
        fc.muda_janela('Distribuição Processo Drone_2025.xlsx')
        pyautogui.press('down')        
