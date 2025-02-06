import pyautogui
import time
import funcoes as fc
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

# "53500.000025/2025-03": "NM 691 730 476 BR",
# "53500.000035/2025-31": "NM633261141BR",
# "53500.000038/2025-74": "NM675742484BR",
# "53500.000017/2025-59": "NM667787387BR",
# "53500.000090/2025-21": "NM665901478BR",

# processos_testar = {
#                     "53500.000282/2025-37": "NM622222235BR",
#                     "53500.000280/2025-48": "NM645610825BR",
#                     "53500.000269/2025-88": "NM665307128BR",
#                     "53500.000262/2025-66": "NM647822059BR",
#                     "53500.000373/2025-72": "NM682706686BR",
#                     "53500.000366/2025-71": "LB585022745HK",
#                     "53500.000376/2025-14": "NM647446993BR",
#                     "53500.000396/2025-87": "NM661969157BR",
#                     "53500.000412/2025-31": "NM669327349BR",
#                     "53500.000657/2025-69": "NM661953979BR",
#                     "53500.000656/2025-14": "CNBR00028920928 Transportadora ANJUN",
#                     "53500.000644/2025-90": "NM678828975BR",
#                     "53500.000635/2025-07": "NM633694417BR",
#                     "53500.000615/2025-28": "NM653938569BR",
#                     "53500.000561/2025-09": "NM671578306BR",
#                     "53500.000518/2025-35": "Nm682984906br",
#                     "53500.000498/2025-01": "NM672514810BR",
#                     "53500.000489/2025-10": "NM678029048BR",
#                     "53500.000431/2025-68": "NM677022542BR",
#                     "53500.000427/2025-08": "NM678343185BR",
#                     "53500.000603/2025-01": "NM689210035BR",
#                     "53500.000580/2025-27": "NM657957325BR",
#                     "53500.000456/2025-61": "NM 669 661 319 BR",
#                     "53500.000425/2025-19": "NM620448545BR",
#                     "53500.000736/2025-70": "NM677165170BR",
#                     "53500.000774/2025-22": "NM658668832BR",
#                     "53500.000859/2025-19": "NM670571686BR",
#                     "53500.000882/2025-03": "NM646621529BR",
#                     "53500.000928/2025-86": "NM628937045BR",
#                     "53500.000755/2025-04": "NM667519156BR",
#                     "53500.000758/2025-30": "NM646676377BR",
#                     "53500.000860/2025-35": "NM681318787BR",
#                     "53500.000900/2025-49": "NM657541866BR",
#                     "53500.000944/2025-79": "NM688225807BR",
#                     "53500.000954/2025-12": "NM680137528BR",
#                     "53500.000969/2025-72": "LB585233483HK",
#                     "53500.000990/2025-78": "NM669483131BR",
#                     "53500.000729/2025-78": "NM687611319BR",
#                     "53500.000728/2025-23": "NM677165170BR",
#                     "53500.001011/2025-07": "NM681629529BR",
#                     "53500.001274/2025-16": "NM669326895BR",
#                     "53500.001226/2025-10": "ND 312 784 241 BR",
#                     "53500.001086/2025-80": "NM 688 903 415 BR",
#                     "53500.001248/2025-80": "NM656600109BR",
#                     "53500.001125/2025-49": "LB585165580HK",
#                     "53500.001115/2025-11": "LM112457192CN",
#                     "53500.001295/2025-23": "ND 404 134 803 BR",
#                     "53500.001560/2025-73": "NM694734534BR",
#                     "53500.001385/2025-14": "MN622949646BR",
#                     "53500.001414/2025-48": "NM610663116BR",
#                     "53500.001510/2025-96": "NM632187997BR",
#                     "53500.001558/2025-02": "NM652228191BR",
#                     "53500.001587/2025-66": "NM688816945BR",
#                     "53500.001654/2025-42": "NM682432015BR",
#                     "53500.001848/2025-48": "NM681265320BR",
#                     "53500.001712/2025-38": "NM616179918BR",
#                     "53500.001701/2025-58": "NM657943683BR",
#                     "53500.001689/2025-81": "NM679559055BR",
#                     "53500.001795/2025-65": "LB585213047HK",
#                     "53500.001888/2025-90": "NM667017083BR",
#                     "53500.001920/2025-37": "LB585273712HK",
#                     "53500.001928/2025-01": "NM645266319BR",
#                     "53500.001939/2025-83": "NM667347181BR",
#                     "53500.001979/2025-25": "NM695217956BR",
#                     "53500.001964/2025-67": "ND354951550BR",
#                     "53500.001999/2025-04": "NM693521294BR",
#                     "53500.002045/2025-19": "NM677832457BR",
#                     "53500.002188/2025-12": "NM693530092BR",
#                     "53500.002000/2025-36": "LB585320545HK",
#                     "53500.002035/2025-75": "ND392878475BR",
#                     "53500.002139/2025-80": "NM695223320BR",
#                     "53500.002183/2025-90":  "ND420151147BR",
#                     "53500.002261/2025-56": "AA605333196BR",
#                     "53500.002273/2025-81": "NM674202545BR ",
#                     "53500.002276/2025-14": "NM683240997BR",
#                     "53500.002429/2025-23": "LB585351030HK",
#                     "53500.002621/2025-10": "NM 696 222 056 BR",
#                     "53500.002572/2025-15": "NM688042066BR",
#                     "53500.002447/2025-13": "ND 401 005 566 BR",
#                     "53500.002448/2025-50": "NM677652092BR",
#                     "53500.002451/2025-73": "NM683086003BR",
#                     "53500.002462/2025-53": "NM670237107BR",
#                     "53500.002471/2025-44": "NM677757858BR",
#                     "53500.002487/2025-57": "NM685165171BR",
#                     "53500.002488/2025-00": "NM669110435BR",
#                     "53500.002490/2025-71": "ND432804760BR",
#                     "53500.002493/2025-12": "NM684341626BR",
#                     "53500.002497/2025-92": "NM670483085BR",
#                     "53500.002503/2025-10": "NM686106991BR",
#                     "53500.002576/2025-01": "NM672066422BR",
#                     "53500.002666/2025-94": "NM666546416BR",
#                     "53500.002679/2025-63": "NM 671 182 489 BR",
#                     "53500.002685/2025-11": "NM676496192BR",
#                     "53500.002686/2025-65": "NM 669 318 228 BR",
#                     "53500.002697/2025-45": "NM 695 217 409 BR",
#                     "53500.002900/2025-83": "NM 684 326 941 BR",
#                     "53500.002990/2025-11": "NM666438249BR",
#                     "53500.002994/2025-91": "NM692008715BR",
#                     "53500.002999/2025-13": "NM693892953BR",
#                     "53500.003052/2025-20": "NM667449441BR",
#                     "53500.003101/2025-24": "NM660466357BR",
#                     "53500.002977/2025-53": "NM694962307BR",
#                     "53500.003060/2025-76": "NM694143979BR",
#                     "53500.003089/2025-58": "NM665872215BR",
#                     "53500.003099/2025-93": "NM695223846BR",
#                     "53500.003104/2025-68": "LB585309553HK",
#                     "53500.003108/2025-46": "NM652884802BR",
#                     "53500.003120/2025-51": "NM692816111BR",
#                     "53500.003153/2025-09": "NM683471655BR",
#                     "53500.003189/2025-84": "NM669363553BR",
#                     "53500.003355/2025-42": "NM652458994BR",
#                     "53500.003259/2025-02": "NM693761459BR",
#                     "53500.003307/2025-54": "NM672755740BR",
#                     "53500.003334/2025-27": "NM667519672BR",
#                     "53500.003354/2025-06": "ND373827952BR",
#                     "53500.003357/2025-31": "NM695788043BR",
#                     "53500.003359/2025-21": "NM669325921BR",
#                     "53500.003376/2025-68": "ND386461036BR",
#                     "53500.003378/2025-57": "ND386461036BR",
#                     "53500.003395/2025-94": "NM696246928BR",
#                     "53500.003482/2025-41": "NM698138696BR",
#                     "53500.003602/2025-19": "NM668006657BR",
#                     "53500.003607/2025-33": "NM689037734BR",
#                     "53500.003607/2025-33": "NM689037734BR",
#                     "53500.003612/2025-46": "NM672173276BR",
#                     "53500.003659/2025-18": "NM693521294BR",
#                     "53500.003671/2025-14": "NM654697609BR",
#                     "53500.003678/2025-36": "NM694861515BR",
#                     "53500.003681/2025-50": "NM695984618BR",
#                     "53500.003731/2025-07": "NM688780424BR",
#                     "53500.003749/2025-09": "NM698305310BR",
#                     "53500.003829/2025-56": "NM652387082BR",
#                     "53500.003869/2025-06": "NM696562823BR",
#                     "53500.003937/2025-29": "NM693457335BR",
#                     "53500.003468/2025-48": "NM667371835BR",
#                     "53500.003493/2025-21":" NM67109205BR",
#                     "53500.003506/2025-62": "ND430773605BR",
#                     "53500.003523/2025-08": "NM 700 011 890 BR",
#                     "53500.003615/2025-80": "ND445604682BR",
#                     "53500.003652/2025-98": "NM681545801BR",
#                     "53500.003658/2025-65": "NM739004325BR",
#                     "53500.003733/2025-98": "NM697332203BR",
#                     "53500.003735/2025-87": "NM672671015BR",
#                     "53500.003747/2025-10": "LM112815558CN",
#                     "53500.003754/2025-11": "NM692112280BR",
#                     "53500.003798/2025-33": "LM113007209CN",
#                     "53500.003863/2025-21": "NM697595465BR",
#                     "53500.003866/2025-64": "NM693750460BR",
#                     "53500.003881/2025-11": "ND451954974BR",
#                     "53500.003891/2025-48": "NM697218645BR",
#                     "53500.004077/2025-41": "CNBR00039909888",
#                     "53500.004227/2025-16": "NM655587896BR",
#                     "53500.004600/2025-39": "NM694852315BR",
#                     "53500.004591/2025-86": "ND434587379BR",
#                     "53500.004583/2025-30": "NM666962249BR",
#                     "53500.004559/2025-09": "NM688316114BR",
#                     "53500.004230/2025-30": "NM738925075BR",
#                     "53500.004232/2025-29": "NM738929236BR",
#                     "53500.004236/2025-15": "NM699053325BR",
#                     "53500.004238/2025-04": "NM738922710BR",
#                     "53500.004241/2025-10": "NM699043694BR",
#                     "53500.004259/2025-11": "ND481442846BR",
#                     "53500.004324/2025-17": "NM699041146BR",
#                     "53500.004342/2025-91": "NM689547114BR",
#                     "53500.004402/2025-75": "NM661841703BR",
#                     "53500.004431/2025-37": "ND 431 701 097 BR",
#                     "53500.004451/2025-16": "NM 671 418 785 BR",
#                     "53500.004556/2025-67": "NM697873212BR",
#                     "53500.004558/2025-56": "NM6987404679BR",
#                     "53500.004570/2025-61": "NM664176806BR",
#                     "53500.004573/2025-02": "LM112807976CN",
#                     "53500.004589/2025-15": "ND433581658BR",
#                     "53500.004590/2025-31": "NM690568672BR",
#                     "53500.004679/2025-06": "NM671407791BR",
#                     "53500.004717/2025-12": "LM112445282CN",
#                     "53500.004729/2025-47": "NM742213535BR",
#                     "53500.004734/2025-50": "NM740063498BR",
#                     "53500.004756/2025-10": "NM695991208BR",
#                     "53500.004769/2025-99": "NM739118384BR",
#                     "53500.004774/2025-00": "ND451465993BR",
#                     "53500.004791/2025-39": "ND451465993BR",
#                     "53500.004965/2025-63": "LM113006526CN",
#                     "53500.004963/2025-74": "NM738943900BR"}

# def reabre_processo():
#     navegador.switch_to.default_content()
#     navegador.switch_to.frame('ifrConteudoVisualizacao')
#     try:
#         fc.clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/processo_reabrir.svg?18')]", 2)
#         print('Processo foi aberto novamente')
#         navegador.switch_to.default_content()
#         return True
#     except:
#         navegador.switch_to.default_content()
#         return False
# def fecha_processo():
#     navegador.switch_to.default_content()
#     navegador.switch_to.frame('ifrConteudoVisualizacao')
#     try:
#         navegador.find_element(By.XPATH, "//img[contains(@src, 'svg/processo_concluir.svg?18')]").click()
#         navegador.switch_to.frame('ifrVisualizacao')
#         fc.clica_noelemento(navegador, By.XPATH, '//*[@id="sbmSalvar"]')
#     except:
#         pass
# def manda_email(n_processo, codigo_rastreio, janela_principal):
#     navegador.switch_to.default_content()
#     navegador.switch_to.frame('ifrConteudoVisualizacao')
#     fc.clica_noelemento(navegador, By.XPATH, "//img[contains(@src, 'svg/email_enviar.svg?18')]")
#     navegador.switch_to.window(navegador.window_handles[-1])
#     select_element = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="selDe"]')))
#     select = Select(select_element)
#     select.select_by_visible_text('ANATEL/E-mail de replicação <nao-responda@anatel.gov.br>')
#     fc.endereco_email("notificacaosei.sp@anatel.gov.br", navegador)
#     fc.endereco_email("notificacaosei.rj@anatel.gov.br", navegador)
#     fc.endereco_email("notificacaosei.pr@anatel.gov.br", navegador)
#     navegador.find_element(By.ID, 'txtAssunto').send_keys(f'Processo SEI nº {n_processo} - Aprovado ({codigo_rastreio})')
#     select_element = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="selTextoPadrao"]')))
#     select = Select(select_element)
#     #SELECIONA O TEXTO PADRAO DE DRONES RETIDOS APROVADOS
#     select.select_by_visible_text('Processo SEI ORCN - APROVAÇÃO - DRONE RETIDO')
#     #ENVIA EMAIL
#     navegador.find_element(By.XPATH, '//*[@id="divInfraBarraComandosInferior"]/button[1]').click()
#     #FECHA ALERTA DO NAVEGADOR
#     time.sleep(0.5)
#     alert = Alert(navegador)
#     alert.accept()
#     time.sleep(2)
#     # chrome=pyautogui.locateOnScreen('imagensAut/chrome.png', confidence=0.7)
#     # pyautogui.click(chrome)
#     navegador.switch_to.window(janela_principal)

# options = webdriver.ChromeOptions()
# options.add_argument("--disable-popup-blocking") 
# #DEFINE O TEMPO DE EXECUÇÃO PARA CADA COMANDO DO PYAUTOGUI
# pyautogui.PAUSE = 0.7
# #INICIA O NAVEGADOR
# # Caminho do ChromeDriver local
# # chrome_driver_path = r"main\chromedriver-win64\chromedriver.exe"  # Altere para o caminho correto do seu ChromeDriver
# # Configura o serviço do ChromeDriver
# servico = Service(ChromeDriverManager().install())
# # Inicia o navegador Chrome
# navegador = webdriver.Chrome(service=servico, options=options)
# navegador.maximize_window()
# #ENTRA NO SEI
# navegador.get('https://sei.anatel.gov.br/')
# janela_principal = navegador.current_window_handle

# while True:
#     input("Logue no SEI e pressione enter")
#     break

# for processo, codigo_rastreio in processos_testar.items():
#     print('\nProcessando processo', processo)
#     fc.vai_para_processo(navegador, processo)
#     processo_reaberto = reabre_processo()
#     navegador.switch_to.frame('ifrArvore')
#     try:
#         fc.clica_noelemento(navegador, By.PARTIAL_LINK_TEXT, "Despacho Decisório")
#         manda_email(processo, codigo_rastreio, janela_principal)
#         retido='Sim'
#         time.sleep(0.2)
#     except Exception as e:
#         print(e)
#     if processo_reaberto:
#         fecha_processo()
#         print('Processo fechado')

pyautogui.hotkey("win", "ll")
