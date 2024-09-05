import pyautogui
import time
import pyperclip
processos = "53500.068611/2024-66"

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
time.sleep(.3)
#COPIA O NUMERO DO PROCESSO NA CELULA PARA SABER SE ELE FOI ENCONTRADO CORRETAMENTE
pyautogui.hotkey('ctrl','c')
time.sleep(0.5)
n_processoConfirmation = pyperclip.paste()
n_processoConfirmation = n_processoConfirmation.replace('\n', '')
n_processoConfirmation = n_processoConfirmation.strip('"')
print(n_processoConfirmation)
print(type(n_processoConfirmation))
if n_processoConfirmation != processos:
    while True:
        procure_manualmente = int(input("Houve um problema para encontrar o processo. Busque o manualmente e digite 1: "))
        if procure_manualmente == 1:
            time.sleep(1)
            pyautogui.click(edge)
            time.sleep(0.7)
            break
        else:
            print("Opcao incorreta")  
#NAVEGA ENTRE AS CELULAS E INSERE OS DADOS SOBRE O PROCESSO        
pyautogui.press('right')
pyautogui.PAUSE = 0.3
pyperclip.copy('teste')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('right')
pyperclip.copy('teste')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('right')