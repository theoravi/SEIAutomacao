#IMPORTS NECESSÁRIOS PARA O FUNCIONAMENTO DO CÓDIGO
import numpy as np
import os
import sys
import subprocess
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
import funcoes as fc

#python -m PyInstaller --onefile theomation.py
def main():
    reset = True
    while reset:
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
        pyautogui.hotkey('ctrl', 'l')
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
                login_button = tk.Button(frame, text="Entrar", font=('Arial', 10), command=fc.preencher_campos, width=20)
                login_button.grid(row=6, columnspan=2)

                #ASSOCIA A TECLA ENTER DO TECLADO PARA APERTA O BOTAO DE LOGIN
                root.bind('<Return>', lambda event: fc.preencher_campos(user_var, senha_var, navegador, root))

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
            if fc.check_element_exists(By.XPATH,'//*[@id="divFiltro"]/div[2]/a', navegador):
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
        resultados = list(executor.map(fc.processa_tr, trs))
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

        #COLETA NOME DO ESTAGIARIO
        nomeEstag=input('Insira seu nome completo: ')
        nomeEstag_sem_acento = unidecode.unidecode(nomeEstag)

        #COLETA CAMINHO DA PLANILHA DE DRONES CONFORMES
        planilhaDrones = input("Insira o caminho da planilha/lista de drones conformes: ")
        #RETIRA ASPAS CASO EXISTA NO CAMINHO DA PLANILHA
        planilhaDrones = planilhaDrones.replace('"' , '')
        while True:
            #MOSTRA OPÇÕES DE EXECUCAO PARA O USUARIO
            print("O que deseja fazer?")
            print("Digite [1] para analisar processo e criar despacho.")
            print("Digite [2] para concluir processos assinados.")
            print("Digite [3] para atribuir processos para si.") 
            print("Digite [4] para analisar um processo específico.")
            print("Digite [5] para reiniciar o programa")
            print("Digite [6] para encerrar o programa")
            opcoes = int(input("Opção: "))
            if opcoes == 1:
                #EXECUTA FUNCAO PARA ANALISAR PROCESSO
                fc.analisaListaDeProcessos(navegador, lista_processos, nomeEstag, planilhaDrones)
                print("Análise finalizada.")
            elif opcoes == 2:
                #EXECUTA FUNCAO PARA CONCLUIR PROCESSO
                fc.concluiProcesso(navegador, lista_procConformes)
                print("Todos os processos foram concluídos.")
            elif opcoes == 3:
                try:
                    #EXECUTA FUNCAO PARA ATRIBUIR PROCESSOS
                    fc.atribuicao(navegador, nomeEstag_sem_acento, nomeEstag)
                except Exception as e:
                    print(f"Ocorreu um erro! {e}")
            elif opcoes == 4:
                #EXECUTA FUNCAO PARA ANALISAR  UM ÚNICO PROCESSO
                fc.analisaApenasUmProcesso(navegador, nomeEstag, planilhaDrones)
            elif opcoes == 5:
                print("Reiniciando programa")
                reset = True
                break
            elif opcoes == 6:
                #ENCERRA O PROGRAMA
                print("Encerrando programa")
                reset = False
                break
            else:
                #CONDICAO DE ERRO
                print("Opção inválida, tente novamente.")
        navegador.quit()

if __name__ == "__main__":
    main()