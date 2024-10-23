#Para baixar as bibliotecas, execute o arquivo bibs_install.bat

#Para gerar o executável, execute o comando abaixo:
#python -m PyInstaller --icon="T_logo-1-.ico" --onefile theomation.py --clean


#IMPORTS NECESSÁRIOS PARA O FUNCIONAMENTO DO CÓDIGO
import unidecode
from concurrent.futures import ThreadPoolExecutor
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
import time
import funcoes as fc
import json

def main():
    reset = True
    while reset:
        navegador = fc.abreChromeEdge()
        while True:
            #INICIA JANELA
            try:
                fc.iniciaJanela(navegador)
                user_name = fc.user_name
            #CONDICAO DE ERRO PARA CASO O USUÁRIO ERRE O SEU LOGIN 
            except Exception:
                print('Ocorreu um erro, tente novamente.')
            try:
                #FECHA ALERTA DO NAVEGADOR
                time.sleep(1)
                alert = Alert(navegador)
                alert.accept()
            except:
                #ENTRAR NA CAIXA DE PROCESSOS ATRIBUIDOS AO USUARIO
                if fc.check_element_exists(By.XPATH,'//*[@id="divFiltro"]/div[2]/a', navegador):
                    navegador.find_element(By.XPATH,'//*[@id="divFiltro"]/div[2]/a').click()
                    break
                #CASO NAO ENCONTRE O FILTRO DE PROCESSOS ELE INFORMA QUE NAO ENTROU NO SEI
                else:
                    #APAGA O USUÁRIO
                    navegador.find_element(By.XPATH,'//*[@id="txtUsuario"]').clear()
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

        #ABRE DICIONARIO
        with open('usuarios/usuarios.json', 'r') as arquivo:
            usuarios = json.load(arquivo)

        try: 
            nomeEstag = usuarios[user_name]
            print("Coletado nome do usuário:", nomeEstag)
        except:
            while True:
                nomeEstag = str(input("Como é seu primeiro acesso, digite seu nome completo: "))
                opcao = int(input("Caso esteja tudo certo, digite [1]. Caso contrario, digite [2]: "))
                if opcao == 1:
                    usuarios[user_name] = nomeEstag
                    break
            with open('usuarios/usuarios.json', 'w') as arquivo:
                json.dump(usuarios, arquivo)
        finally:
            nomeEstag_sem_acento = unidecode.unidecode(nomeEstag)

        #COLETA CAMINHO DAS PLANILHAS DE EQUIPAMENTOS CONFORMES
   
        planilhaDrones = f"C:\\Users\\{user_name}\\ANATEL\\ORCN - Drones\\Lista de Drones Anatel_Corrigida.xlsx"
        planilhaRadios = f"C:\\Users\\{user_name}\\ANATEL\\ORCN - Rádios\\Lista Radiamador.xlsx"
        planilhaGeral = f"C:\\Users\\{user_name}\\ANATEL\\ORCN - DRONES SEI PLANILHA\\Distribuição Processo Drone.xlsx"

        # #IMPRIME CAMINHOS ENCONTRADOS
        # print("Coletado caminho da planilha de drones conformes", planilhaDrones)
        # print("Coletado caminho da planilha de rádios conformes", planilhaRadios)
        # print("Coletado caminho da planilha geral", planilhaGeral)

        while True:
            #MOSTRA OPÇÕES DE EXECUCAO PARA O USUARIO
            print("O que deseja fazer?\n",
                  "Digite [1] para analisar processo e criar despacho.\n",
                  "Digite [2] para concluir processos assinados.\n",
                  "Digite [3] para atribuir processos para si.\n",
                  "Digite [4] para analisar um processo específico.\n",
                  "Digite [5] para reiniciar o programa\n",
                  "Digite [6] para encerrar o programa\n",
                  )
            opcoes = int(input("Opção: "))
            if opcoes == 1:
                #EXECUTA FUNCAO PARA ANALISAR PROCESSO
                fc.analisaListaDeProcessos(navegador, lista_processos, nomeEstag, planilhaDrones, planilhaRadios)
                print("Análise finalizada.")
            elif opcoes == 2:
                #EXECUTA FUNCAO PARA CONCLUIR PROCESSO
                fc.concluiProcesso(navegador, lista_procConformes, nomeEstag, planilhaGeral)
                print("Todos os processos foram concluídos.")
            elif opcoes == 3:
                try:
                    #EXECUTA FUNCAO PARA ATRIBUIR PROCESSOS
                    fc.atribuicao(navegador, nomeEstag_sem_acento, nomeEstag, planilhaGeral)
                except Exception as e:
                    print(f"Ocorreu um erro! {e}")
            elif opcoes == 4:
                #EXECUTA FUNCAO PARA ANALISAR  UM ÚNICO PROCESSO
                fc.analisaApenasUmProcesso(navegador, nomeEstag, planilhaDrones, planilhaRadios)
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