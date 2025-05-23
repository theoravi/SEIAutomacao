#Para baixar as bibliotecas, execute o arquivo bibs_install.bat

#Para gerar o executável, execute o comando abaixo:
#python -m PyInstaller --icon="icons\T_logo-1-.ico" --onefile theomation.py --clean

#IMPORTS NECESSÁRIOS PARA O FUNCIONAMENTO DO CÓDIGO
import unidecode
#from concurrent.futures import ThreadPoolExecutor
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
import time
import funcoes as fc
import json
import subprocess
from datetime import datetime

#DEV_MODE = True
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
        
        lista_procConformes, lista_procCancelamento, lista_processos = fc.fazListasProcessos(navegador)
            
        #IMPRIME A QUANTIDADE DE PROCESSOS QUE SERÃO ANALISADOS E CONCLUÍDOS
        print("Quantidade de processos para analisar", len(lista_processos))
        print("Quantidade de processos para concluir", len(lista_procConformes))
        print("Quantidade de pedidos de cancelamento a concluir", len(lista_procCancelamento))

        

        #ABRE DICIONARIO
        with open('usuarios/usuarios.json', 'r') as arquivo:
            usuarios = json.load(arquivo)
        try: 
            nomeEstag = usuarios[user_name]
        except:
            while True:
                nomeEstag = str(input("Como é seu primeiro acesso, digite seu nome completo: "))
                opcao = str(input("Caso esteja tudo certo, digite [1]. Caso contrario, digite [2]: "))
                if opcao == '1':
                    usuarios[user_name] = nomeEstag
                    break
            with open('usuarios/usuarios.json', 'w') as arquivo:
                json.dump(usuarios, arquivo)
        finally:
            nomeEstag_sem_acento = unidecode.unidecode(nomeEstag)
            apelidoEstag = nomeEstag.split()[0]
            now = datetime.now()
            if now.hour < 12:
                print(f"Bom dia, {apelidoEstag}!")
            elif now.hour < 18:
                print(f"Boa tarde, {apelidoEstag}!")
            else:
                print(f"Boa noite, {apelidoEstag}!")

        #COLETA CAMINHO DAS PLANILHAS DE EQUIPAMENTOS CONFORMES
   
        planilhaDrones = f"C:\\Users\\{user_name}\\ANATEL\\ORCN - Drones\\Lista de Drones Anatel_Corrigida.xlsx"
        planilhaRadios = f"C:\\Users\\{user_name}\\ANATEL\\ORCN - Rádios\\Lista Radiamador.xlsx"
        planilhaGeral2025 = f"C:\\Users\\{user_name}\\ANATEL\\ORCN - DRONES SEI PLANILHA\\Distribuição Processo Drone_2025.xlsx"
        planilhaGeral2024 = f"C:\\Users\\{user_name}\\ANATEL\\ORCN - DRONES SEI PLANILHA\\Distribuição Processo Drone_2024.xlsx"
        
        # #IMPRIME CAMINHOS ENCONTRADOS
        # print("Coletado caminho da planilha de drones conformes", planilhaDrones)
        # print("Coletado caminho da planilha de rádios conformes", planilhaRadios)
        # print("Coletado caminho da planilha geral", planilhaGeral)

        while True:
            #MOSTRA OPÇÕES DE EXECUCAO PARA O USUARIO
            print("O que deseja fazer?",
                  "Digite [1] para analisar processo e criar despacho.",
                  "Digite [2] para concluir processos assinados.",
                  "Digite [3] para atribuir processos para si.",
                  "Digite [4] para analisar um processo específico.",
                  "Digite [5] para refazer a lista de processos a analisar e a concluir.",
                  "Digite [6] para reiniciar o programa",
                  "Digite [7] para encerrar o programa",
                  sep='\n'
                  )
            opcoes = str(input("Opção: "))
            if opcoes == '1':
                #EXECUTA FUNCAO PARA ANALISAR PROCESSO
                fc.analisaListaDeProcessos(navegador, lista_processos, nomeEstag, planilhaDrones, planilhaRadios, planilhaGeral2025)
                print("Análise finalizada.")        
            elif opcoes == '2':
                #EXECUTA FUNCAO PARA CONCLUIR PROCESSO
                fc.concluiProcesso(navegador, lista_procConformes, lista_procCancelamento, nomeEstag, planilhaGeral2024, planilhaGeral2025)
                print("Todos os processos foram concluídos.")
            elif opcoes == '3':
                try:
                    #EXECUTA FUNCAO PARA ATRIBUIR PROCESSOS
                    fc.atribuicao(navegador, nomeEstag_sem_acento, nomeEstag, planilhaGeral2025)
                except Exception as e:
                    print(f"Ocorreu um erro! {e}")

                # Clica no icone que mostra a lista de processos 
                navegador.switch_to.default_content()
                fc.clica_noelemento(navegador, By.XPATH, '//*[@id="lnkControleProcessos"]')

                # Clica no filtro de processsos atribuidos ao usuario
                try:
                    fc.clica_noelemento(navegador, By.XPATH,'//*[@id="divFiltro"]/div[2]/a', 2)
                except:
                    pass
                lista_procConformes, lista_procCancelamento, lista_processos = fc.fazListasProcessos(navegador)
                print("Lista de processos a analisar e a concluir refeita:")
                print("Quantidade de processos para analisar", len(lista_processos))
                print("Quantidade de processos para concluir", len(lista_procConformes))
                print("Quantidade de pedidos de cancelamento a concluir", len(lista_procCancelamento))
            elif opcoes == '4':
                #EXECUTA FUNCAO PARA ANALISAR  UM ÚNICO PROCESSO
                fc.analisaApenasUmProcesso(navegador, nomeEstag, planilhaDrones, planilhaRadios, planilhaGeral2025)
            elif opcoes == '5':
                # Clica no icone que mostra a lista de processos 
                navegador.switch_to.default_content()
                fc.clica_noelemento(navegador, By.XPATH, '//*[@id="lnkControleProcessos"]')

                # Clica no filtro de processsos atribuidos ao usuario
                try:
                    fc.clica_noelemento(navegador, By.XPATH,'//*[@id="divFiltro"]/div[2]/a', 2)
                except:
                    pass
                lista_procConformes, lista_procCancelamento, lista_processos = fc.fazListasProcessos(navegador)
                print("Lista de processos a analisar e a concluir refeita:")
                print("Quantidade de processos para analisar", len(lista_processos))
                print("Quantidade de processos para concluir", len(lista_procConformes))
                print("Quantidade de pedidos de cancelamento a concluir", len(lista_procCancelamento))
            elif opcoes == '6':
                print("Reiniciando programa")
                reset = True
                break
            elif opcoes == '7':
                #ENCERRA O PROGRAMA
                print("Encerrando programa")
                reset = False
                break
            else:
                #CONDICAO DE ERRO
                print("Opção inválida, tente novamente.")
        navegador.quit()

if __name__ == "__main__":
    # if not DEV_MODE:
        # # Chamar o script de atualização antes de executar a automação
        # subprocess.run(["python", "updater.py"])
        # # Continuar com a lógica principal da automação
        # print("Iniciando a automação...")
        # # Coloque aqui o código principal da sua automação
    main()