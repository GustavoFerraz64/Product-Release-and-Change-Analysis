from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import os
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import sys
import shutil

os.path.expanduser('~')

class ExtraiAIT:
    def __init__(self, caminho_downloads, arquivo_lap, caminho_lap):
        self.caminho_downloads = caminho_downloads
        self.arquivo_lap = arquivo_lap
        self.caminho_lap = caminho_lap
        self.link_ait = r'https://ait.br.tkelevator.com/scripts/gisprod.pl/ait/ait_login.html'
        try:
            self.service = Service(ChromeDriverManager().install())
        except:
            os.environ['WDM_SSL_VERIFY'] = '0'
            self.service = Service(ChromeDriverManager().install())

        self.driver = webdriver.Chrome(service=self.service)
        self.driver.maximize_window()

    def limpa_pasta_download(self):
        '''Função que limpa o arquivo RelatórioLAP.csv da pasta de downloads do usuário, 
        caso o arquivo exista.'''

        #Para remover RelatorioLAP no formato CSV
        try:
            os.remove(self.caminho_downloads + "\\" + self.arquivo_lap)
        except:
            pass

        #Para remover RelatorioLAP no formato XLSX (Excessões)
        try:
            os.remove(self.caminho_downloads + "RelatorioLAP.xlsx")
        except:
            pass
    
    def ler_login(self):
        with open(r'\\srvfile01.tsur.local\DADOS_PBI\Compartilhado_BI\DPCP\9. DEPM\Bases de Dados\Ω Parâmetros Scripts\login gustavo.txt') as file:
            self.user = file.readline().strip()
            self.key = file.readline().strip()

    def entra_ait(self):
        self.driver.get(r'https://ait.br.tkelevator.com/scripts/gisprod.pl/ait/ait_login.html')
        
        #Campo usuário
        self.driver.find_element(By.ID, 'usuario').clear()
        self.driver.find_element(By.ID, 'usuario').send_keys(self.user)

        #Campo senha
        self.driver.find_element(By.ID, 'chave').clear()
        self.driver.find_element(By.ID, 'chave').send_keys(self.key)

        #Clica no botão de entrar
        self.driver.find_element(By.XPATH, '//*[@id="generico"]/div/table/tbody/tr[4]/td[2]/button').click()
    
    def navegacao(self):
        '''Navega pelo AIT até achar o local onde está o Relatório LAP'''
        wait = WebDriverWait(self.driver, timeout=50)

        #Contenção utilizada no dia 31/01/2025 para fechar a mensagem de alerta do Service Desk.
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[1]/div/button')))
            self.driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/button').click()
        except:
            pass

        #Vai até a página de download do Relatório LAP
        self.driver.find_element(By.ID, 'atalho').clear()
        self.driver.find_element(By.ID, 'atalho').send_keys('22479')
        
        self.driver.find_element(By.XPATH, '//*[@id="divExecutarPrograma"]/input[3]').click()


        iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, 'iframe')))
        self.driver.switch_to.frame(iframe)

        elemento_download = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="excel"]/td[1]/a/span')))

        self.driver.find_element(By.XPATH, '//*[@id="excel"]/td[1]/a/span').click()

    def verifica_se_download_concluido(self):
        '''Verifica se o download foi concluído.
        O tempo de espera máximo é de um minuto, caso tenha passado o tempo e o arquivo não esteja na pasta de downloads do usuário,
        o script é encerrado.'''
        tempo_inicio = time.time()

        while True:
            if os.path.exists(self.caminho_downloads + "\\" + self.arquivo_lap):
                print('Download concluído!')
                break
            if time.time() - tempo_inicio > 60: #Tempo máximo: 1 minuto
                print('Tempo limite de download excedido. O script será encerrado')
                sys.exit()
            time.sleep(1)
    
    def fecha_navegador(self):
        '''Após o download do Relatório LAP, o navegador é fechado.'''
        self.driver.quit()




                
