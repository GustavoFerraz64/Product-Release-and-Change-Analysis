from ExtracaoAIT import *
from AnaliseLAP import *
from Email import *
import warnings
warnings.filterwarnings("ignore")

class App:
    '''Este código tem como objetivo extrair automaticamente a base de dados das LAPs pelo AIT, realizar análise de dados, atualizar arquivo Excel destinado à justificativa
    de atrasos LAPs, e enviar e-mail aos departamentos em relação às LAPs atrasadas

    @autor: Gustavo Nunes Ferraz
    @departamento: DPCP
    @datas de modificações:
        -19/12/2024: parte de análise de dados e envio de e-mail ao DPCP concluída
        -07/01/2025: Foi realizada uma organização geral do código, e acrescentada a etapa de extração do relatório no AIT + envio de e-mail aos departamentos restantes concluída;
        -09/01/2025: Foi incluso o README.md, detalhando o funcionamento do script.
        -21/01/2025: Envio dos e-mails aos departamentos funcionando normalmente.
        -29/01/2025: Alterado a lógica da definição de prazos para cada departamento. Foi informado que o DEQF sempre terá como Pai o DCOM, portanto a coluna Pai do Relatório LAP é manipulado
        quando uma LAP que possuí DEQF e DCOM juntas, garantindo que o DCOM sempre fique em um nível abaixo de DEQF.
        -31/01/2025: Adicionado na parte extracao_ait.navegacao() o comando de fechar o aviso informando a substituição da Service Desk TI pela ServiceNow.
    '''
    def __init__(self):
        self.caminho_downloads = os.path.expanduser('~') + "\\" + 'Downloads'
        self.caminho_lap = r'\\srvfile01.tsur.local\DADOS_PBI\Compartilhado_BI\DPCP\25. LAPs'
        self.arquivo_lap = 'RelatorioLAP.csv'
        self.outlook = win32com.client.Dispatch('outlook.application')

    def executar(self):
        print("Iniciando o Script")
        print("Extraindo LAPs do AIT...")
        extracao_ait = ExtraiAIT(self.caminho_downloads, self.arquivo_lap, self.caminho_lap)
        extracao_ait.limpa_pasta_download()
        extracao_ait.ler_login()
        extracao_ait.entra_ait()
        extracao_ait.navegacao()
        extracao_ait.verifica_se_download_concluido()
        extracao_ait.fecha_navegador()

        analise_lap = AnaliseLAP(self.caminho_downloads, self.caminho_lap, self.arquivo_lap)
        print("Lendo calendário TKE...")
        analise_lap.ler_calendario()
        analise_lap.definir_dias_uteis()
        print("Lendo o arquivo relatório LAP...")
        analise_lap.ler_arquivo_relatorio_lap()
        print("Tratando datas...")
        analise_lap.trata_datas()
        print("Analisando os dados das LAPs...")
        analise_lap.calcula_lead_time()
        analise_lap.define_data_liberacao_lap()
        analise_lap.define_atraso_engenharia()
        analise_lap.define_numero_departamentos_por_lap()
        print("Analisando os prazos das LAPs...")
        analise_lap.cria_coluna_prazo()
        analise_lap.ajustar_prazo(analise_lap.df_relatorio_lap['Prazo LAP'])
        analise_lap.definir_laps_em_fluxo_e_atraso()
        analise_lap.definir_liberadas_em_atraso()
        analise_lap.definir_atraso_laps_liberadas()
        print("Calculando lead time por departamento...")
        analise_lap.calcula_lt_por_departamento()
        print("Definindo prazos por departamento...")
        analise_lap.define_prazos_departamentos()
        analise_lap.ajustar_prazo(analise_lap.df_relatorio_lap['Prazo Departamento'])
        analise_lap.define_departamentos_atrasados()
        print("Atualizando a planilha de justificativas de atraso...")
        analise_lap.atualiza_planilha_justificativa_atrasos()
        print("Gravando o Relatório da Análise...")
        analise_lap.gravar_relatorio_analisado()

        print("Iniciando o envio de e-mails...")
        envio_email = EnvioEmail(analise_lap.df_relatorio_lap, analise_lap.df_justificativa_atrasos, self.outlook)
        envio_email.desconsidera_laps()
        envio_email.trata_prazo_departamentos()
        envio_email.envia_email_pcp()
        envio_email.ler_emails_destinatarios()
        envio_email.cria_rotulo_laps()
        envio_email.envia_email_demais_departamentos()
        print("Script finalizado com sucesso")
        time.sleep(3)