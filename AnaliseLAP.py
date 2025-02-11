import pandas as pd
import datetime
import numpy as np
import math
from tqdm import tqdm


class AnaliseLAP:
    def __init__(self, caminho_download, caminho_lap, arquivo_lap):
        self.caminho_emails = r'\\srvfile01.tsur.local\DADOS_PBI\Compartilhado_BI\DPCP\9. DEPM\Bases de Dados\E-mails\LaP Report.xlsx'
        self.caminho_calendario = r'\\srvfile01.tsur.local\DADOS_PBI\Compartilhado_BI\DPCP\9. DEPM\Bases de Dados\Calendário'
        self.caminho_download= caminho_download
        self.arquivo_calendario = r'Calendário TKE.xlsx'
        self.hoje = pd.Timestamp(datetime.datetime.today())
        self.caminho_lap = caminho_lap
        self.arquivo_lap = arquivo_lap
        self.arquivo_relatorio_analisado = r'RelatorioLAP_Analisado.xlsx'
        self.arquivo_justificativa_atrasos = r'Justificativa de LAPs em Atraso.xlsx'
        self.prazo_adequacao_produto = 60
        self.prazo_correcao_critica = 20
        self.prazo_liberacao_produto_novo = 60
        self.prazo_tabela_aplicacao = 45
        self.prazo_homologacao = 60

    def ler_calendario(self):
        '''Função que lê o calendário da TKE'''
        self.df_calendario_tke = pd.read_excel(self.caminho_calendario + "\\" + self.arquivo_calendario)

    def definir_dias_uteis(self):
        '''Criar Series de dias úteis TKE'''
        self.dias_uteis = self.df_calendario_tke[self.df_calendario_tke['Dia útil'] == 1]['Data']
        self.dias_uteis = pd.to_datetime(self.dias_uteis, format='%d/%m/%Y')
        self.dias_uteis = self.dias_uteis.to_list()
        self.dias_uteis.append(None)

    def ler_arquivo_relatorio_lap(self):
        '''Lê o arquivo de relatório LAP baixado pelo AIT.
        Espera-se que o arquivo esteja na pasta de Downloads do usuário.'''
        try:
            self.df_relatorio_lap = pd.read_csv(self.caminho_download + "\\" + self.arquivo_lap, sep=';', encoding='latin-1')
        except:
            print("Erro ao ler o arquivo RelatórioLAP.csv")
            raise FileNotFoundError

    def trata_datas(self):
        '''Transformar colunas de data do Data Frame em Datetime.
        Valores estranhos são substituídos por None'''
        self.df_relatorio_lap['Previsão Liberação da Engenharia'] = self.df_relatorio_lap['Previsão Liberação da Engenharia'].replace('?', None)
        self.df_relatorio_lap['Previsão Liberação da Engenharia'] = self.df_relatorio_lap['Previsão Liberação da Engenharia'].replace('05/04/0202', None)
        self.df_relatorio_lap['Previsão Liberação da Engenharia'] = self.df_relatorio_lap['Previsão Liberação da Engenharia'].replace('02/09/0222', None)
        self.df_relatorio_lap['Data Prev EA'] = pd.to_datetime(self.df_relatorio_lap['Data Prev EA'], format='%d/%m/%y', dayfirst=True)
        self.df_relatorio_lap['Previsão Liberação da Engenharia'] = pd.to_datetime(self.df_relatorio_lap['Previsão Liberação da Engenharia'], format='%d/%m/%Y', dayfirst=True)
        self.df_relatorio_lap['Data de liberação'] = pd.to_datetime(self.df_relatorio_lap['Data de liberação'], format='%d/%m/%Y', dayfirst=True)
        self.df_relatorio_lap['Data Lib EA'] = pd.to_datetime(self.df_relatorio_lap['Data Lib EA'], format='%d/%m/%y', dayfirst=True)
        self.df_relatorio_lap['Dt. Abertura'] = pd.to_datetime(self.df_relatorio_lap['Dt. Abertura'], format='%d/%m/%Y')
    def calcula_lead_time(self):
        '''Função que calcula o Lead Time da LAP, subtraindo a coluna de liberação do DPCP e a coluna de liberação da Engenharia'''
        self.df_relatorio_lap['Lead Time da LAP'] = None
        self.df_laps_finalizadas = self.df_relatorio_lap[(self.df_relatorio_lap['Depto'] == 'DPCP') & (pd.notna(self.df_relatorio_lap['Data Lib EA'])) & (self.df_relatorio_lap['Status'] != 'CANCELADA')]
        self.codigo_laps_finalizadas = self.df_laps_finalizadas['Código']
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['Código'].isin(self.codigo_laps_finalizadas)) & (self.df_relatorio_lap['Depto'] == 'DPCP'), 'Lead Time da LAP'] = self.df_relatorio_lap[self.df_relatorio_lap['Código'].isin(self.codigo_laps_finalizadas)]['Data Lib EA'] - self.df_relatorio_lap[self.df_relatorio_lap['Código'].isin(self.codigo_laps_finalizadas)]['Data de liberação']

        for index, row in self.df_laps_finalizadas.iterrows():
            self.df_relatorio_lap.loc[self.df_relatorio_lap['Código'] == row['Código'], 'Lead Time da LAP'] = self.df_relatorio_lap[self.df_relatorio_lap['Código'] == row['Código']]['Lead Time da LAP'].ffill()

        self.df_relatorio_lap.loc[(pd.isna(self.df_relatorio_lap['Lead Time da LAP'])) & (pd.notna(self.df_relatorio_lap['Data de liberação'])), 'Lead Time da LAP'] = 'LAP Em Fluxo'
        self.df_relatorio_lap.loc[pd.isna(self.df_relatorio_lap['Data de liberação']), 'Lead Time da LAP'] = 'Pendente de Liberação da Engenharia'

    def define_data_liberacao_lap(self):
        '''Função que cria coluna de liberação da LAP, se baseia na data de liberação da LAP pelo DPCP'''
        self.df_laps_finalizadas = self.df_laps_finalizadas.rename(columns={'Data Lib EA': 'Data de Finalização da LAP'})
        self.df_relatorio_lap = self.df_relatorio_lap.merge(self.df_laps_finalizadas[['Código', 'Data de Finalização da LAP']], how='left', on='Código')

    def define_atraso_engenharia(self):
        '''Cria coluna booleana indicando se houve atraso pela engenharia'''
        #self.df_relatorio_lap['Atraso Liberação Engenharia'] = (self.df_relatorio_lap['Previsão Liberação da Engenharia'] < self.df_relatorio_lap['Data de liberação']) & (self.df_relatorio_lap['Previsão Liberação da Engenharia'] < self.hoje)
        self.df_relatorio_lap['Atraso Liberação Engenharia'] = None
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['Previsão Liberação da Engenharia'] < self.df_relatorio_lap['Data de liberação']) & (self.df_relatorio_lap['Previsão Liberação da Engenharia'] < self.hoje), 'Atraso Liberação Engenharia'] = 'Liberado em Atraso'
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['Previsão Liberação da Engenharia'] < self.df_relatorio_lap['Data de liberação']) & (self.df_relatorio_lap['Previsão Liberação da Engenharia'] < self.hoje), 'Atraso Liberação Engenharia'] = 'Liberado em Atraso'
        self.df_relatorio_lap['Atraso Liberação Engenharia'] = (self.df_relatorio_lap['Previsão Liberação da Engenharia'] < self.df_relatorio_lap['Data de liberação']) & (self.df_relatorio_lap['Previsão Liberação da Engenharia'] < self.hoje)

    def define_numero_departamentos_por_lap(self):
        '''Define o número de departamentos envolvidos em cada código de LAP'''
        self.dpto_por_lap = self.df_relatorio_lap.groupby('Código')['Depto'].count()
        self.dpto_por_lap.name = 'Número de departamentos por LAP'
        self.df_relatorio_lap = self.df_relatorio_lap.merge(self.dpto_por_lap, how='left', on='Código')
 
    def cria_coluna_prazo(self):
        '''Gerar a coluna de prazo máximo de finalização da LAP, conforme o definido pelo DPCP'''
        self.df_relatorio_lap['Prazo LAP'] = None
        self.df_relatorio_lap.loc[(pd.notna(self.df_relatorio_lap['Data de liberação'])) & (self.df_relatorio_lap['Motivo'] == 'ADEQUACAO DO PRODUTO'), 'Prazo LAP'] = self.df_relatorio_lap[(pd.notna(self.df_relatorio_lap['Data de liberação'])) & (self.df_relatorio_lap['Motivo'] == 'ADEQUACAO DO PRODUTO')]['Data de liberação'] + datetime.timedelta(self.prazo_adequacao_produto)
        self.df_relatorio_lap.loc[(pd.notna(self.df_relatorio_lap['Data de liberação'])) & (self.df_relatorio_lap['Motivo'] == 'CORRECAO CRITICA'), 'Prazo LAP'] = self.df_relatorio_lap[(pd.notna(self.df_relatorio_lap['Data de liberação'])) & (self.df_relatorio_lap['Motivo'] == 'CORRECAO CRITICA')]['Data de liberação'] + datetime.timedelta(self.prazo_correcao_critica)
        self.df_relatorio_lap.loc[(pd.notna(self.df_relatorio_lap['Data de liberação'])) & (self.df_relatorio_lap['Motivo'] == 'LIBERACAO DE PRODUTO NOVO'), 'Prazo LAP'] = self.df_relatorio_lap[(pd.notna(self.df_relatorio_lap['Data de liberação'])) & (self.df_relatorio_lap['Motivo'] == 'LIBERACAO DE PRODUTO NOVO')]['Data de liberação'] + datetime.timedelta(self.prazo_liberacao_produto_novo)
        self.df_relatorio_lap.loc[(pd.notna(self.df_relatorio_lap['Data de liberação'])) & (self.df_relatorio_lap['Motivo'] == 'TABELA DE APLICAÇÃO'), 'Prazo LAP'] = self.df_relatorio_lap[(pd.notna(self.df_relatorio_lap['Data de liberação'])) & (self.df_relatorio_lap['Motivo'] == 'TABELA DE APLICAÇÃO')]['Data de liberação'] + datetime.timedelta(self.prazo_tabela_aplicacao)
        self.df_relatorio_lap.loc[(pd.notna(self.df_relatorio_lap['Data de liberação'])) & (self.df_relatorio_lap['Motivo'] == 'HOMOLOGACAO'), 'Prazo LAP'] = self.df_relatorio_lap[(pd.notna(self.df_relatorio_lap['Data de liberação'])) & (self.df_relatorio_lap['Motivo'] == 'HOMOLOGACAO')]['Data de liberação'] + datetime.timedelta(self.prazo_homologacao)

    # def ajustar_prazo(self):
    #     '''Ajustar as datas de Prazo LAP, caso o prazo caia em um dia não útil, diminuir o número de datas até o dia ser útil'''
    #     while True:
    #         self.df_relatorio_lap['Prazo LAP'] = np.where(
    #             self.df_relatorio_lap['Prazo LAP'].isin(self.dias_uteis),
    #             self.df_relatorio_lap['Prazo LAP'],
    #             self.df_relatorio_lap['Prazo LAP'] - datetime.timedelta(-1)
    #         )
    #         if all(self.df_relatorio_lap['Prazo LAP'].isin(self.dias_uteis)):
    #             break
    
    def ajustar_prazo(self, coluna_data):
        '''Ajustar as datas de Prazo LAP, caso o prazo caia em um dia não útil, diminuir o número de datas até o dia ser útil'''
        coluna_data = pd.Series(coluna_data)
        while True:
            coluna_data = np.where(coluna_data.isin(self.dias_uteis), coluna_data, coluna_data - datetime.timedelta(-1))
            coluna_data = pd.Series(coluna_data)
            if all(coluna_data.isin(self.dias_uteis)):
                break
        

    def definir_laps_em_fluxo_e_atraso(self):
        '''Definir as LAPS que estão em fluxo e que estão em atraso'''
        self.df_relatorio_lap['LAP no Fluxo e em Atraso'] = None
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['Lead Time da LAP'] == 'LAP Em Fluxo') & (self.df_relatorio_lap['Prazo LAP'] < self.hoje) & (pd.notna(self.df_relatorio_lap['Prazo LAP'])), 'LAP no Fluxo e em Atraso'] = 'Em Atraso'
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['Lead Time da LAP'] == 'LAP Em Fluxo') & (self.df_relatorio_lap['Prazo LAP'] >= self.hoje) & (pd.notna(self.df_relatorio_lap['Prazo LAP'])), 'LAP no Fluxo e em Atraso'] = 'Dentro do Prazo'
        self.df_relatorio_lap.loc[((self.df_relatorio_lap['Status'] == 'LIBERADA') | (self.df_relatorio_lap['Status'] == 'CANCELADA')), 'LAP no Fluxo e em Atraso'] = 'LAP Liberada/Cancelada'
        self.df_relatorio_lap.loc[pd.isna(self.df_relatorio_lap['Prazo LAP']), 'LAP no Fluxo e em Atraso'] = 'Prazo Não Definido para Categoria da LAP'

    def definir_liberadas_em_atraso(self):
        '''Criar rótulo das LAPs entregue em atraso'''
        self.df_relatorio_lap['LAP Entregue em Atraso'] = None
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['Lead Time da LAP'] != 'LAP Em Fluxo') & (self.df_relatorio_lap['Lead Time da LAP'] != 'Pendente de Liberação da Engenharia') & (self.df_relatorio_lap['Data de Finalização da LAP'] <= self.df_relatorio_lap['Prazo LAP']), 'LAP Entregue em Atraso'] = 'Entregue ON Time'
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['Lead Time da LAP'] != 'LAP Em Fluxo') & (self.df_relatorio_lap['Lead Time da LAP'] != 'Pendente de Liberação da Engenharia') & (self.df_relatorio_lap['Data de Finalização da LAP'] > self.df_relatorio_lap['Prazo LAP']), 'LAP Entregue em Atraso'] = 'Entregue em Atraso'
        self.df_relatorio_lap.loc[((self.df_relatorio_lap['Lead Time da LAP'] == 'LAP Em Fluxo') | (self.df_relatorio_lap['Lead Time da LAP'] == 'Pendente de Liberação da Engenharia')), 'LAP Entregue em Atraso'] = 'LAP Não Entregue'
        self.df_relatorio_lap.loc[(pd.isna(self.df_relatorio_lap['Prazo LAP']), 'LAP Entregue em Atraso')] = 'Prazo Não Definido para Categoria da LAP'

    def definir_atraso_laps_liberadas(self):
        '''Definir o número de dias de atraso das LAPs já liberadas'''
        self.df_relatorio_lap['Número de Dias de LAP Entregue em Atraso'] = None
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['LAP Entregue em Atraso'] == True), 'Número de Dias de LAP Entregue em Atraso'] = self.df_relatorio_lap[(self.df_relatorio_lap['LAP Entregue em Atraso'] == True)]['Data de Finalização da LAP'] - self.df_relatorio_lap[(self.df_relatorio_lap['LAP Entregue em Atraso'] == True)]['Prazo LAP']
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['LAP Entregue em Atraso'] == False), 'Número de Dias de LAP Entregue em Atraso'] = 'Liberada ON Time'
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['LAP Entregue em Atraso'] == 'LAP Não Entregue'), 'Número de Dias de LAP Entregue em Atraso'] = 'LAP Não Entregue'
        self.df_relatorio_lap.loc[(pd.isna(self.df_relatorio_lap['Prazo LAP'])), 'Número de Dias de LAP Entregue em Atraso'] = 'Prazo Não Definido para Categoria da LAP'

    def definir_dias_de_atraso_das_laps_em_fluxo(self):
        self.df_relatorio_lap['Atraso LAPs Pendentes (Dias)'] = None
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['Status'] == 'EM ANDAMENTO') & (self.df_relatorio_lap['LAP no Fluxo e em Atraso'] == True) & (pd.notna(self.df_relatorio_lap['Prazo LAP'])), 'Atraso LAPs Pendentes (Dias)'] = self.hoje - self.df_relatorio_lap[(self.df_relatorio_lap['LAP no Fluxo e em Atraso'] == True) & (pd.notna(self.df_relatorio_lap['Prazo LAP']))]['Prazo LAP']
        self.df_relatorio_lap.loc[(pd.isna(self.df_relatorio_lap['Prazo LAP'])), 'Atraso LAPs Pendentes (Dias)'] = 'Prazo Não Definido para Categoria da LAP'

    def calcula_lt_por_departamento(self):
        def define_lt_depto_independentes(df):
            df_departamentos_independentes = df[(~pd.notna(df['Pai'])) & (df['Depto'] != 'DPCP')]
            
            #Definição de Lead Time
            self.df_relatorio_lap.loc[df_departamentos_independentes.index, 'Lead Time Departamento'] =  df.loc[df_departamentos_independentes.index, 'Data Lib EA'] - df.loc[df_departamentos_independentes.index, 'Data de liberação'] 

        def define_lt_depto_dependentes(df):
            df_dependencia = pd.DataFrame({'Depto': df['Depto'], 'Pai': df['Pai']})
            df_dependencia = df_dependencia.dropna()

            if df_dependencia.empty:
                pass
            else:
                for index, row in df_dependencia.iterrows():
                    self.df_relatorio_lap.loc[index, 'Lead Time Departamento'] = df.loc[index, 'Data Lib EA'] - df.loc[df['Depto'] == row['Pai'], 'Data Lib EA'].iloc[0]

        def define_lt_pcp(df):
            df_deptos = df[(df['Depto'] != 'DPCP')]['Data Lib EA']
            if (not pd.isna(df_deptos).all()) and (pd.notna(df[(df['Depto'] == 'DPCP')]['Data Lib EA'].values[0])):
                ultimo_prazo = df.loc[df['Depto'] != 'DPCP', 'Data Lib EA'].max()
                self.df_relatorio_lap.loc[df[df['Depto'] == 'DPCP'].index, 'Lead Time Departamento'] = self.df_relatorio_lap.loc[df[df['Depto'] == 'DPCP'].index, 'Data Lib EA'] - ultimo_prazo


        self.df_relatorio_lap['Lead Time Departamento'] = None
        group_codigo_lap = self.df_relatorio_lap.groupby('Código')[['Código', 'Data de liberação', 'Depto', 'Data Lib EA', 'Pai']]
        for group, df in tqdm(group_codigo_lap, desc='Lead Time por Departamento Definidos'):
            define_lt_depto_independentes(df)
            define_lt_depto_dependentes(df)
            define_lt_pcp(df)
    
    def define_prazos_departamentos(self):
        ''''Lógica que define os prazos para cada departamento'''
        def analise_motivo(df):
            match df['Motivo'].unique()[0]:
                case 'CORRECAO CRITICA':
                    prazo_lap = self.prazo_correcao_critica
                case 'ADEQUACAO DO PRODUTO':
                    prazo_lap = self.prazo_adequacao_produto
                case 'LIBERACAO DE PRODUTO NOVO':
                    prazo_lap = self.prazo_adequacao_produto
                case 'TABELA DE APLICAÇÃO':
                    prazo_lap = self.prazo_tabela_aplicacao
                case 'HOMOLOGACAO':
                    prazo_lap = self.prazo_homologacao
                case _:
                    prazo_lap = None
            return prazo_lap

        def define_prazo_dpcp(df):
            '''Esta função determina o prazo do DPCP para a LAP. O seu prazo será sempre o prazo da LAP (Pois é sempre o último
            departamento no processo)'''
            df.loc[(self.df_relatorio_lap['Código'] == group) & (self.df_relatorio_lap['Depto'] == 'DPCP'), 'Prazo Departamento'] = df.loc[(df['Código'] == group) & (df['Depto'] == 'DPCP'), 'Prazo LAP']

        def reorganiza_departamento_pai_e_filho(df):
            '''Regra adicionada: Caso uma LAP tenha o departamento DEQF e DCOM, o DEQF terá como Departamento Pai o DCOM, obrigatóriamente, de acordo com a regra de negócio explicada por William
            Machado de Moraes. Caso o DEQF já tenha um departamento Pai, este é reorganizado.
            '''

            departamento_substituido = None
            manipulacoes_realizadas = False

            #As LAPs que geram um LOOP infinito ao reorganizar departamentos pai e filho, de acordo com as regras estabelecidas pelo código abaixo, são poucas excessões e terão seu prazo calculado normalmente, sem considerar regra entre DEQF e DCOM.
            match df['Código'].values[0]:
                case '0415/22':
                    return df, manipulacoes_realizadas, departamento_substituido
                case '0747/21':
                    return df, manipulacoes_realizadas, departamento_substituido
                case '0861/21':
                    return df, manipulacoes_realizadas, departamento_substituido
                
            #Verificar se na LAP contém DEQF e DCOM juntas
            if df['Depto'].str.contains('DEQF').any() and df['Depto'].str.contains('DCOM').any():
                
                if (df.loc[df['Depto'] == 'DCOM', 'Pai'] == 'DEQF').all(): #Caso o DCOM tenha como departamento pai DEQF (caso raro, só aconteceu na LAP 0307/23). Neste caso, a manipulação da coluna Pai não é realizada.
                    return df, manipulacoes_realizadas, departamento_substituido

                #Caso a LAP tenha DEQF e DCOM juntas, o próximo passo é verificar se DCOM realiza a LAP antes do DEQF        
                pai_deqf = df.loc[df['Depto'] == 'DEQF', 'Pai'].values[0]
                if pd.isna(pai_deqf):
                    dcom_em_nivel_abaixo = False #Caso o DEQF não haja departamento Pai, não é necessário descer até o último nível, e não 
                elif pai_deqf == 'DCOM':
                    dcom_em_nivel_abaixo = True
                    manipulacoes_realizadas = False
                elif pai_deqf: #Caso DEQF haja pai, vai descer até o último nível ou até encontrar o DCOM (Retorna True), se não encontrar DCOM retorna False
                    nivel_abaixo = pai_deqf #Cria uma cópia da variável do pai do DEQF, pois esta irá ir descendo até o último nível. A variável pai_deqf pode ser usada posteriormente
                    while True:
                        #Encontrar o pai do pai do DEQF
                        nivel_abaixo = df.loc[df['Depto'] == nivel_abaixo, 'Pai'].values[0]
                        if pd.notna(nivel_abaixo): #Se encontrou departamento em um nível mais abaixo, iteração continua.
                            if nivel_abaixo == 'DCOM':
                                dcom_em_nivel_abaixo = True
                                manipulacoes_realizadas = False
                                departamento_substituido = None
                                break
                            else:
                                continue
                        else: #Caso não haja mais departamentos, a iteração é parada e é informado que DCOM não está no nível abaixo, e que manipulações devem ser realizadas.
                            dcom_em_nivel_abaixo = False
                            break
                            
                #Caso a LAP não tenha o DCOM antes do DEQF, manipulações devem ser realizadas.
                if not dcom_em_nivel_abaixo:
                    
                    #Caso o DEQF não tenha departamento Pai, o DCOM apenas é inserido como Pai do DEQF.
                    if pd.isna(pai_deqf):
                        df.loc[df['Depto'] == 'DEQF', 'Pai'] = 'DCOM'
                        manipulacoes_realizadas = False #Esta variável indica se o departamento substituído terão as datas regredidas a partir do DEQF    
                    else: #Caso o DEQF tenha departamento Pai, manipulações da coluna Pai são realizadas
                        departamento_substituido = df.loc[df['Depto'] == 'DEQF', 'Pai'].values[0]

                        #Substitui o pai do DEQF como DCOM
                        df.loc[df['Depto'] == 'DEQF', 'Pai'] = 'DCOM'
                    
                        manipulacoes_realizadas = True

                        #Verifica se o DCOM possuí departamento Pai. Se não tiver, a regra será: DCOM e departamento substituído terão a mesma data
                        if pd.isna(df.loc[df['Depto'] == 'DCOM', 'Pai'].values[0]):

                            #É definido um rótulo informando que DCOM e o departamento substituído devem ter as mesma data.
                            df.loc[(df['Depto'] == departamento_substituido) | (df['Depto'] == 'DCOM'), 'regra_mudanca_deqf'] = 'mesma_data'
                            
                        else: #Caso DCOM tenha departamento pai, a regra será: Departamento substituíto terão datas regredidas a partir da data já definida por DEQF.
                            df.loc[df['Depto'] == departamento_substituido, 'regra_mudanca_deqf'] = 'inferior_a_deqf'
                            
                else: #Caso a LAP tenha DCOM nates do DEQF, manipulações não são feitas
                    manipulacoes_realizadas = False

            return df, manipulacoes_realizadas, departamento_substituido



        def definir_deptos_impactantes_e_sem_impacto(df, departamento_substituido):
            '''Esta função é uma lógica que define os departamentos que impactam a continuidade dos demais departamentos
            das LAPs'''
            #Primeira parte da função: definir os departamentos com e sem impacto
            departamentos_nao_impactantes = df[(df['Depto'] != 'DPCP') & (~df['Depto'].isin(df['Pai']))][['Depto','Pai']]
            departamentos_com_impacto = df[(df['Depto'] != 'DPCP') & (df['Depto'].isin(df['Pai']))]['Depto']

            if departamento_substituido:
                departamentos_nao_impactantes = departamentos_nao_impactantes[departamentos_nao_impactantes['Depto'] != departamento_substituido]
                departamentos_com_impacto = departamentos_com_impacto._append(pd.Series([departamento_substituido]))
            #Caso não hava nenhum departamento que impacte o processo, a lógica abaixo não será executada pois não seria necessário
            if departamentos_com_impacto.empty:
                lap_sem_deptos_com_impacto = True
                penultimo_depto = None

                return departamentos_nao_impactantes, departamentos_com_impacto, penultimo_depto, lap_sem_deptos_com_impacto
            
            else:
                #Achar o penúltimo departamento no processo, sendo o último o DPCP
                series_aux = df.loc[df['Depto'].isin(departamentos_com_impacto)]
                #pai = series_aux[pd.notna(df['Pai'])]['Pai']
                pai = series_aux.loc[(pd.notna(df['Pai'])), 'Pai']
                penultimo_depto = series_aux.loc[~series_aux['Depto'].isin(pai)][['Depto', 'Pai']]

                #Caso um departamento sem impacto em outros departamentos tenha como o seu 'Pai' o penúltimo departamento do 'Caminho Crítico' (Departamentos que impactam outros departamentos), este é categorizado como departamento com impacto. 
                if departamentos_nao_impactantes['Pai'].isin(penultimo_depto['Depto']).any():

                    #Atualiza a variável de departamentos com impacto
                    departamentos_com_impacto = pd.concat([departamentos_com_impacto, departamentos_nao_impactantes[departamentos_nao_impactantes['Pai'].isin(penultimo_depto['Depto'])]['Depto']])
                    
                    #Atualiza a variável de departamentos sem impacto
                    departamentos_nao_impactantes = departamentos_nao_impactantes[~departamentos_nao_impactantes['Pai'].isin(penultimo_depto['Depto'])]

                    #Transforma o 'departamentos_nao_impactantes' novamente em uma series, para o funcionamento correto do código
                    departamentos_nao_impactantes = departamentos_nao_impactantes['Depto']
                
                    #Atualizar a variável de penultimo_depto
                    series_aux = df.loc[df['Depto'].isin(departamentos_com_impacto)]
                    #pai = series_aux[pd.notna(df['Pai'])]['Pai']
                    pai = series_aux.loc[(pd.notna(df['Pai'])), 'Pai']
                    penultimo_depto = series_aux.loc[~series_aux['Depto'].isin(pai)][['Depto', 'Pai']]
                
                #Caso não satisfeito a condição acima, não é realizado alterações
                else:
                    #Transforma o 'departamentos_nao_impactantes' novamente em uma series, para o funcionamento do código abaixo
                    departamentos_nao_impactantes = departamentos_nao_impactantes['Depto']
                
                lap_sem_deptos_com_impacto = False


            return departamentos_nao_impactantes, departamentos_com_impacto, penultimo_depto, lap_sem_deptos_com_impacto

        def define_prazos_se_lap_tem_somente_departamentos_que_nao_impactam(df):
            '''Caso não haja nenhum departamento com impacto em outros departamentos, o número de dias para cada departamento é dividido
            igualmente entre o DPCP e os demais departamentis'''

            num_dias_cada_depto = math.floor(prazo_lap/2)
            df.loc[df['Depto'] != 'DPCP', 'Prazo Departamento'] = (df.loc[df['Depto'] == 'DPCP', 'Prazo Departamento'] - datetime.timedelta(num_dias_cada_depto)).values[0]
            
            return df

        def define_prazos_se_lap_tem_departamentos_que_impactam(df, departamentos_com_impacto, departamentos_nao_impactantes):

            num_dias_cada_depto = math.floor(prazo_lap/(len(departamentos_com_impacto) + 1))

            #Achar o penúltimo departamento
            series_aux = df.loc[df['Depto'].isin(departamentos_com_impacto)]
            #pai = series_aux[pd.notna(df['Pai'])]['Pai']
            pai = series_aux.loc[(pd.notna(df['Pai'])), 'Pai']
            penultimo_depto = series_aux.loc[~series_aux['Depto'].isin(pai)][['Depto', 'Pai']]
            #df.loc[df['Depto'].isin(penultimo_depto['Depto']), 'Prazo Departamento'] = calendario_tke.soma_dias_uteis(df.loc[df['Depto'] == 'DPCP', 'Prazo Departamento'], - num_dias_cada_depto)
            df.loc[df['Depto'].isin(penultimo_depto['Depto']), 'Prazo Departamento'] = (df.loc[df['Depto'] == 'DPCP', 'Prazo Departamento'] - datetime.timedelta(num_dias_cada_depto)).values[0]

            departamento_abaixo = df.loc[df['Depto'].isin(departamentos_com_impacto) & (pd.notna(df['Prazo Departamento']))][['Depto', 'Pai', 'Prazo Departamento']]

            #Definir prazos de PCP para baixo
            while True:
                #df.loc[df['Depto'].isin(departamento_abaixo['Pai']), 'Prazo Departamento'] = calendario_tke.soma_dias_uteis(df.loc[df['Depto'].isin(departamento_abaixo['Depto']), 'Prazo Departamento'], -num_dias_cada_depto)
                df.loc[df['Depto'].isin(departamento_abaixo['Pai']), 'Prazo Departamento'] = (df.loc[df['Depto'].isin(departamento_abaixo['Depto']), 'Prazo Departamento'] - datetime.timedelta(num_dias_cada_depto)).values[0]
                df['Prazo Departamento'] = pd.to_datetime(df['Prazo Departamento'])        
                departamento_abaixo = df.loc[df['Depto'].isin(departamentos_com_impacto) & ((df['Prazo Departamento'] == df['Prazo Departamento'].min()))][['Depto', 'Pai', 'Prazo Departamento']]
                if df.loc[df['Depto'].isin(departamentos_com_impacto)]['Prazo Departamento'].notnull().all():
                    break

            #Definir prazo dos departamentos que não impactam

            df.loc[df['Depto'].isin(departamentos_nao_impactantes), 'Prazo Departamento'] = df.loc[df['Depto'].isin(departamentos_com_impacto), 'Prazo Departamento'].max()

            ultimo_departamento = departamento_abaixo

            return df, ultimo_departamento

        def define_prazos_se_lap_teve_manipulacoes_realizadas(df, departamento_substituido):
            '''Esta função calcula os prazos somente das LAPs que tiveram manipulações realizadas: caso onde a LAP que tenha DEQF e DCOM juntas, e coluna 'regra_mudanca_deqf'
            da LAP ter algum valor'''
            
            num_dias_cada_depto = math.floor(prazo_lap/(len(departamentos_com_impacto) + 1))

            #Caso onde o DCOM e o departamento substituído foram categorizados para terem a mesma data
            if df['regra_mudanca_deqf'].str.contains('mesma_data').any():
                
                #Achar o penúltimo departamento
                series_aux = df.loc[df['Depto'].isin(departamentos_com_impacto)]
                #pai = series_aux[pd.notna(df['Pai'])]['Pai']
                pai = series_aux.loc[(pd.notna(df['Pai'])), 'Pai']
                penultimo_depto = series_aux.loc[~series_aux['Depto'].isin(pai)][['Depto', 'Pai']]
                #df.loc[df['Depto'].isin(penultimo_depto['Depto']), 'Prazo Departamento'] = calendario_tke.soma_dias_uteis(df.loc[df['Depto'] == 'DPCP', 'Prazo Departamento'], - num_dias_cada_depto)
                df.loc[df['Depto'].isin(penultimo_depto['Depto']), 'Prazo Departamento'] = (df.loc[df['Depto'] == 'DPCP', 'Prazo Departamento'] - datetime.timedelta(num_dias_cada_depto)).values[0]

                departamento_abaixo = df.loc[df['Depto'].isin(departamentos_com_impacto) & (pd.notna(df['Prazo Departamento']))][['Depto', 'Pai', 'Prazo Departamento']]

                #Definir prazos de PCP para baixo
                while True:
                    #df.loc[df['Depto'].isin(departamento_abaixo['Pai']), 'Prazo Departamento'] = calendario_tke.soma_dias_uteis(df.loc[df['Depto'].isin(departamento_abaixo['Depto']), 'Prazo Departamento'], -num_dias_cada_depto)
                    df.loc[df['Depto'].isin(departamento_abaixo['Pai']), 'Prazo Departamento'] = (df.loc[df['Depto'].isin(departamento_abaixo['Depto']), 'Prazo Departamento'] - datetime.timedelta(num_dias_cada_depto)).values[0]
                    df['Prazo Departamento'] = pd.to_datetime(df['Prazo Departamento'])        
                    departamento_abaixo = df.loc[df['Depto'].isin(departamentos_com_impacto) & ((df['Prazo Departamento'] == df['Prazo Departamento'].min()))][['Depto', 'Pai', 'Prazo Departamento']]
                    
                    if pd.notna(df.loc[df['Depto'] == 'DCOM', 'Prazo Departamento'].values[0]):
                        df.loc[df['Depto'] == departamento_substituido, 'Prazo Departamento'] = df.loc[df['Depto'] == 'DCOM', 'Prazo Departamento'].values[0]
                
                    if df.loc[df['Depto'].isin(departamentos_com_impacto)]['Prazo Departamento'].notnull().all():
                        break
                    
                df.loc[df['Depto'].isin(departamentos_nao_impactantes), 'Prazo Departamento'] = df.loc[df['Depto'].isin(departamentos_com_impacto), 'Prazo Departamento'].max()
            
            elif df['regra_mudanca_deqf'].str.contains('inferior_a_deqf').any():
                #Achar o penúltimo departamento
                series_aux = df.loc[df['Depto'].isin(departamentos_com_impacto)]
                #pai = series_aux[pd.notna(df['Pai'])]['Pai']
                pai = series_aux.loc[(pd.notna(df['Pai'])), 'Pai']
                penultimo_depto = series_aux.loc[~series_aux['Depto'].isin(pai)][['Depto', 'Pai']]
                #df.loc[df['Depto'].isin(penultimo_depto['Depto']), 'Prazo Departamento'] = calendario_tke.soma_dias_uteis(df.loc[df['Depto'] == 'DPCP', 'Prazo Departamento'], - num_dias_cada_depto)
                df.loc[df['Depto'].isin(penultimo_depto['Depto']), 'Prazo Departamento'] = (df.loc[df['Depto'] == 'DPCP', 'Prazo Departamento'] - datetime.timedelta(num_dias_cada_depto)).values[0]

                departamento_abaixo = df.loc[df['Depto'].isin(departamentos_com_impacto) & (pd.notna(df['Prazo Departamento']))][['Depto', 'Pai', 'Prazo Departamento']]

                depto_inferior_a_deqf = df.loc[df['regra_mudanca_deqf'] == 'inferior_a_deqf', 'Depto'].values[0]

                #Definir prazos dos departamentos, exceto o departamento substituído
                while True:

                    #Caso o departamento abaixo contenha o departamento inferior ao DEQF, este é removido do Loop, pois o mesmo terá sua data regredida posteriormente pelo DEQF
                    if departamento_abaixo['Depto'].str.contains(depto_inferior_a_deqf).any():
                        departamento_abaixo = departamento_abaixo.loc[departamento_abaixo['Depto'] != depto_inferior_a_deqf]
                    
                    
                    df.loc[df['Depto'].isin(departamento_abaixo['Pai']), 'Prazo Departamento'] = (df.loc[df['Depto'].isin(departamento_abaixo['Depto']), 'Prazo Departamento'] - datetime.timedelta(num_dias_cada_depto)).values[0]
                    df['Prazo Departamento'] = pd.to_datetime(df['Prazo Departamento'])        
                    departamento_abaixo = df.loc[df['Depto'].isin(departamentos_com_impacto) & ((df['Prazo Departamento'] == df['Prazo Departamento'].min()))][['Depto', 'Pai', 'Prazo Departamento']]
                    
                    #Caso todos os departamentos estejam corretamente preenchidos, com excessão do departamento inferior a deqf, a iteração para
                    if pd.isna(departamento_abaixo['Pai']).all():
                        break
                
                #Definir prazos dos departamentos, a partir do DEQF
                df.loc[df['Depto'] == departamento_substituido, 'Prazo Departamento'] = (df.loc[df['Depto'] == 'DEQF', 'Prazo Departamento'] - datetime.timedelta(num_dias_cada_depto)).values[0]
                departamento_abaixo = df.loc[df['Depto'] == departamento_substituido][['Depto', 'Pai']]
                while True:
                    if pd.isna(departamento_abaixo['Pai'].values[0]):
                        break
                    df.loc[df['Depto'] == departamento_abaixo['Pai'].values[0], 'Prazo Departamento'] = (df.loc[df['Depto'] == departamento_abaixo['Depto'].values[0], 'Prazo Departamento'] - datetime.timedelta(num_dias_cada_depto)).values[0]
                    departamento_abaixo = df.loc[df['Depto'] == departamento_abaixo['Pai'].values[0]]

        def definir_status_departamentos_se_lap_tem_somente_departamentos_que_nao_impactam(df):
            '''Esta função tem como objetivo definir o status dos departamentos caso a lap tenha
            somente departamentos que não impactam uns aos outros. Os status são: 'Liberada', 'No fluxo' e 'Aguardando
            Liberação dos Demais Departamentos' '''

            df.loc[pd.notna(df['Data Lib EA']), 'Status Departamento'] = 'Liberada'
            df.loc[(~pd.notna(df['Data Lib EA'])), 'Status Departamento'] = 'Aguardando Liberação dos Demais Departamentos'
            self.df_relatorio_lap.loc[df.index, 'Status Departamento'] = df['Status Departamento']
            
            return df

        def definir_status_departamentos_se_lap_tem_departamentos_que_impactam(df, ultimo_departamento):
            '''Esta função tem como objetivo definir o status dos departamentos caso a lap tenha
            departamentos que impactam outro. Os status são: 'Liberada', 'No fluxo' e 'Aguardando
            Liberação dos Demais Departamentos' '''

            #Definir departamentos que já liberaram sua LAP
            departamentos_que_ja_liberaram = df.loc[pd.notna(df['Data Lib EA']), 'Depto']

            #Se nenhum departamento liberou, o status de 'Realizando Liberação' é alocada para o primeiro departamento
            if departamentos_que_ja_liberaram.empty:
                df.loc[df['Depto'].isin(ultimo_departamento['Depto']), 'Status Departamento'] = 'Realizando Liberação'    
                df.loc[((pd.isna(df['Pai'])) & (df['Depto'] != 'DPCP')), 'Status Departamento'] = 'Realizando Liberação'

                df.loc[((pd.notna(df['Pai'])) | (df['Depto'] == 'DPCP')), 'Status Departamento'] = 'Aguardando Liberação dos Demais Departamentos'
            else:
                #Definir status 'Liberada'
                df.loc[df['Depto'].isin(departamentos_que_ja_liberaram), 'Status Departamento'] = 'Liberou'

                #Definir status 'No fluxo' (Departamentos dependentes de outros que já liberaram)
                df.loc[(pd.isna(df['Data Lib EA'])) & (df['Pai'].isin(departamentos_que_ja_liberaram)), 'Status Departamento'] = 'Realizando Liberação'    

                #Definir status 'No fluxo' (Departamentos Independentes)
                df.loc[((pd.isna(df['Pai'])) & (df['Depto'] != 'DPCP')) & (pd.isna(df['Status Departamento'])), 'Status Departamento'] = 'Realizando Liberação'

                #Definir no status 'Aguardando Liberação dos Demais Departamentos'
                df.loc[(pd.isna(df['Data Lib EA'])) & (~df['Pai'].isin(departamentos_que_ja_liberaram)) & (pd.isna(df['Status Departamento'])), 'Status Departamento'] = 'Aguardando Liberação dos Demais Departamentos'

            self.df_relatorio_lap.loc[df.index, 'Status Departamento'] = df['Status Departamento']
            return df

        def transferir_ao_df_relatorio(df):
            '''Função que transfere prazos do dataframe resultante da iteração ao dataframe relatório'''

            self.df_relatorio_lap.loc[df.index, 'Prazo Departamento'] = df['Prazo Departamento']

        group_codigo = self.df_relatorio_lap.loc[pd.notna(self.df_relatorio_lap['Data de liberação'])].groupby('Código')

        for group, df in tqdm(group_codigo, desc='LAPs com Departamentos com Prazos Definidos'):
            prazo_lap = analise_motivo(df)

            if not prazo_lap:
                #Caso não encontre algum motivo com prazo para LAP, os prazos para cada departamento não é calculado 
                continue
            else:
                define_prazo_dpcp(df)
                df, manipulacoes_realizadas, departamento_substituido = reorganiza_departamento_pai_e_filho(df)


                if not manipulacoes_realizadas: #Caso o DEQF não seja considerado como ponto de partida (definido na função "reorganiza_departamento_pai_e_filho"), a função é executada da maneira normal.

                    departamentos_nao_impactantes, departamentos_com_impacto, penultimo_depto, lap_sem_deptos_com_impacto = definir_deptos_impactantes_e_sem_impacto(df, departamento_substituido)
                    
                    if lap_sem_deptos_com_impacto:
                        df = define_prazos_se_lap_tem_somente_departamentos_que_nao_impactam(df)
                    else:
                        df, ultimo_departamento = define_prazos_se_lap_tem_departamentos_que_impactam(df, departamentos_com_impacto, departamentos_nao_impactantes)

                elif manipulacoes_realizadas: #Caso o DEQF seja considerado como ponto de partida (definido na função "reorganiza_departamento_pai_e_filho"), a lógica abaixo é executada.
                    departamentos_nao_impactantes, departamentos_com_impacto, penultimo_depto, lap_sem_deptos_com_impacto = definir_deptos_impactantes_e_sem_impacto(df, departamento_substituido)
                    define_prazos_se_lap_teve_manipulacoes_realizadas(df, departamento_substituido)


                    #df, ultimo_departamento = define_prazos_se_lap_tem_departamentos_que_impactam(df, departamentos_com_impacto, departamentos_nao_impactantes)
                
                
                transferir_ao_df_relatorio(df)    

                if (df['Status'].unique()[0] == 'EM ANDAMENTO'):
                    if lap_sem_deptos_com_impacto:
                        df = definir_status_departamentos_se_lap_tem_somente_departamentos_que_nao_impactam(df)
                    else:
                        df = definir_status_departamentos_se_lap_tem_departamentos_que_impactam(df, ultimo_departamento)

    def define_departamentos_atrasados(self):
        '''Define se os departamentos entregaram ON Time ou em atraso, de acordo com os prazos definidos pelo próprio script.'''
        self.df_relatorio_lap['Atraso Departamento'] = None
        self.df_relatorio_lap.loc[(pd.notna(self.df_relatorio_lap['Prazo Departamento'])) & (self.df_relatorio_lap['Status'] == 'LIBERADA') & (self.df_relatorio_lap['Data Lib EA'] > self.df_relatorio_lap['Prazo Departamento']), 'Atraso Departamento'] = 'Entregue em Atraso'
        self.df_relatorio_lap.loc[(pd.notna(self.df_relatorio_lap['Prazo Departamento'])) & (self.df_relatorio_lap['Status'] == 'LIBERADA') & (self.df_relatorio_lap['Data Lib EA'] <= self.df_relatorio_lap['Prazo Departamento']), 'Atraso Departamento'] = 'Entregue ON TIME'

    def atualiza_planilha_justificativa_atrasos(self):
        '''Atualiza a planilha onde o Analista do DPCP deve justificar as LAPs que estão ou foram entregues em atraso'''
        self.df_justificativa_atrasos = pd.read_excel(self.caminho_lap + "\\" + self.arquivo_justificativa_atrasos)
        
        #Cria Dataframe de todas as LAPs atrasadas, a partir do dia 01 de novembro de 2024
        laps_atrasadas = self.df_relatorio_lap[((self.df_relatorio_lap['LAP Entregue em Atraso'] == 'Entregue em Atraso') | (self.df_relatorio_lap['LAP no Fluxo e em Atraso'] == 'Em Atraso')) & (self.df_relatorio_lap['Dt. Abertura'] >= datetime.datetime(2024, 11, 1))][['Código','Data de liberação','Data de Finalização da LAP','Motivo']]
        laps_atrasadas = laps_atrasadas.drop_duplicates(['Código'])
        laps_atrasadas.loc[:, 'Justificativa do Atraso'] = None
        laps_atrasadas.loc[:, 'Observação'] = None

        #Atualiza o Dataframe de justificativa de atrasos, fazendo com que novas LAPs atrasadas sejam incluídas neste dataframe
        novas_linhas = laps_atrasadas[~laps_atrasadas['Código'].isin(self.df_justificativa_atrasos['Código'])]
        self.df_justificativa_atrasos = pd.concat([self.df_justificativa_atrasos, novas_linhas], ignore_index=False)

        #Atualiza a data de finalização da LAP do dataframe de justificativa atrasos
        self.df_justificativa_atrasos = self.df_justificativa_atrasos.merge(self.df_relatorio_lap[['Código', 'Data de Finalização da LAP']].drop_duplicates(['Código']), how='inner', on='Código')
        self.df_justificativa_atrasos = self.df_justificativa_atrasos.drop(columns=['Data de Finalização da LAP_x'])
        self.df_justificativa_atrasos = self.df_justificativa_atrasos.rename(columns={'Data de Finalização da LAP_y': 'Data de Finalização da LAP'})
        self.df_justificativa_atrasos = self.df_justificativa_atrasos[['Código', 'Data de liberação','Data de Finalização da LAP' ,'Motivo', 'Justificativa do Atraso', 'Observação']]
        self.df_justificativa_atrasos['Data de Finalização da LAP'] = self.df_justificativa_atrasos['Data de Finalização da LAP'].fillna('LAP em fluxo e atrasada')
        self.df_justificativa_atrasos.to_excel(self.caminho_lap + "\\" + self.arquivo_justificativa_atrasos, index=False)

    def gravar_relatorio_analisado(self):
        '''Grava o relatório final, após o script realizar as análises necessárias para o BI e o posterior envio de e-mail aos departamentos.'''
        self.df_relatorio_lap.to_excel(self.caminho_lap + "\\" + self.arquivo_relatorio_analisado, index=False)