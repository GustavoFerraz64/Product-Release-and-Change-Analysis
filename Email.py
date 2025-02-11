import win32com.client
import pandas as pd
import datetime

class EnvioEmail:
    def __init__(self, df_relatorio_lap, df_justificativa_atrasos, conexao_outlook):
        self.df_relatorio_lap = df_relatorio_lap
        self.df_justificativa_atrasos = df_justificativa_atrasos
        self.path_destinatarios = r'\\srvfile01.tsur.local\DADOS_PBI\Compartilhado_BI\DPCP\9. DEPM\Bases de Dados\E-mails'
        self.file_destinatarios = r'LaP Report.xlsx'
        self.outlook = conexao_outlook
        self.lista_todos_deptos_email = []
        self.lista_departamentos_atrasados = []
        self.lista_departamentos_somente_atrasados = []
    
    def desconsidera_laps(self):
        '''Desconsiderar algumas LAPs que não serão enviadas e-mail, conforme combinado com William de Moraes'''
        self.df_relatorio_lap = self.df_relatorio_lap.loc[self.df_relatorio_lap['Código'] != '0254/23']

    def trata_prazo_departamentos(self):
        '''Converte a coluna Prazo Departamento para Datetime, para operações posteriores no código.'''
        self.df_relatorio_lap['Prazo Departamento'] = pd.to_datetime(self.df_relatorio_lap['Prazo Departamento'])
    
    def envia_email_pcp(self):
        '''Esta função envia e-mail aos responsável pelo preenchimento de justificativa de LAPs atrasadas
        '''
        df_justificativa_atrasos_email = self.df_justificativa_atrasos.loc[pd.isna(self.df_justificativa_atrasos['Justificativa do Atraso']), ['Código', 'Data de liberação', 'Data de Finalização da LAP', 'Motivo']]
        
        if not df_justificativa_atrasos_email.empty:

            email = self.outlook.CreateItem(0)
            email.Subject = 'Preenchimento de Justificativa de LAPs Atrasadas'

            self.df_tabela_email = df_justificativa_atrasos_email.to_html(index=False)
            
            #Enviar o e-mail para o analista responsável
            email.To = 'william.moraes@tkelevator.com'
            #email.CC = 'nicolas.ferretti@tkelevator.com'

            #Cria o corpo do email
            lista_email = ['---Mensagem Automática---</p>']
            lista_email.append('<p>Olá!</p>')
            lista_email.append(f'<p>Existem LAPs em atraso pendentes de justificativa<p>')
            lista_email.append(f'<p>Segue a tabela abaixo destas LAPs:<p>')
            lista_email.append(self.df_tabela_email)
            lista_email.append('<br>')
            lista_email.append(f'<p>Favor justificar o atraso destas LAPs em: \\srvfile01.tsur.local\DADOS_PBI\Compartilhado_BI\DPCP\25. LAPs\Justificativa de LAPs em Atraso.xlsx<p>')
            lista_email.append('<br>')
            lista_string = '\n'.join(lista_email)

            email.HTMLBody = lista_string

            #Envia o e-mail
            email.Send()
        
        else:
            print('Não há LAPs atrasadas pendentes de justificativa, e-mail ao DPCP não será enviado.')

    def ler_emails_destinatarios(self):
        '''Lê o arquivo Excel que contém os destinatários de cada departamento'''
        self.lista_destinatarios = pd.read_excel(self.path_destinatarios + "\\" + self.file_destinatarios)
        #self.lista_destinatarios['Departamento'] = self.lista_destinatarios['Departamento'].ffill()

    def cria_rotulo_laps(self):
        '''Define quais departamentos estão com suas LAPs atrasadas, ou que irão vencer daqui a uma semana.'''
        #Cria rótulo para indicar departamentos no fluxo e que estão ON time e D+7
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['Status Departamento'] == 'Realizando Liberação') 
                            & (self.df_relatorio_lap['Prazo Departamento'] <= (datetime.datetime.today() + datetime.timedelta(7))) 
                            & (self.df_relatorio_lap['Prazo Departamento'] >= datetime.datetime.today()), 'ON Time e Vence D+7'] = True

        #Cria rótulo para indicar departamentos no fluxo e que estão em atraso
        self.df_relatorio_lap.loc[(self.df_relatorio_lap['Status Departamento'] == 'Realizando Liberação') 
                            & (self.df_relatorio_lap['Prazo Departamento'] < datetime.datetime.today()),
                            'Departamento em Atraso'] = True

        self.lista_todos_deptos_email = set(self.df_relatorio_lap.loc[(self.df_relatorio_lap['ON Time e Vence D+7'] == True) | (self.df_relatorio_lap['Departamento em Atraso'] == True), 'Depto'].unique().tolist())

        self.lista_departamentos_que_possui_lap_ontime_d7 = set(self.df_relatorio_lap.loc[self.df_relatorio_lap['ON Time e Vence D+7'] == True, 'Depto'].unique().tolist())

        self.lista_departamentos_que_possui_lap_atrasada = set(self.df_relatorio_lap.loc[self.df_relatorio_lap['Departamento em Atraso'] == True, 'Depto'].unique().tolist())

        self.lista_departamentos_que_somente_possui_lap_ontime_d7 = self.lista_departamentos_que_possui_lap_ontime_d7 - self.lista_departamentos_que_possui_lap_atrasada

        self.lista_departamentos_que_possuem_somente_laps_atrasadas = self.lista_departamentos_que_possui_lap_atrasada - self.lista_departamentos_que_possui_lap_ontime_d7 

    def envia_email_demais_departamentos(self):
        '''Função que envia e-mail aos departamentos informanto as LAPs que irão vencer ou que estão em atraso, de acordo com os prazos estabelecidos
        pelo próprio script.'''

        def envia_email_deptos_somente_laps_atrasadas(departamento):
            df_email = self.df_relatorio_lap.loc[(self.df_relatorio_lap['Departamento em Atraso'] == True) & (self.df_relatorio_lap['Depto'] == departamento), ['Código', 'Item', 'Prazo Departamento']]
            df_email = df_email.rename(columns={'Prazo Departamento': f'Prazo de liberação {departamento}'})

            #Transforme colunas datetime em string, para um melhor visual do e-mail
            df_email[f'Prazo de liberação {departamento}'] = df_email[f'Prazo de liberação {departamento}'].dt.strftime("%d/%m/%Y")
            df_email = df_email.style.applymap(
                lambda x: 'background-color: yellow' if df_email['Código'].isin([x]).any() else '',
                subset=['Código']
            ).hide(axis='index').set_table_attributes('border="1" cellspacing="0" cellpadding="5"').to_html(escape=False)

            email = self.outlook.CreateItem(0)
            email.Subject = f"LAPs Em Atraso - {departamento}"

            email_to = self.lista_destinatarios[self.lista_destinatarios['Departamento'] == departamento]['E-mail']
            email_to = email_to.to_list()
            #email_to = 'gustavo.ferraz@tkelevator.com'
            #email_to = email_to.to_list()
            email.To = ';'.join(email_to)
            email.CC = 'gustavo.ferraz@tkelevator.com;william.moraes@tkelevator.com'

            #Cria o corpo do email
            lista_email = []
            lista_email.append('<p>Mensagem Automática - Departamento de Planejamento e Controle da Produção,</p>')
            lista_email.append('<p>Prezados(as),<p>')
            lista_email.append(f'<p>Existem LAPs que estão em atraso de acordo com o prazos estabelecidos pela auditoria. Segue a tabela de relação abaixo:</p>')
            lista_email.append(df_email)
            lista_email.append('<p>')
            lista_email.append('Por gentileza, dar atenção a liberação destas LAPs')
            lista_string = '\n'.join(lista_email)

            email.HTMLBody = lista_string

            #Envia o e-mail
            email.Send()

        def envia_email_deptos_somente_laps_ontime(departamento):
            df_email = self.df_relatorio_lap.loc[(self.df_relatorio_lap['ON Time e Vence D+7'] == True) & (self.df_relatorio_lap['Depto'] == departamento), ['Código', 'Item', 'Prazo Departamento']]
            df_email = df_email.rename(columns={'Prazo Departamento': f'Prazo de liberação {departamento}'})

            #Transforme colunas datetime em string, para um melhor visual do e-mail
            df_email[f'Prazo de liberação {departamento}'] = df_email[f'Prazo de liberação {departamento}'].dt.strftime("%d/%m/%Y")

            df_email = df_email.style.applymap(
                lambda x: 'background-color: yellow' if df_email['Código'].isin([x]).any() else '',
                subset=['Código']
            ).hide(axis='index').set_table_attributes('border="1" cellspacing="0" cellpadding="5"').to_html(escape=False)

            email = self.outlook.CreateItem(0)
            email.Subject = f"LAPs Pendentes de Liberação - {departamento}"
            email_to = self.lista_destinatarios[self.lista_destinatarios['Departamento'] == departamento]['E-mail']
            email_to = email_to.to_list()
            #email.To = 'gustavo.ferraz@tkelevator.com'
            email.To = ';'.join(email_to)
            email.CC = 'gustavo.ferraz@tkelevator.com;william.moraes@tkelevator.com'

            #Cria o corpo do email
            lista_email = []
            lista_email.append('<p>Mensagem Automática - Departamento de Planejamento e Controle da Produção,</p>')
            lista_email.append('<p>Prezados(as),<p>')
            lista_email.append(f'<p>Existem LAPs com prazo de vencimento dentro de uma semana, de acordo com os prazos definidos pela auditoria. Segue a tabela de relação abaixo:</p>')
            lista_email.append(df_email)
            lista_email.append('<p>')
            lista_email.append('Por gentileza, realizar a liberação destas LAPs o quanto antes.')
            lista_string = '\n'.join(lista_email)

            email.HTMLBody = lista_string

            #Envia o e-mail
            email.Send()

        def envia_email_deptos_ontime_atrasados(departamento):
            df_email_ontime = self.df_relatorio_lap.loc[(self.df_relatorio_lap['ON Time e Vence D+7'] == True) & (self.df_relatorio_lap['Depto'] == departamento), ['Código', 'Item', 'Prazo Departamento']]
            df_email_ontime = df_email_ontime.rename(columns={'Prazo Departamento': f'Prazo de liberação {departamento}'})

            df_email_atraso = self.df_relatorio_lap.loc[(self.df_relatorio_lap['Departamento em Atraso'] == True) & (self.df_relatorio_lap['Depto'] == departamento), ['Código', 'Item', 'Prazo Departamento']]
            df_email_atraso = df_email_atraso.rename(columns={'Prazo Departamento': f'Prazo de liberação {departamento}'})

            #Transforme colunas datetime em string, para um melhor visual do e-mail
            df_email_ontime[f'Prazo de liberação {departamento}'] = df_email_ontime[f'Prazo de liberação {departamento}'].dt.strftime("%d/%m/%Y")

            df_email_atraso[f'Prazo de liberação {departamento}'] = df_email_atraso[f'Prazo de liberação {departamento}'].dt.strftime("%d/%m/%Y")

            df_email_ontime = df_email_ontime.style.applymap(
                lambda x: 'background-color: yellow' if df_email_ontime['Código'].isin([x]).any() else '',
                subset=['Código']
            ).hide(axis='index').set_table_attributes('border="1" cellspacing="0" cellpadding="5"').to_html(index=False, escape=False)

            df_email_atraso = df_email_atraso.style.applymap(
                lambda x: 'background-color: yellow' if df_email_atraso['Código'].isin([x]).any() else '',
                subset=['Código']
            ).hide(axis='index').set_table_attributes('border="1" cellspacing="0" cellpadding="5"').to_html(index=False, escape=False)

            email = self.outlook.CreateItem(0)
            email.Subject = f"LAPs Em Atraso - {departamento}"

            email_to = self.lista_destinatarios[self.lista_destinatarios['Departamento'] == departamento]['E-mail']
            email_to = email_to.to_list()
            #email.To = 'gustavo.ferraz@tkelevator.com'
            email.To = ';'.join(email_to)
            email.CC = 'gustavo.ferraz@tkelevator.com;william.moraes@tkelevator.com'

            #Cria o corpo do email
            lista_email = []
            lista_email.append('<p>Mensagem Automática - Departamento de Planejamento e Controle da Produção,</p>')
            lista_email.append('<p>Prezados(as),<p>')
            lista_email.append(f'<p>Existem LAPs que estão em atraso de acordo com os prazos estabelecidos pela auditoria. Segue a tabela de relação abaixo:</p>')
            lista_email.append(df_email_atraso)
            lista_email.append('<p>')
            lista_email.append('<br>')
            lista_email.append(f'<p>Segue outra tabela de relação abaixo com LAPs que irão vencer dentro de uma semana:</p>')
            lista_email.append(df_email_ontime)
            lista_email.append('<br>')
            lista_email.append('Por gentileza, realizar a liberação destas LAPs o quanto antes.')
            lista_string = '\n'.join(lista_email)

            email.HTMLBody = lista_string

            #Envia o e-mail
            email.Send()

        #Lógica de envio dos e-mails
        if self.lista_todos_deptos_email:
            for departamento in self.lista_todos_deptos_email:
                #Primeiro tipo de e-mail: para departamentos que possui somente LAPs em atraso
                if departamento in self.lista_departamentos_que_possuem_somente_laps_atrasadas:
                    print('Eviando e-mail para:', departamento)
                    envia_email_deptos_somente_laps_atrasadas(departamento)
                elif departamento in self.lista_departamentos_que_somente_possui_lap_ontime_d7:
                    print('Eviando e-mail para:', departamento)
                    envia_email_deptos_somente_laps_ontime(departamento)
                else: 
                    print('Eviando e-mail para:', departamento)
                    envia_email_deptos_ontime_atrasados(departamento)