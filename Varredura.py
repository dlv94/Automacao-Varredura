import os
import time
import getpass
import winsound
import pandas as pd
from datetime import datetime,timedelta
from selenium import webdriver
import selenium.webdriver.chrome.options
from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC

## cd C:\Program Files (x86)\Google\Chrome\Application
## chrome.exe --remote-debugging-port=9222 --user-data-dir=C:\Py\ChromeProfile
## https://chromedriver.chromium.org/downloads

fechamentos_rf_atlas = {'FALTA DE ENERGIA NAS RESIDÊNCIAS': '(1113)Sinal Reestabelecido - Não Acionado Gerador',
                        'ATIVO EQUALIZAÇÃO': '(1123)Reequalizacao Do Amplificador - Emergência',
                        'INGRESSO DE RUÍDO - NORMALIZADO DURANTE IDENTIFICAÇÃO': '(1122)Limpeza De Ruído Corretivo',
                        'QUEDA DE ENERGIA - FORNECIMENTO REESTABELECIDO': '(1113)Sinal Reestabelecido - Não Acionado Gerador',
                        'INGRESSO DE RUÍDO - NORMALIZADO SEM INTERVENÇÃO': '(1122)Limpeza De Ruído Corretivo',
                        'QUEDA DE ENERGIA - JANELA DE SINCRONISMO': '(1113)Sinal Reestabelecido - Não Acionado Gerador',
                        'ADEQUAÇÃO DE REDE COAXIAL': '(2126)Adequações, Construções, Reespinamentos De Rede',
                        'ATIVO MODULO COM DEFEITO': '(1124)Amplificador De Rede Com Defeito - Emergência',
                        'REFEITA CONEXÃO': '(1131)Conexao De Rede Externa Danificada - Emergência',
                        'INGRESSO DE RUÍDO - LOGRADOURO FILTRADO': '(1122)Limpeza De Ruído Corretivo',
                        'QUEDA DE ENERGIA - ACIONADO GERADOR': '(1111)Falta Energia - Acionado Gerador/Aliment Alternativa',
                        'ERRO DE ABERTURA': '(2117)Resolvido Pelo Noc - Rede',
                        'ATIVO COM MAU CONTATO': '(2111)Mau Contato Em Equipamentos De Rede Externa',
                        'PROBLEMA INTERNO - CLIENTE': 'Notificado Via Newmonitor',
                        'DISJUNTOR DESARMADO': '(1115)Disjuntor De Fonte Desarmado',
                        'ATIVO QUEIMADO': '(1124)Amplificador De Rede Com Defeito - Emergência',
                        'CABO COAXIAL DANIFICADO': '(1130)Cabo Coaxial De Rede Externa Danificado - Emergência',
                        'CABOS DROP ROMPIDO': 'Notificado Via Newmonitor',
                        'LIMPEZA DE RUÍDO - MANUTENÇÃO REDE COAXIAL': '(1122)Limpeza De Ruído Corretivo',
                        'ACESSO PROIBIDO AO PONTO DE FALHA': '(3115)Fechada P/ Impossib. Acesso Equips-Cliente Ciente',
                        'INGRESSO DE RUÍDO - LOGRADOURO ATENUADO': '(1122)Limpeza De Ruído Corretivo',
                        'PASSIVO DANFICADO': '(1125)Passivo De Rede Com Defeito - Emrgência',
                        'ATIVO MODULO DE RETORNO COM DEFEITO': '(1124)Amplificador De Rede Com Defeito - Emergência',
                        'FUSÍVEL QUEIMADO': '(1112)Fusível Queimado Na Rede',
                        'LIMPEZA DE RUÍDO - LOGRADOURO IDENTIFICADO SEM FILTRO': '(1122)Limpeza De Ruído Corretivo',
                        'PASSIVO COM ÁGUA': '(1125)Passivo De Rede Com Defeito - Emrgência',
                        'TX / RX COM DEFEITO': '(1214)Equalizacao Do Amplificador',
                        'ATIVO COM ÁGUA': '(1124)Amplificador De Rede Com Defeito - Emergência',
                        'ROMPIMENTO CABO  COAXIAL - VANDALISMO / CORTE': '(1117)Rompimento De Cabo Coaxial',
                        'PASSIVO QUEIMADO': '(1125)Passivo De Rede Com Defeito - Emrgência',
                        'MODULO FONTE DANIFICADO': '(1114)Problema Com Modulo De Fonte',
                        'ROMPIMENTO CABO COAXIAL - CARGA ALTA': '(1117)Rompimento De Cabo Coaxial',
                        'REPROJETO DE REDE COAXIAL': '(2127)Sar/Reprojeto De Rede',
                        'PROTETOR DE SURTO DANIFICADO': '(1116)Problemas Fonte(Bater./Conex/Protetor Surto/Acidente)',
                        'CONEXÃO FONTE DANIFICADA': '(1116)Problemas Fonte(Bater./Conex/Protetor Surto/Acidente)',
                        'ROMPIMENTO CABO  COAXIAL - CABO BANDOLADO': '(1117)Rompimento De Cabo Coaxial',
                        'CHECK LIST FONTE': '(2123)Preventiva De Fonte',
                        'FUSÍVEL DE SAÍDA DA FONTE QUEIMADO': '(1112)Fusível Queimado Na Rede',
                        'LPI DANIFICADA': '(1116)Problemas Fonte(Bater./Conex/Protetor Surto/Acidente)',
                        'ROMPIMENTO CABO COAXIAL - ACIDENTE DE TRANSITO': '(1117)Rompimento De Cabo Coaxial',
                        'ATIVO CONECTOR INTERNO': '(2111)Mau Contato Em Equipamentos De Rede Externa',
                        'ROMPIMENTO CABO  COAXIAL - OBRAS PÚBLICAS': '(1117)Rompimento De Cabo Coaxial',
                        'LEVANTAMENTO REDE COAXIAL': '(2126)Adequações, Construções, Reespinamentos De Rede',
                        'SPI INTERNA FONTE DANIFICADA': '(1116)Problemas Fonte(Bater./Conex/Protetor Surto/Acidente)',
                        'INGRESSO DE RUÍDO - ALTERADO FREQ/MODULAÇÃO': '(1122)Limpeza De Ruído Corretivo',
                        'CABO COAXIAL FURTADO': '(1117)Rompimento De Cabo Coaxial',
                        'CAIXA DE EMENDA DANIFICADA': '(1121)Rompimento De Cabo Optico',
                        'ROMPIMENTO CABO OPTICO - VANDALISMO': '(1121)Rompimento De Cabo Optico',
                        'MANOBRA ÓPTICA - CABO DANIFICADO': '(1121)Rompimento De Cabo Optico',
                        'MANOBRA ÓPTICA - MELHORIA DE REDE': '(1121)Rompimento De Cabo Optico',
                        'MANOBRA ÓPTICA - TESTE  COMUTAÇÃO COLETORES': '(1121)Rompimento De Cabo Optico',
                        'MANOBRA ÓPTICA - MIGRAÇÃO DE SITE': '(1121)Rompimento De Cabo Optico',
                        'MANOBRA ÓPTICA - ENGROSSAMENTO DE ROTA': '(1121)Rompimento De Cabo Optico',
                        'ROMPIMENTO CABO OPTICO - CARGA ALTA': '(1121)Rompimento De Cabo Optico',
                        'ROMPIMENTO CABO OPTICO - CABO BANDOLADO': '(1121)Rompimento De Cabo Optico',
                        'ROMPIMENTO CABO OPTICO - ACIDENTE DE TRANSITO': '(1121)Rompimento De Cabo Optico',
                        'ROMPIMENTO CABO ÓPTICO - TROCA DE POSTE / EMERGENCIAL': '(1121)Rompimento De Cabo Optico',
                        'ROMPIMENTO CABO OPTICO - DANO PARCIAL': '(1121)Rompimento De Cabo Optico',
                        'REPARO EM LINK DE FIBRA ÓPTICA NET - CARGA ALTA': '(1121)Rompimento De Cabo Optico',
                        'ROMPIMENTO CABO OPTICO - CABO QUEIMADO': '(1121)Rompimento De Cabo Optico',
                        'ROMPIMENTO CABO OPTICO - PODA DE ARVORE': '(1121)Rompimento De Cabo Optico',
                        'REPARO EM LINK DE FIBRA ÓPTICA NET - VANDALISMO': '(1121)Rompimento De Cabo Optico',
                        'FIBRA OPTICA ATENUADA': '(1121)Rompimento De Cabo Optico',
                        'REPARO EM LINK DE FIBRA ÓPTICA NET - DANO PARCIAL': '(1121)Rompimento De Cabo Optico',
                        'FECHAMENTO DE FUSÕES': '(1121)Rompimento De Cabo Optico',
                        'LEVANTAMENTO DE ROTA ÓPTICA': '(1121)Rompimento De Cabo Optico',
                        'ALINHAMENTO LINK OPTICO': '(1121)Rompimento De Cabo Optico',
                        'DESVIO DE FIBRA/GRUPO DANIFICADO': '(1121)Rompimento De Cabo Optico',
                        'CABO DROP': '(1121)Rompimento De Cabo Optico',
                        'ROMPIMENTO CABO ÓPTICO - TROCA DE POSTE / PROGRAMADA': '(1121)Rompimento De Cabo Optico',
                        'CORDÃO OPTICO DANIFICADO': '(1121)Rompimento De Cabo Optico',
                        'ROMPIMENTO CABO OPTICO - OBRAS PÚBLICAS': '(1121)Rompimento De Cabo Optico',
                        'SERVICE CABLE DANIFICADO': '(1121)Rompimento De Cabo Optico',
                        'REPARO EM LINK DE FIBRA ÓPTICA NET  TROCA DE POSTE / EMERGENCIAL': '(1121)Rompimento De Cabo Optico',
                        'MANOBRA ÓPTICA - CAIXA DE EMENDA DANIFICADA': '(1121)Rompimento De Cabo Optico',
                        'SUBSTITUIÇÃO DE NAP': '(1121)Rompimento De Cabo Optico',
                        'REPARO EM LINK DE FIBRA ÓPTICA NET - CABO QUEIMADO': '(1121)Rompimento De Cabo Optico',
                        'REPARO EM LINK DE FIBRA ÓPTICA NET - OBRAS PÚBLICAS': '(1121)Rompimento De Cabo Optico',
                        'SUBSTITUIÇÃO DE SPLITTER': '(1121)Rompimento De Cabo Optico',
                        'CONECTOR ÓPTICO DANIFICADO': '(1121)Rompimento De Cabo Optico',
                        'LIMPEZA DE CONECTORES OPTICOS': '(1121)Rompimento De Cabo Optico',
                        'REPARO EM LINK DE FIBRA ÓPTICA NET - PODA DE ARVORE': '(1121)Rompimento De Cabo Optico',
                        'EXECUÇÃO DE NOVOS LINKS': '(1121)Rompimento De Cabo Optico'
                        }



# FUNÇÃO PARA VERIFICAR OUTAGES EM ABERTO NA FERRAMENTA DE RELATORIO DE OUTAGE (RECORRÊNCIA)
def checkelementonw( css_elemento ):
    try:
        driver.find_element( By.CSS_SELECTOR,f"{css_elemento}" ).size
    except NoSuchElementException:
        return False
    return True



# CONFIGURAÇÃO DO CHROME WEBDRIVER
chrome_options = selenium.webdriver.chrome.options.Options()
chrome_options.add_argument( "user-data-dir=Configure seu diretório aqui" )
chrome_driver = r'Configure seu diretório aqui'
driver = webdriver.Chrome( chrome_driver,chrome_options = chrome_options )
wait = WebDriverWait( driver,10 )

# CARREGANDO A PLANILHA BASE DAS CIDADES/LINKS
planilhabase = pd.read_excel( r'Configure seu diretório aqui',sheet_name = 'Plan1',
                              index_col = None )  # , header=['Regional','Cidade CNS','Cidade ATLAS']

# GERANDO AS FUTURAS PLANILHAS PARA A VARREDURA
infoadicional = [ ]
infoadicional = pd.DataFrame( infoadicional,
                              columns = [ 'Regional','Cidade','Node','HP','Imóvel','Outros','Total' ] )
planilhanode = [ ]
planilhanode = pd.DataFrame( planilhanode,columns = [ 'Regional','Cidade','N Notificação','Tipo','Status','Obs' ] )
planilhahp = [ ]
planilhahp = pd.DataFrame( planilhahp,columns = [ 'Regional','Cidade','N Notificação','Tipo','Status','Obs' ] )
planilhamdu = [ ]
planilhamdu = pd.DataFrame( planilhamdu,columns = [ 'Regional','Cidade','N Notificação','Tipo','Status','Obs' ] )

# ABRE A PAGINA DO NEWMONITOR COMO BASE
driver.get( 'link' )
webnm = driver.window_handles[ 0 ]

#

#
#
#
#
# TELA PARA LOGIN NO NEWMONITOR
while True:
    time.sleep( 1 )
    if driver.title == 'nome da guia':
        break
    else:
        os.system( 'cls' )
        print( '\nFerramenta desenvolvida por Daniel Ledezma Vieira.\n' )
        print( 'Faça o login na ferramenta NEWMONITOR' )

# TELA PARA INSERIR LOGIN E SENHA DO ATLAS
inciarok = 'A'
while True:
    if inciarok == '':
        break
    else:
        os.system( 'cls' )
        print( '\nFerramenta desenvolvida por Daniel Ledezma Vieira.\n' )
        print( 'Após logar no NEWMONITOR, digite seu login e senha do ATLAS' )
        loginatlas = input( '\nDigite seu login do ATLAS: ' )
        if loginatlas == '' or len( loginatlas ) <= 5:
            continue
        senhaatlas = getpass.getpass( 'Digite sua senha do ATLAS: ' )
        inciarok = input(
            '\nSe tudo estiver ok, PRESSIONE ENTER, se não, digite qualquer tecla para inserir login e senha novamente: ' )

# ABRE PAGINA DE RECORRENCIA E VOLTA 5 DIAS
driver.get( 'link' )
os.system( 'cls' )
print( '\nFerramenta desenvolvida por Daniel Ledezma Vieira.\n' )
print( 'Voltando alguns dias para busca na recorrência' )
time.sleep( 1 )
d = datetime.now() - timedelta( days = 5 )
d = d.strftime( "%d/%m/%Y" )
WebDriverWait( driver,15 ).until( EC.presence_of_element_located( (By.XPATH,'//input[@id="edt_de"]') ) ).clear()
time.sleep( 0.1 )
driver.find_element( By.XPATH,'//input[@id="edt_de"]' ).send_keys( d )

os.system( 'cls' )
print( '\nFerramenta desenvolvida por Daniel Ledezma Vieira.\n' )
print(
    'A Varredura irá iniciar em 5 segundos.\n\nFique tranquilo(a) que está configurado para corrigir páginas com mensagens de erro, apenas aguarde que o programa corrige automaticamente e segue a varredura.' )
print(
    '\nCaso você ouça que está tocando o som do Windows ou está travado, volte até a pagina inicial do atlas da cidade, ou clique em logout e aguarde.' )
time.sleep( 5 )

##
##
## Possivel back quando da o erro após o login
## css selector 'body > center:nth-child(1) > h4 > a'
##
##

print( '\n\n>>>>>INICIANDO A VARREDURA<<<<<' )

# Inicia a varredura de acordo com a planilha
for index,linhabase in planilhabase.iterrows():
    driver.execute_script( f"window.open('{linhabase[ 'Link' ]}','_blank');" )

    contadorhp = 0
    contadornode = 0
    contadormdu = 0

    # Entra no Loop dentro do Atlas e validação em caso de identificação de ticket
    a = 1
    while a == 1:
        try:
            print( f"\nIniciando verificação na cidade de {linhabase[ 'Cidade NM' ]}." )
            # Entra no Atlas e pega a qtd de notificações aberta
            webatlas = driver.window_handles[ 1 ]
            driver.switch_to.window( webatlas )

            #
            # Depois daqui, as vezes acontece um "404 Not Found"
            #                  
            try:
                WebDriverWait( driver,20 ).until(
                    EC.presence_of_element_located( (By.XPATH,"//input[@name='pUs_Codigo']") ) ).click()
            except:
                while driver.title == '404 Not Found' or driver.title == 'nome da guia':
                    print( '\n\nParece que a página não carregou.\nRecarregando a página\n' )
                    driver.refresh()
                    time.sleep( 3 )
                WebDriverWait( driver,20 ).until(
                    EC.presence_of_element_located( (By.XPATH,"//input[@name='pUs_Codigo']") ) ).click()
            WebDriverWait( driver,20 ).until(
                EC.presence_of_element_located( (By.XPATH,"//input[@name='pUs_Codigo']") ) ).clear()
            WebDriverWait( driver,15 ).until(
                EC.presence_of_element_located( (By.XPATH,"//input[@name='pUs_Codigo']") ) ).send_keys( loginatlas )
            WebDriverWait( driver,15 ).until(
                EC.presence_of_element_located( (By.XPATH,"//input[@name='pSenha']") ) ).click()
            WebDriverWait( driver,15 ).until(
                EC.presence_of_element_located( (By.XPATH,"//input[@name='pSenha']") ) ).clear()
            WebDriverWait( driver,15 ).until(
                EC.presence_of_element_located( (By.XPATH,"//input[@name='pSenha']") ) ).send_keys( senhaatlas )
            WebDriverWait( driver,15 ).until(
                EC.presence_of_element_located( (By.XPATH,"//select[@name='pCi_Codigo']") ) ).click()
            WebDriverWait( driver,15 ).until( EC.presence_of_element_located(
                (By.XPATH,f"//option[normalize-space()= '{linhabase[ 'Cidade ATLAS' ]}']") ) ).click()
            WebDriverWait( driver,15 ).until(
                EC.presence_of_element_located( (By.XPATH,"//input[@value='Confirma']") ) ).click()
            try:  # resolver o b.o do confirma após login do atlas
                driver.find_element( By.CSS_SELECTOR,
                                     'body > table:nth-child(7) > tbody > tr > td:nth-child(1) > form > input[type=submit]:nth-child(4)' ).click()
                # driver.find_element(By.CSS_SELECTOR, 'body > table > tbody > tr > td:nth-child(1) > a > img').click()
            except:
                try:
                    driver.find_element( By.CSS_SELECTOR,'body > center:nth-child(1) > h4 > a' ).click()
                except:
                    pass

            WebDriverWait( driver,3 ).until(
                EC.presence_of_element_located( (By.XPATH,"//img[@alt='bot5.jpg (8850 bytes)']") ) ).click()
            WebDriverWait( driver,15 ).until(
                EC.presence_of_element_located( (By.XPATH,"(//img[@alt='bot1.jpg (8776 bytes)'])[6]") ) ).click()
            WebDriverWait( driver,15 ).until(
                EC.presence_of_element_located( (By.XPATH,"(//a[contains(text(),'Notificações')])[3]") ) ).click()
            WebDriverWait( driver,30 ).until(
                EC.presence_of_element_located( (By.XPATH,"//small[contains(.,'Sem Filtro')]") ) ).click()

            # Pegando a qtd de notificações aberta
            qtd = driver.find_element( By.CSS_SELECTOR,'body > center > center > table > caption > font' ).text
            qtd = qtd.split( ' ' )
            qtd = qtd[ 2 ]
            qtd = int( qtd )
            qtdoriginal = qtd
            # Em caso de haver notificações
            if qtd > 0:
                notqtd = 2
                qtd = qtd + 1
                print( 'A cidade de',linhabase[ 'Cidade NM' ],'não está vazia, foram encontradas',qtdoriginal,
                       'notificações.\nIniciando verificação.' )

                # Verificando os tipo da notificação
                while notqtd <= qtd:  # Coleta as informações importante da lista
                    cssnotificacao = 'body > center > center > table > tbody > tr:nth-child(' + str(
                        notqtd ) + ') > td:nth-child(1)'
                    cssobs = 'body > center > center > table > tbody > tr:nth-child(' + str(
                        notqtd ) + ') > td:nth-child(3)'
                    csstipo = 'body > center > center > table > tbody > tr:nth-child(' + str(
                        notqtd ) + ') > td:nth-child(5)'
                    csslinkfechar = 'body > center > center > table > tbody > tr:nth-child(' + str(
                        notqtd ) + ') > td:nth-child(12) > a'
                    notqtd = notqtd + 1

                    # Em caso de ser HP:
                    if driver.find_element( By.CSS_SELECTOR,csstipo ).text == 'HP:':
                        ftipo = driver.find_element( By.CSS_SELECTOR,csstipo ).text
                        fnot = driver.find_element( By.CSS_SELECTOR,cssnotificacao ).text
                        fobs = driver.find_element( By.CSS_SELECTOR,cssobs ).text
                        contadorhp = contadorhp + 1
                        fstatus = 'INC não localizado'
                        fobsbkp = fobs
                        fobs = fobs.split( ' ' )  # Divide por espaço e tenta localizar

                        # Busca possiveis tickets
                        for n,item in enumerate( fobs ):
                            if item[ 0:4 ] == 'INC0':
                                fobs = item
                                fstatus = 'Identificado INC:'
                                break
                            else:
                                continue
                        if fobs[ 0:5 ] != 'INC00':
                            for n,item in enumerate( fobs ):
                                if len( item ) == 7:
                                    if item[ 0 ] == '4':
                                        fobs = item
                                        fstatus = 'Identificado INC:'
                                        break
                                    else:
                                        continue
                                else:
                                    continue

                        if fstatus == 'INC não localizado':
                            fobs = fobsbkp

                        novalinha = {'Regional': linhabase[ 'Regional' ],'Cidade': linhabase[ 'Cidade ATLAS' ],
                                     'N Notificação': fnot,'Tipo': ftipo,'Status': fstatus,'Obs': fobs}
                        planilhahp = planilhahp.append( novalinha,ignore_index = True )

                    # Em caso de Node
                    if driver.find_element( By.CSS_SELECTOR,csstipo ).text[ 0:4 ] == 'Node':
                        ftipo = driver.find_element( By.CSS_SELECTOR,csstipo ).text
                        fnot = driver.find_element( By.CSS_SELECTOR,cssnotificacao ).text
                        fobs = driver.find_element( By.CSS_SELECTOR,cssobs ).text
                        contadornode = contadornode + 1
                        otgfechamento = ''
                        fobsbkp = fobs
                        fobs = fobs.split( ' ' )  # Divide por espaço e tenta localizar

                        # Busca possiveis tickets
                        for n,item in enumerate( fobs ):
                            if len( item ) == 8:
                                if item[ 0:2 ] == '12':
                                    fobs = item
                                    break
                                else:
                                    continue
                            else:
                                continue

                        if len( fobs ) == 8:
                            if fobs[ 0:2 ] == '12':
                                verificastatus = 'link' + fobs
                                print( '\nVerificando ticket:' )
                                print( verificastatus )
                                time.sleep( 1 )
                                driver.execute_script(
                                    f"window.open('{verificastatus}','_blank');" )  # .format(link)) (dentro do fim das aspas)
                                webstatus = driver.window_handles[ 2 ]
                                driver.switch_to.window( webstatus )
                                time.sleep( 1 )
                                try:
                                    fstatus = WebDriverWait( driver,10 ).until( EC.presence_of_element_located( (
                                                                                                                By.CSS_SELECTOR,
                                                                                                                '#divTicket > table > tbody > tr:nth-child(11) > td.ticket_value') ) ).text
                                    if fstatus == "Fechado":
                                        WebDriverWait( driver,10 ).until( EC.presence_of_element_located( (
                                                                                                          By.CSS_SELECTOR,
                                                                                                          '#divDados > table > tbody > tr:nth-child(1) > td:nth-child(4) > a') ) ).click()
                                        otgfechamento = WebDriverWait( driver,10 ).until(
                                            EC.presence_of_element_located( (By.CSS_SELECTOR,
                                                                             '#divComp > table > tbody > tr:nth-child(4) > td.ticket_value') ) ).text
                                        otgbase = WebDriverWait( driver,10 ).until( EC.presence_of_element_located( (
                                                                                                                    By.CSS_SELECTOR,
                                                                                                                    '#divComp > table > tbody > tr:nth-child(3) > td.ticket_value') ) ).text
                                except:
                                    driver.get( 'link' )
                                    time.sleep( 1 )
                                    fstatus = WebDriverWait( driver,10 ).until( EC.presence_of_element_located( (
                                                                                                                By.CSS_SELECTOR,
                                                                                                                '#divTicket > table > tbody > tr:nth-child(11) > td.ticket_value') ) ).text

                                if driver.find_element( By.CSS_SELECTOR,
                                                        '#divTicket > table > tbody > tr:nth-child(6) > td.ticket_value' ).text != \
                                        linhabase[ 'Cidade NM' ]:
                                    fstatus = 'Ticket não localizado, possivelmente cidade divergente (verificar a planilha)'
                                    fobs = fobsbkp
                                time.sleep( 1 )
                                driver.close()
                                driver.switch_to.window( webatlas )
                        else:
                            fstatus = 'Ticket não localizado'
                            fobs = fobsbkp

                        if fstatus == 'Ticket não localizado':
                            fnode = ftipo
                            fnode = fnode.split( ' ' )
                            fnode = fnode[ -1 ]
                            if len( fnode ) > 6:
                                fnode = 'nodenotfound'

                            driver.switch_to.window( webnm )
                            WebDriverWait( driver,10 ).until(
                                EC.presence_of_element_located( (By.CSS_SELECTOR,'#cmb_cidade') ) ).click()
                            WebDriverWait( driver,10 ).until( EC.presence_of_element_located(
                                (By.XPATH,f"//option[contains(text(), '{linhabase[ 'Cidade NM' ]}')]") ) ).click()
                            WebDriverWait( driver,10 ).until(
                                EC.presence_of_element_located( (By.CSS_SELECTOR,'#edt_node_mdu') ) ).clear()
                            WebDriverWait( driver,10 ).until(
                                EC.presence_of_element_located( (By.CSS_SELECTOR,'#edt_node_mdu') ) ).send_keys( fnode )
                            WebDriverWait( driver,10 ).until(
                                EC.presence_of_element_located( (By.NAME,'submit') ) ).click()
                            time.sleep( 2 )

                            checknw = 3
                            elemento = '#divResposta > table > tbody > tr:nth-child(' + str( checknw ) + ')'

                            while checkelementonw( elemento ) == True:
                                verificanw = driver.find_element( By.CSS_SELECTOR,elemento ).text
                                print( 'verificando node:',fnode,verificanw )
                                if verificanw == 'Sem dados disponíveis':
                                    fstatus = 'Ticket não localizado'
                                    fobs = fobsbkp
                                    driver.switch_to.window( webatlas )
                                    break
                                elementoxpath = '//*[@id="divResposta"]/table/tbody/tr[' + str( checknw ) + ']/td[11]'
                                testnd = driver.find_element( By.XPATH,elementoxpath ).text
                                if testnd == 'N/D':
                                    fstatus = 'Ticket em aberto'
                                    fobs = fobsbkp
                                    driver.switch_to.window( webatlas )
                                    break
                                else:
                                    fstatus = 'Ticket fechado'
                                    fobs = fobsbkp
                                    checknw = checknw + 1
                                    elemento = '#divResposta > table > tbody > tr:nth-child(' + str( checknw ) + ')'
                            print( fstatus,fobs )
                            driver.switch_to.window( webatlas )

                        if fstatus in [ "Ticket fechado","Fechado" ]:

                            print( 'Status',fstatus )

                            flinkfechar = driver.find_element( By.CSS_SELECTOR,csslinkfechar ).get_attribute( 'href' )
                            driver.switch_to.new_window()
                            webatlasfechar = driver.current_window_handle
                            driver.switch_to.window( webatlasfechar )
                            driver.get( flinkfechar )
                            notvalidacaofechar = 'body > center > font > b'
                            fnotvalidacao = driver.find_element( By.CSS_SELECTOR,notvalidacaofechar ).text
                            fnotvalidacao = fnotvalidacao.split( '.' )
                            fnotvalidacao = fnotvalidacao[ -1 ]
                            print( '\n\n\n' )
                            print( 'Notificação:',fnot )
                            print( 'Not validação',fnotvalidacao )  # fnotvalidacao

                            # if fnot == fnotvalidacao:
                            # print('Notificação igual')
                            # if otgbase == 'REDE COAXIAL' or otgbase == 'OUTROS':
                            #    if otgfechamento in fechamentos_rf_atlas:
                            #        print('Fechamento do outage: ', otgfechamento,'\nFechamento para o Atlas',fechamentos_rf_atlas[otgfechamento])
                            #        otgfechamento = fechamentos_rf_atlas[otgfechamento]
                            #    else:
                            #        print('-- ',otgfechamento,'-- | Fechamento não existe no cadastro')
                            #        otgfechamento = 'Notificado Via Newmonitor'
                            # else:
                            #    print('-- ',otgfechamento,'-- | Fechamento não existe no cadastro')
                            #    otgfechamento = 'Notificado Via Newmonitor'

                            if fnot == fnotvalidacao:
                                if otgfechamento in fechamentos_rf_atlas:
                                    print( 'Fechamento do outage: ',otgfechamento,'\nFechamento para o Atlas',
                                           fechamentos_rf_atlas[ otgfechamento ] )
                                    otgfechamento = fechamentos_rf_atlas[ otgfechamento ]
                                else:
                                    print( 'Fechamento:',otgfechamento,'| não localizado no cadastro:' )
                                    otgfechamento = 'Notificado Via Newmonitor'

                                WebDriverWait( driver,15 ).until( EC.presence_of_element_located(
                                    (By.XPATH,f"//option[normalize-space()= '{otgfechamento}']") ) ).click()
                                time.sleep( 1 )
                                WebDriverWait( driver,15 ).until( EC.presence_of_element_located( (By.CSS_SELECTOR,
                                                                                                   'body > center > form > table:nth-child(6) > tbody > tr > td:nth-child(1) > input[type=submit]') ) ).click()
                                WebDriverWait( driver,15 ).until( EC.presence_of_element_located( (By.CSS_SELECTOR,
                                                                                                   'body > center > form > table > tbody > tr:nth-child(5) > td > input[type=submit]') ) ).click()
                                time.sleep( 1 )

                                #
                                #
                                # FAZER A VALIDAÇÃO DO FECHAMENTO AQUI \/
                                #
                                #

                                valida_fechamento = WebDriverWait( driver,15 ).until(
                                    EC.presence_of_element_located( (By.CSS_SELECTOR,'body > center > font') ) ).text
                                if valida_fechamento == 'Notificações':
                                    print( 'Notificação fechada com sucesso' )
                                else:
                                    winsound.PlaySound( "SystemHand",winsound.SND_NOSTOP )
                                    print( f'\nNotificação {fnot} da cidade de ',linhabase[ 'Cidade NM' ],
                                           ' não foi fechada corretamente, necessário verificar posteriormente' )
                                    fstatus = 'BUG - necessário fechar manualmente'

                                fobsbkp = fobs
                                fobs = f'### -- Noficação {fnot} fechada como: ' + otgfechamento + " Obs do ticket: " + fobs

                            driver.close()
                            driver.switch_to.window( webatlas )

                        novalinha2 = {'Regional': linhabase[ 'Regional' ],'Cidade': linhabase[ 'Cidade ATLAS' ],
                                      'N Notificação': fnot,'Tipo': ftipo,'Status': fstatus,'Obs': fobs}
                        planilhanode = planilhanode.append( novalinha2,ignore_index = True )

                    if driver.find_element( By.CSS_SELECTOR,csstipo ).text[ 0:6 ] == 'Imóvel':
                        ftipo = driver.find_element( By.CSS_SELECTOR,csstipo ).text
                        fnot = driver.find_element( By.CSS_SELECTOR,cssnotificacao ).text
                        fobs = driver.find_element( By.CSS_SELECTOR,cssobs ).text
                        contadormdu = contadormdu + 1
                        fobsbkp = fobs
                        fobs = fobs.split( ' ' )  # Divide por espaço e tenta localizar

                        # Busca possiveis tickets
                        for n,item in enumerate( fobs ):
                            if len( item ) == 8:
                                if item[ 0:2 ] == '12':
                                    fobs = item
                                    break
                                else:
                                    fstatus = 'Ticket não localizado'
                                    fobs = fobsbkp
                                    continue
                            else:
                                fstatus = 'Ticket não localizado'
                                fobs = fobsbkp
                                continue

                        # if len(fobs) == 8:
                        # print('@15')
                        # if fobs[0:2] == '12':
                        # print('@16')
                        # verificastatus = 'link'+fobs
                        # print(verificastatus)
                        # time.sleep(1)
                        # driver.execute_script(f"window.open('{verificastatus}','_blank');") #.format(link)) (dentro do fim das aspas)
                        # webstatus = driver.window_handles[2]
                        # driver.switch_to.window(webstatus)
                        # time.sleep(1)
                        # try:
                        # print('@17')
                        # fstatus = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#divTicket > table > tbody > tr:nth-child(11) > td.ticket_value'))).text
                        # if fstatus == "Fechado":
                        # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#divDados > table > tbody > tr:nth-child(1) > td:nth-child(4) > a'))).click()
                        # otgfechamento = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#divComp > table > tbody > tr:nth-child(4) > td.ticket_value'))).text
                        # except:
                        # print('@18')
                        # driver.get('link')
                        # time.sleep(1)
                        # fstatus = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#divTicket > table > tbody > tr:nth-child(11) > td.ticket_value'))).text

                        # print('@19')
                        # if driver.find_element(By.CSS_SELECTOR, '#divTicket > table > tbody > tr:nth-child(6) > td.ticket_value').text != linhabase['Cidade NM']:
                        # print('@20')
                        # fstatus = 'Ticket não localizado'
                        # fobs = fobsbkp
                        # time.sleep(1)
                        # driver.close()
                        # driver.switch_to.window(webatlas)

                        # else:
                        # fstatus = 'Ticket não localizado'
                        # fobs = fobsbkp

                        # if fstatus in ["Ticket fechado","Fechado"]:
                        # flinkfechar = driver.find_element(By.CSS_SELECTOR, csslinkfechar).get_attribute('href')
                        # driver.switch_to.new_window()
                        # webatlasfechar = driver.current_window_handle
                        # driver.switch_to.window(webatlasfechar)
                        # driver.get(flinkfechar)
                        # notvalidacaofechar = 'body > center > font > b'
                        # fnotvalidacao = driver.find_element(By.CSS_SELECTOR, notvalidacaofechar).text
                        # fnotvalidacao = fnotvalidacao.split('.')
                        # fnotvalidacao = fnotvalidacao[-1]
                        # print('\n\n\n')
                        # print('Notificação:', fnot)
                        # print('Not validação', fnotvalidacao, otgfechamento) #fnotvalidacao

                        # if fnot == fnotvalidacao:
                        # if otgfechamento in fechamentos_rf_atlas:
                        # print('fechamento existe')
                        # otgfechamento = fechamentos_rf_atlas[otgfechamento]
                        # else:
                        # print('fechamento não existe no cadastro')
                        # otgfechamento = 'Notificado Via Newmonitor'
                        # WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, f"//option[normalize-space()= '{otgfechamento}']"))).click()
                        # time.sleep(1)
                        # WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'body > center > form > table:nth-child(6) > tbody > tr > td:nth-child(1) > input[type=submit]'))).click()
                        # inciarok = '3'
                        # inciarok = input('Se tudo estiver ok, digite xxx: ')
                        # if inciarok == 'xxx':
                        # WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'body > center > form > table > tbody > tr:nth-child(5) > td > input[type=submit]'))).click()
                        # time.sleep(1)
                        # fobsbkp = fobs
                        # fobs = f'### -- Noficação {fnot} fechada -- ### mdu?  ' + otgfechamento + fobs
                        # fobs = fobsbkp

                        # driver.close()
                        # driver.switch_to.window(webatlas)

                        novalinha3 = {'Regional': linhabase[ 'Regional' ],'Cidade': linhabase[ 'Cidade ATLAS' ],
                                      'N Notificação': fnot,'Tipo': ftipo,'Status': fstatus,'Obs': fobs}
                        planilhamdu = planilhamdu.append( novalinha3,ignore_index = True )

            else:
                print( f"Não consta notificações na cidade de {linhabase[ 'Cidade NM' ]}." )

            novalinha = {'Regional': linhabase[ 'Regional' ],'Cidade': linhabase[ 'Cidade ATLAS' ],'Total': qtdoriginal,
                         'Node': contadornode,'HP': contadorhp,'Imóvel': contadormdu}  # 'Outros':contaodroutros
            infoadicional = infoadicional.append( novalinha,ignore_index = True )
            WebDriverWait( driver,10 ).until(
                EC.presence_of_element_located( (By.XPATH,"//img[@alt='Finaliza Sessão']") ) ).click()
            time.sleep( 1 )
            driver.close()
            driver.switch_to.window( webnm )
            a = 2

        except:

            try:
                driver.find_element( By.CSS_SELECTOR,
                                     'body > center:nth-child(3) > form > input[type=submit]' ).click()  # body > center:nth-child(4) > form > input[type=submit] - antigo
                print( 'Recomeçando a cidade de:',linhabase[ 'Cidade NM' ] )


            except:

                #
                # IMPLEMENTAR LOG DE ERRO AQUI
                #
                try:
                    driver.find_element( By.CSS_SELECTOR,'body > center > form > input[type=submit]' ).click()
                except:
                    pass

                winsound.PlaySound( "SystemHand",winsound.SND_NOSTOP )
                time.sleep( 2 )

# FAZ O CALCULO DAS PLANILHAS
infoadicional[ 'Outros' ] = infoadicional[ 'Total' ] - infoadicional[ 'Node' ] - infoadicional[ 'HP' ] - infoadicional[
    'Imóvel' ]
novalinha = {'Regional': 'Nenhum','Cidade': 'Total','Node': infoadicional[ 'Node' ].sum(),
             'HP': infoadicional[ 'HP' ].sum(),'Imóvel': infoadicional[ 'Imóvel' ].sum(),
             'Outros': infoadicional[ 'Outros' ].sum(),'Total': infoadicional[ 'Total' ].sum()}
infoadicional = infoadicional.append( novalinha,ignore_index = True )

# PEGA O DIA E HORARIO ATUAL PARA ACRESCENTAR NO NOME DA PLANILHA
now = datetime.now()
s1 = now.strftime( "%Y-%m-%d -- %HH%M" )

# FAZ A GERAÇÃO/UNIÃO DAS PLANILHAS
with pd.ExcelWriter( fr'C:\Atlas\Varredura Atlas {s1} .xlsx',engine = 'openpyxl' ) as writer:
    infoadicional.to_excel( writer,sheet_name = 'Cidades',index = False )
    planilhanode.to_excel( writer,sheet_name = 'Node',index = False )
    planilhahp.to_excel( writer,sheet_name = 'Hp',index = False )
    planilhamdu.to_excel( writer,sheet_name = 'Imóvel',index = False )

driver.close()
print( '\n\n\nVarredura finalizada\nVarredura finalizada\nVarredura finalizada\nVarredura finalizada' )
winsound.PlaySound( "SystemHand",winsound.SND_NOSTOP )