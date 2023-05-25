from datetime import datetime
import time
import os
import pdfplumber
import shutil

import funcoes   as f
import importar  as sicredi
import importarc as cresol

#Relatorio
from docxtpl import DocxTemplate
from docx2pdf import convert

#Interface
import PySimpleGUI as sg

#registro windows
import winreg as wrg

vPath    = 'C:/Temp/Fichas_Graficas'
versao   = ''
lancamentos = []
fichas_importadas = []
tela = None

def alimentaDetalhesRelatorio(lista, data, descricao, valor, correcao, corrigido, juros, total):
    lista.append({"data":data, "descricao":descricao, "valor":valor,"correcao":correcao,"corrigido":corrigido,"juros":juros,"total":total})

def converterPDF(vCaminho_arquivo):    
    try:
        ## Carrega arquivo
        pdf = pdfplumber.open(vCaminho_arquivo)

        vNome        = vCaminho_arquivo.split('/')
        vNomeArquivo = vNome[-1][:-4]

        ## Converte PDF em TXT
        for pagina in pdf.pages:
            texto = pagina.extract_text(x_tolerance=1)

            vArquivo_final = vPath + '/' + vNomeArquivo + '.txt'
            
            with open(vArquivo_final, 'a') as arquivo_txt:
                arquivo_txt.write(str(texto))       
                
        pdf.close()         

        return vArquivo_final
    
    except Exception as erro:
        with open(vPath + '/ERRO_LOG.txt', 'a') as arquivo_txt:
            arquivo_txt.write(str(erro))
    
        return None
    
def identificaVersao(caminho_txt):
    with open(caminho_txt, 'r') as arquivo_txt:        
        vlinha = 1
        for linha in arquivo_txt:
            if vlinha <= 20:
                if("CRESOL" in linha):
                    return 'cresol'
                    break

                if("COOP CRED POUP E INVEST" in linha):
                    return 'sicredi'
                    break
            else:
                break 
            
            vlinha = vlinha + 1
    
    return versao

def verificaLicenca():
    try:
        registry_key = wrg.OpenKey(wrg.HKEY_CURRENT_USER, r"SOFTWARE\\CorretorMonetario", 0,
                                       wrg.KEY_READ)
        value, regtype = wrg.QueryValueEx(registry_key, "Chave")
        wrg.CloseKey(registry_key)
        return value
    except WindowsError:
        return None
    
def licenca():

    local = wrg.HKEY_CURRENT_USER            
    soft  = wrg.OpenKeyEx(local, r"SOFTWARE\\")
  

    sg.theme('Reddit')

    layout_ativar = [    
                [sg.Text(text='Chave de Ativação', text_color="BLACK", font=("Arial"))],                                        
                [sg.Text(text='Insira a chave para ativa o produto: ', text_color="BLACK", font=("Arial",10))],                      
                [sg.InputText(key='ed_chave')],                
                [sg.Button('Registrar'), sg.Button('Fechar')]      
             ]
    tela_ativar = sg.Window('Ativar Produto', layout_ativar, modal=True)

    while True:                    
        eventos, valores = tela_ativar.read(timeout=0.1)
        
        if eventos == 'Registrar':            
            valor = valores['ed_chave']            
            key_1 = wrg.CreateKey(soft, "CorretorMonetario")
                        
            wrg.SetValueEx(key_1, "Chave", 0, wrg.REG_SZ,
                        valor)            
            
            if key_1:
                wrg.CloseKey(key_1)

            tela_ativar.Close()

        if eventos == sg.WINDOW_CLOSED:
            break
        if eventos == 'Fechar':
            tela_ativar.close()            
    
def main():            
    nome_arquivo       = ''
    tipo_arquivo       = ''
    versao             = '' 
    informacoes        = '' 
    arquivo_importar   = ''
    remove_arquivo     = False    
    global tela
    global fichas_importadas

    if not os.path.isdir(vPath): # vemos de este diretorio ja existe        
        os.mkdir(vPath)
        os.mkdir(vPath + '/Processados')  
    
    sg.theme('Reddit')
    lista_arquivo = []
    lista_arquivo.append(sg.Input(key='-INPUT-', enable_events=True))
    lista_arquivo.append(sg.FileBrowse(button_text='Procurar', enable_events=True, key='arquivo'))
    frame_arquivo = sg.Frame('Selecione a Ficha Gráfica', [lista_arquivo], expand_x=True)
    
    lista_indices = []
    lista_indices.append(sg.Checkbox(text='IGPM', key='ed_igpm'))
    lista_indices.append(sg.Checkbox(text='IPCA', key='ed_ipca'))
    lista_indices.append(sg.Checkbox(text='CDI', key='ed_cdi'))
    lista_indices.append(sg.Checkbox(text='INPC', key='ed_inpc'))
    lista_indices.append(sg.Checkbox(text='TR', key='ed_tr'))
    frame_indices = sg.Frame('Índices de Correção', [lista_indices], expand_x=True)

    lista_informacoes = []
    lista_informacoes.append(sg.Multiline(default_text=informacoes, expand_x=True, size=(None, 5), key="ed_informacoes", enable_events=True, auto_refresh=True) )    
    frame_informacoes = sg.Frame('Informações da Ficha Gráfica', [lista_informacoes], expand_x=True, visible=False, key='Info')      
    
    lista_multa = []
    lista_multa.append(sg.Text(text="Percentual:", font=("Arial",10, "bold")))    
    lista_multa.append(sg.Input(key="ed_multa_perc", enable_events=True, size=(10,1)))    
    lista_multa.append(sg.Text(text="Valor Fixo:", font=("Arial",10, "bold")))    
    lista_multa.append(sg.Input(key="ed_multa_valor", enable_events=True,size=(20,1)))    
    lista_multa.append(sg.Text(text="Incidência:", font=("Arial",10, "bold")))    
    lista_multa.append(sg.Combo(['Sobre o Valor Corrigido+Juros Principais','Sobre o Valor Corrigido','Sobre o Valor Original+Juros Principais','Sobre o Valor Original'], key="ed_multa_incidencia", enable_events=True))    
    frame_multa = sg.Frame('Multa', [lista_multa], expand_x=True, key="frame_multa")

    lista_honorarios = []
    lista_honorarios.append(sg.Text(text="Percentual:", font=("Arial",10, "bold")))    
    lista_honorarios.append(sg.Input(key="ed_honorarios_perc", enable_events=True, size=(10,1)))    
    lista_honorarios.append(sg.Text(text="Valor Fixo:", font=("Arial",10, "bold")))    
    lista_honorarios.append(sg.Input(key="ed_honorarios_valor", enable_events=True, size=(20,1)) )    
    frame_honorarios = sg.Frame('Honorários', [lista_honorarios], expand_x=True, key="frame_honorarios")

    lista_outros = []
    lista_outros.append(sg.Text(text="Valor R$:", font=("Arial",10, "bold")))    
    lista_outros.append(sg.Input(key="ed_outros_valor", size=(20,1), default_text='0,00', enable_events=True))        
    frame_outros = sg.Frame('Outros Valores', [lista_outros], expand_x=True)    
    
    ## Campos para defini~çao dos parâmetros padrão (Aba Parâmetros)    
    lista_indices_parametros = []
    lista_indices_parametros.append(sg.Checkbox(text='IGPM', key='ed_igpm_param'))
    lista_indices_parametros.append(sg.Checkbox(text='IPCA', key='ed_ipca_param'))
    lista_indices_parametros.append(sg.Checkbox(text='CDI', key='ed_cdi_param'))
    lista_indices_parametros.append(sg.Checkbox(text='INPC', key='ed_inpc_param'))
    lista_indices_parametros.append(sg.Checkbox(text='TR', key='ed_tr_param'))
    frame_indices_parametros = sg.Frame('Índices de Correção', [lista_indices_parametros], expand_x=True)
    
    lista_multa_parametros = []
    lista_multa_parametros.append(sg.Text(text="Percentual:", font=("Arial",10, "bold")))    
    lista_multa_parametros.append(sg.Input(key="ed_multa_perc_param", enable_events=True, size=(10,1)))    
    lista_multa_parametros.append(sg.Text(text="Valor Fixo:", font=("Arial",10, "bold")))    
    lista_multa_parametros.append(sg.Input(key="ed_multa_valor_param", enable_events=True,size=(20,1)))    
    lista_multa_parametros.append(sg.Text(text="Incidência:", font=("Arial",10, "bold")))    
    lista_multa_parametros.append(sg.Combo(['Sobre o Valor Corrigido+Juros Principais','Sobre o Valor Corrigido','Sobre o Valor Original+Juros Principais','Sobre o Valor Original'], key="ed_multa_incidencia_param", enable_events=True))    
    frame_multa_parametros = sg.Frame('Multa', [lista_multa_parametros], expand_x=True, key="frame_multa_param")

    lista_honorarios_parametros = []
    lista_honorarios_parametros.append(sg.Text(text="Percentual:", font=("Arial",10, "bold")))    
    lista_honorarios_parametros.append(sg.Input(key="ed_honorarios_perc_param", enable_events=True, size=(10,1)))    
    lista_honorarios_parametros.append(sg.Text(text="Valor Fixo:", font=("Arial",10, "bold")))    
    lista_honorarios_parametros.append(sg.Input(key="ed_honorarios_valor_param", enable_events=True, size=(20,1)) )    
    frame_honorarios_parametros = sg.Frame('Honorários', [lista_honorarios_parametros], expand_x=True, key="frame_honorarios_param")

    lista_outros_parametros = []
    lista_outros_parametros.append(sg.Text(text="Valor R$:", font=("Arial",10, "bold")))    
    lista_outros_parametros.append(sg.Input(key="ed_outros_valor_param", size=(20,1), default_text='0,00', enable_events=True))        
    frame_outros_parametros = sg.Frame('Outros Valores', [lista_outros_parametros], expand_x=True)     
    
    principal = [    
                [sg.Text(text='Corretor Monetário', text_color="Black", font=("Arial",22, "bold"), expand_x=True, justification='center')],                                      
                [frame_arquivo],      
                [frame_informacoes],                                              
                [frame_indices],    
                [frame_multa],                
                [frame_honorarios],
                [frame_outros],                
                [sg.Text(text='Aguardando Operação', key='ed_situacao', text_color="green")],
                [sg.ProgressBar(100, orientation='h', size=(50, 4), key='progressbar', visible=False, expand_x=True)],
                [sg.Button('Calcular'), sg.Button('Fechar'), sg.Text('Versão: 1.0.0', expand_x=True, justification='right', font=("Verdana",10))]      
             ]
    
    lista_filtros = []
    lista_filtros.append(sg.Text("Título"))
    lista_filtros.append(sg.Input(key="ed_filtro_titulo", size=(20,1), enable_events=True))
    lista_filtros.append(sg.Text("Associado"))
    lista_filtros.append(sg.Input(key="ed_filtro_associado", size=(30,1), enable_events=True))
    lista_filtros.append(sg.Button("Filtrar", key='btn_atualizar_importadas'))
    frame_filtros = sg.Frame('Filtros', [lista_filtros], expand_x=True)    
    importadas = [    
                    [sg.Text(text='Fichas Importadas', text_color="black", font=("Arial",12), expand_x=True, justification='center')],
                    [frame_filtros],
                    [sg.Table(values=fichas_importadas, headings=['ID', 'Título', 'Versão', 'Associado', 'N° Parcelas', 'Valor Financiado', 'Taxa Juros'], auto_size_columns=True, display_row_numbers=False, justification='center', key='-TABLE-', selected_row_colors='red on yellow', enable_events=True, expand_x=True, expand_y=True,enable_click_events=True)]                    
                 ]    

    parametros = [    
                    [sg.Text(text='Parâmetros', text_color="black", font=("Arial",12), expand_x=True, justification='center')],                                        
                    [sg.Text(text='Parâmetros Padrão do Sistema', text_color="black", font=("Arial",9), expand_x=True, justification='left')],
                    [sg.Text(text='Atenção: Opções definidas aqui serão carregadas como padrão para todos ao inicializar uma nova importação.', text_color="red", font=("Verdana",7), expand_x=True, justification='left')],
                    [frame_indices_parametros],
                    [frame_multa_parametros],
                    [frame_honorarios_parametros],
                    [frame_outros_parametros],
                    [sg.Button('Salvar Parâmetros', key='btn_salvar_parametros_gerais'), sg.Button('Cancelar')]
                 ]

    tabgrp = [
                [sg.TabGroup
                    (
                        [
                            [sg.Tab('Principal', principal, title_color='Red',border_width =10), sg.Tab('Importadas', importadas, key='aba_importadas'), sg.Tab('Parâmetros', parametros, key='aba_parametros')]                            
                        ], enable_events=True, key='-abas-'
                    )
                ]
             ]  
    

    busca_licena = verificaLicenca() 
    if ((busca_licena == None) or (busca_licena != 'THMPV-77D6F-94376-8HGKG-VRDRQ')):
        licenca()

    def atualizaInfo():
        global informacoes
        informacoes = ''
        while True:            
            tela['Info'].update(visible=True)

            eventos, valores = tela.read(timeout=0.1)
            caminho = valores['arquivo']

            arquivo_selecionado = caminho.split('/')
            nome_arquivo  = arquivo_selecionado[-1][:-4]
            formato       = arquivo_selecionado[-1].split('.')
            tipo_arquivo  = formato[-1]
            
            if tipo_arquivo == 'pdf':                  
                arquivo_tmp = converterPDF(caminho)
                print(str(arquivo_tmp))                              
            else:
                arquivo_tmp = caminho

            if identificaVersao(arquivo_tmp) == 'cresol':
                dados = cresol.importar_cabecalho(arquivo_tmp, True)                
            else:
                dados = sicredi.importaFichaGrafica(arquivo_tmp, True)                

            for l in dados:
                informacoes = informacoes + ' ' + str(l) + '\n'

            tela['ed_informacoes'].update(informacoes)                
            break

    def carregaParametros(tela):
        parametros = f.carregaParametrosGerais()

        if parametros != None:
            tela['ed_igpm'].update(parametros[0])
            tela['ed_igpm_param'].update(parametros[0])
            tela['ed_ipca'].update(parametros[1])
            tela['ed_ipca_param'].update(parametros[1])
            tela['ed_cdi'].update(parametros[2])
            tela['ed_cdi_param'].update(parametros[2])
            tela['ed_inpc'].update(parametros[3])
            tela['ed_inpc_param'].update(parametros[3])
            tela['ed_tr'].update(parametros[4])
            tela['ed_tr_param'].update(parametros[4])

            tela['ed_multa_perc'].update(parametros[5])
            tela['ed_multa_perc_param'].update(parametros[5])
            tela['ed_multa_valor'].update(parametros[6])
            tela['ed_multa_valor_param'].update(parametros[6])
            tela['ed_multa_incidencia'].update(parametros[7])
            tela['ed_multa_incidencia_param'].update(parametros[7])

            tela['ed_honorarios_perc'].update(parametros[8])
            tela['ed_honorarios_perc_param'].update(parametros[8])
            tela['ed_honorarios_valor'].update(parametros[9])
            tela['ed_honorarios_valor_param'].update(parametros[9])

            tela['ed_outros_valor'].update(parametros[10])
            tela['ed_outros_valor_param'].update(parametros[10])

            print('Carregou os parâmetros padrão.')
        return True
        
    
    busca_licena = verificaLicenca() 
    if (busca_licena == 'THMPV-77D6F-94376-8HGKG-VRDRQ'):                
        tela = sg.Window('Corretor', tabgrp)
        progress_bar = tela['progressbar']

        carregou_parametros = False

        while True:                
            eventos, valores = tela.read(timeout=0.1)
            
            if carregou_parametros == False:
                carregou_parametros  = carregaParametros(tela)

            if eventos is None or eventos == "Fechar":
                break
            
            tela['ed_situacao'].update("Aguardando Operação")
            tela['ed_situacao'].update(text_color="Green")
            progress_bar.update(visible=False)
            progress_bar.UpdateBar(0)
            
            #if eventos == '-abas-':        
            #    print(valores)

            if '+CLICKED+' in eventos:
                print(valores['-TABLE-']) 
                #sg.popup("You clicked row:{} Column: {}".format(eventos[2][0], eventos[2][1]))

            if eventos == '-INPUT-':
                atualizaInfo()     
                
            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_multa_perc': 
                if valores['ed_multa_perc'] != '':          
                    if valores['ed_multa_perc'][-1] not in ('0123456789,.'):                                                
                        tela['ed_multa_perc'].update(valores['ed_multa_perc'][:-1])
            
            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_multa_valor': 
                if valores['ed_multa_valor'] != '':          
                    if valores['ed_multa_valor'][-1] not in ('0123456789,.'):                                                
                        tela['ed_multa_valor'].update(valores['ed_multa_valor'][:-1]) 
                        
            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_honorarios_perc': 
                if valores['ed_honorarios_perc'] != '':          
                    if valores['ed_honorarios_perc'][-1] not in ('0123456789,.'):                                                
                        tela['ed_honorarios_perc'].update(valores['ed_honorarios_perc'][:-1])

            ## Tratamento para campos de valor não aceitar letras            
            if eventos == 'ed_honorarios_valor': 
                if valores['ed_honorarios_valor'] != '':          
                    if valores['ed_honorarios_valor'][-1] not in ('0123456789,.'):                                                
                        tela['ed_honorarios_valor'].update(valores['ed_honorarios_valor'][:-1])
                        
            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_outros_valor': 
                if valores['ed_outros_valor'] != '':          
                    if valores['ed_outros_valor'][-1] not in ('0123456789,.'):                                                
                        tela['ed_outros_valor'].update(valores['ed_outros_valor'][:-1])

            if eventos == 'btn_atualizar_importadas':                                                 
                db     = f.conexao()
                cursor = db.cursor()
                
                if valores['ed_filtro_titulo'] != '':
                    v_titulo_filtro = ' AND titulo like "%' + valores['ed_filtro_titulo'] + '%" '
                else:
                    v_titulo_filtro = ''
                    
                if valores['ed_filtro_associado'] != '':
                    v_associado_filtro = ' AND associado like "%' + valores['ed_filtro_associado'] + '%" '
                else:
                    v_associado_filtro = ''  
                
                vfiltros = v_titulo_filtro + ' ' + v_associado_filtro
                
                sql_consulta = 'SELECT f.id,f.titulo,f.versao,f.associado,f.nro_parcelas,f.valor_financiado,f.tx_juro FROM ficha_grafica as f WHERE f.situacao="ATIVO" ' + vfiltros + ' order by id DESC'
                
                cursor.execute(sql_consulta)
                dados = cursor.fetchall()
                
                fichas_importadas = []
                for ficha in dados:                                        
                    fichas_importadas.append(ficha)                    
                    
                tela['-TABLE-'].update(values = fichas_importadas)    
                
                cursor.close()

            if eventos == 'btn_salvar_parametros_gerais':
                print('Salvar parâmetros gerais')

                vincidencia_param = valores['ed_multa_incidencia_param']

                ## Trata campos dos parâmetros
                if valores['ed_multa_perc_param'] == '':
                    vmulta_perc_param = 0
                else:
                    vmulta_perc_param = valores['ed_multa_perc_param'].replace(',','.')
                    vmulta_perc_param = float(vmulta_perc_param)

                if valores['ed_multa_valor_param'] == '':
                    vmulta_valor_param = 0
                else:
                    vmulta_valor_param = valores['ed_multa_valor_param'].replace(',','.')
                    vmulta_valor_param = float(vmulta_valor_param)

                if valores['ed_honorarios_perc_param'] == '':
                    vhonorarios_perc_param = 0
                else:
                    vhonorarios_perc_param = valores['ed_honorarios_perc_param'].replace(',','.')
                    vhonorarios_perc_param = float(vhonorarios_perc_param)

                if valores['ed_honorarios_valor_param'] == '':
                    vhonorarios_valor_param = 0
                else:
                    vhonorarios_valor_param = valores['ed_honorarios_valor_param'].replace(',','.')
                    vhonorarios_valor_param = float(vhonorarios_valor_param)
                
                if valores['ed_outros_valor_param'] == '':
                    voutros_valor_param = 0
                else:
                    voutros_valor_param = valores['ed_outros_valor_param'].replace(',','.')
                    voutros_valor_param = float(voutros_valor_param)

                try:
                   if f.salvarParametrosGerais(valores['ed_igpm_param'], valores['ed_ipca_param'], valores['ed_cdi_param'], valores['ed_inpc_param'], valores['ed_tr_param'], vmulta_perc_param, vmulta_valor_param, vincidencia_param, vhonorarios_perc_param, vhonorarios_valor_param, voutros_valor_param):
                       carregaParametros(tela)
                       sg.popup('Parâmetros definidos com sucesso!')
                except Exception as erro:
                    sg.popup('Ocorreu um erro ao tentar salvar. ' + str(erro))                    
            
            if eventos == 'Calcular':
                ## Aqui colocar validações dos campos

                vincidencia = valores['ed_multa_incidencia']

                ## Trata campos dos parâmetros
                if valores['ed_multa_perc'] == '':
                    vmulta_perc = 0
                else:
                    vmulta_perc = valores['ed_multa_perc'].replace(',','.')
                    vmulta_perc = float(vmulta_perc)

                if valores['ed_multa_valor'] == '':
                    vmulta_valor = 0
                else:
                    vmulta_valor = valores['ed_multa_valor'].replace(',','.')
                    vmulta_valor = float(vmulta_valor)

                if valores['ed_honorarios_perc'] == '':
                    vhonorarios_perc = 0
                else:
                    vhonorarios_perc = valores['ed_honorarios_perc'].replace(',','.')
                    vhonorarios_perc = float(vhonorarios_perc)

                if valores['ed_honorarios_valor'] == '':
                    vhonorarios_valor = 0
                else:
                    vhonorarios_valor = valores['ed_honorarios_valor'].replace(',','.')
                    vhonorarios_valor = float(vhonorarios_valor)
                
                if valores['ed_outros_valor'] == '':
                    voutros_valor = 0
                else:
                    voutros_valor = valores['ed_outros_valor'].replace(',','.')
                    voutros_valor = float(voutros_valor)
                
                parametros = {
                    'igpm'             : valores['ed_igpm'],
                    'ipca'             : valores['ed_ipca'],
                    'cdi'              : valores['ed_cdi'],
                    'inpc'             : valores['ed_inpc'],
                    'tr'               : valores['ed_tr'],
                    'multa_perc'       : vmulta_perc,
                    'multa_valor'      : vmulta_valor,
                    'multa_incidencia' : vincidencia,
                    'honorarios_perc'  : vhonorarios_perc,
                    'honorarios_valor' : vhonorarios_valor,
                    'outros_valor'     : voutros_valor
                }    
                
                progress_bar.update(visible=True)
                progress_bar.UpdateBar(50)

                tela['ed_situacao'].update("Processando Arquivo... Aguarde...")
                tela['ed_situacao'].update(text_color="Blue")
                eventos, valores = tela.read(timeout=1)
                remove_arquivo = False

                if valores['arquivo'] != '':
                    progress_bar.UpdateBar(10)
                    
                    caminho = valores['arquivo']

                    arquivo_selecionado = caminho.split('/')                

                    nome_arquivo  = arquivo_selecionado[-1][:-4]
                    formato       = arquivo_selecionado[-1].split('.')
                    tipo_arquivo  = formato[-1]                

                    path_destino  = vPath + '/Processados/' + nome_arquivo

                    progress_bar.UpdateBar(27)
                    if not os.path.isdir(path_destino): 
                        os.mkdir(path_destino)

                    if tipo_arquivo == 'pdf':                    
                        arquivo_importar = converterPDF(caminho)
                        remove_arquivo   = True                    
                    else:
                        arquivo_importar = caminho

                    progress_bar.UpdateBar(34)

                    if arquivo_importar != None:
                        
                        progress_bar.UpdateBar(50)
                        if identificaVersao(arquivo_importar) == 'sicredi':
                            versao = 'sicredi'
                            titulo = sicredi.importarSicredi(arquivo_importar, parametros)    
                            progress_bar.UpdateBar(74)                
                        else:
                            versao = 'cresol'
                            titulo = cresol.importarCresol(arquivo_importar, parametros)
                            progress_bar.UpdateBar(74)

                        db     = f.conexao()
                        cursor = db.cursor()
                        tipo = 'Correcao_Comum'     
                        
                        ## Aqui Iniciaria a busca dos dados do banco e o calculo, adicionando as linhas calculadas para a confecção do relatório               
                        
                        if versao == 'sicredi':                        
                            ##Busca dados do cabeçalho no BD                            
                            sql_consulta = 'select titulo,associado,nro_parcelas,parcela,valor_financiado,tx_juro,multa,liberacao from ficha_grafica WHERE titulo = %s AND situacao = "ATIVO"'
                            cursor.execute(sql_consulta, (titulo,))                                                                        
                            dados_cabecalho = cursor.fetchall()
                                                                                                                                                                                                                              
                            ##busca detalhes do BD
                            ##Alimenta variaveis para relatorio
                        else:
                            sql_consulta = 'select titulo,associado,modalidade_amortizacao,nro_parcelas,parcela,valor_financiado,tx_juro,multa,liberacao from ficha_grafica WHERE titulo = %s AND situacao = "ATIVO"'
                            cursor.execute(sql_consulta, (titulo,))                                                                        
                            dados_cabecalho = cursor.fetchall()                                        
                            
                        context = {
                                      "nome_associado" : dados_cabecalho[0][1],
                                      "tipo_correcao"  : tipo,
                                      "numero_titulo"  : dados_cabecalho[0][0],
                                      "forma_calculo"  : "Parcelas Atualizadas Individualmente De 27/09/2013 a 11/04/2023 sem correção Multa de 2,0000 sobre o valor corrigido+juros principais+juros moratórios",
                                      "forma_juros"    : "Juros ok",
                                      "lancamentos"    : lancamentos
                                  }
                        progress_bar.UpdateBar(86)
                        ## Gera o relatório e transforma em .pdf
                        template = DocxTemplate('C:/Temp/Fichas_Graficas/Template.docx')                    
                        template.render(context)
                        template.save(path_destino + '/' + tipo + '.docx')
                        convert(path_destino + '/' + tipo + '.docx', path_destino + '/' + tipo + '.pdf')
                        os.remove(path_destino + '/' + tipo + '.docx')
                        
                        progress_bar.UpdateBar(96)                                                
                    try:
                        if remove_arquivo:
                            os.remove(arquivo_importar)
                    except:
                        sg.popup('Falha ao tentar remover o arquivo txt temporário.')    

                    progress_bar.UpdateBar(100)
                    sg.popup('Processo Finalizado')
                    
            if eventos == sg.WINDOW_CLOSED:
                break
            if eventos == 'Fechar':
                break                                                                                                                                      
                
main()

