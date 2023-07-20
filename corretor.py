from datetime import datetime
import time
import os
import sys
import pdfplumber
import shutil
from datetime import datetime

import funcoes   as f
import importar  as sicredi
import importarc as cresol

#Relatorio
from docxtpl import DocxTemplate
from docx2pdf import convert

#Interface
import PySimpleGUI as sg

vPath    = 'C:/Temp/Fichas_Graficas'
versao   = ''
lancamentos       = []
lancamentosCDI    = []
fichas_importadas = []
fichas_alcadas    = []  
tela = None

id_ficha_grafica = 0

def atualizaStatus(texto_adicional):
    global tela
    tela['ed_situacao'].update("Processando Arquivo... Aguarde... " + 'Importando: ' + texto_adicional)

def alimentaDetalhesRelatorio(lista, data, descricao, valor, correcao, corrigido, juros, total):
    lista.append({"data":data, "descricao":descricao, "valor":valor,"correcao":correcao,"corrigido":corrigido,"juros":juros,"total":total})

def converterPDF(vCaminho_arquivo):    
    try:
        ## Carrega arquivo
        pdf = pdfplumber.open(vCaminho_arquivo)

        vNome        = vCaminho_arquivo.split('/')
        vNomeArquivo = vNome[-1][:-4]

        try:
            if os.path.isfile(vPath + '/' + vNomeArquivo + '.txt'):
                os.remove(vPath + '/' + vNomeArquivo + '.txt')
        except Exception as erro:
            print('Erro ao tentar remover arquivo txt temporário. ' + str(erro))

        print('Convertendo ' + str(len(pdf.pages)) + ' página PDF para txt')

        ## Converte PDF em TXT
        for pagina in pdf.pages:
            texto = pagina.extract_text(x_tolerance=1)

            vArquivo_final = vPath + '/' + vNomeArquivo + '.txt'
            
            with open(vArquivo_final, 'a') as arquivo_txt:
                arquivo_txt.write(str(texto))       

        pdf.close()         

        return vArquivo_final
    
    except Exception as erro:        
        f.gravalog('Falha ao tentar converter o PDF para txt. ' + str(erro), True)
    
        return None    
    
def telaAlcadas(parametrosGerais=False):       
    global fichas_alcadas 

    if (parametrosGerais): 
        dados = []       
        retorno = f.carregarParametrosAlcadas()

        for item in retorno:
            dados.append(item)          
    else:        
        dados = fichas_alcadas

    sg.theme('Reddit')
    
    alcadas = [    
                [sg.Text(text='Definição de Alçadas', text_color="Black", font=("Arial",12, "bold"), expand_x=True, justification='center')],  
                [sg.Table(values=dados, headings=['Alçada', 'Perc. %'], auto_size_columns=True, display_row_numbers=False, justification='center', key='-TABLE_ALCADAS-', selected_row_colors='red on yellow', enable_events=True, expand_x=True, expand_y=True,enable_click_events=True)],
                [sg.Text(text='Alçada %: '),sg.Input(key='ed_valor_alcada', size=(15,1), enable_events=True),sg.Button(' + Adicionar', key='btn_incluir_alcadas'), sg.Button(' - Remover', key='btn_remover_alcadas')],
                [sg.Button('Salvar', key='btn_salvar_alcadas'), sg.Button('Cancelar', key='btn_cancelar_alcadas')]      
              ]
    
    tela = sg.Window('Alçadas', alcadas)                

    while True:                
        eventos, valores = tela.read(timeout=0.1)

        if eventos == 'ed_valor_alcada': 
            if valores['ed_valor_alcada'] != '':                  
                if valores['ed_valor_alcada'][-1] not in ('0123456789,.'):                                                
                    tela['ed_valor_alcada'].update(valores['ed_valor_alcada'][:-1])

        if eventos == 'btn_salvar_alcadas':            
            fichas_alcadas = dados
            tela.close()

            ## Salvar parametros no banco
            if parametrosGerais:
                f.salvaParametrosAlcadas(dados)
            else:
                fichas_alcadas = dados

            return None        

        if eventos is None or eventos == "Cancelar":            
            tela.close() 

            if not(parametrosGerais):
                fichas_alcadas = dados

            return None           

        if eventos == 'btn_incluir_alcadas':
            i = 1            
            for alcadas in dados:            
                i = i + 1

            if valores['ed_valor_alcada'] == '':
                sg.popup('É necessário informar um percentual')
            else:

                if float(valores['ed_valor_alcada'].replace(",",".")) > 0:                    
                    dados.append([i,valores['ed_valor_alcada'].replace(",",".")])                    
                    tela['ed_valor_alcada'].update('')                    
                    tela['-TABLE_ALCADAS-'].update(values=dados)

                    if not(parametrosGerais):
                        fichas_alcadas = dados
        
        if eventos == 'btn_remover_alcadas':                        
            if valores['-TABLE_ALCADAS-'] != []:
                dados.pop(valores['-TABLE_ALCADAS-'][0])

                ## Atualiza Numeração das alçadas
                i = 1
                x = 0                
                for alcadas in dados:                    
                    dados[x] = [i, alcadas[1]]
                    i = i + 1
                    x = x + 1
                
                tela['-TABLE_ALCADAS-'].update(values=dados)
                
                if not(parametrosGerais):
                    fichas_alcadas = dados            

def main():
    nome_arquivo       = ''
    tipo_arquivo       = ''
    versao             = '' 
    informacoes        = '' 
    arquivo_importar   = ''
    remove_arquivo     = False    
    
    ip                 = f.pegarIp()         
    internet           = f.testarInternet()
    maquina            = f.pegarNomeMaquina() 
    ultima_versao      = f.BuscaUltimaVersao()
    
    versao_exe         = '1.0.0' 
    
    versao_descricao   = 'Atualizado'
    if(versao_exe != ultima_versao):
        versao_descricao = 'Atualização Disponível'

    global tela
    global fichas_importadas
    global fichas_alcadas
    global id_ficha_grafica

    if not os.path.isdir(vPath): # vemos de este diretorio ja existe        
        os.mkdir(vPath)
        os.mkdir(vPath + '/Processados')  
    
    ## Monta o Layout da tela
    sg.theme('Reddit')
    
    lista_arquivo = []
    lista_arquivo.append(sg.Input(key='-INPUT-', enable_events=True))
    lista_arquivo.append(sg.FileBrowse(button_text='Procurar', enable_events=True, key='arquivo'))
    lista_arquivo.append(sg.Text("Vazio", font=(15,15), key="txt_titulo"))
    lista_arquivo.append(sg.Button("Cancelar", key='btn_cancelar', visible=False))
    frame_arquivo = sg.Frame('Selecione a Ficha Gráfica', [lista_arquivo], expand_x=True)
    
    lista_indices = []
    lista_indices.append(sg.Checkbox(text='IGPM', key='ed_igpm'))
    lista_indices.append(sg.Checkbox(text='IPCA', key='ed_ipca'))
    lista_indices.append(sg.Checkbox(text='CDI', key='ed_cdi'))
    lista_indices.append(sg.Checkbox(text='INPC', key='ed_inpc'))
    lista_indices.append(sg.Checkbox(text='TR', key='ed_tr'))
    lista_indices.append(sg.Button("Alçadas", key='btn_alcadas'))
    frame_indices = sg.Frame('Índices de Correção', [lista_indices], expand_x=True)

    lista_informacoes = []
    lista_informacoes.append(sg.Multiline(default_text=informacoes, expand_x=True, size=(None, 5), key="ed_informacoes", enable_events=True, auto_refresh=True) )    
    frame_informacoes = sg.Frame('Informações da Ficha Gráfica Selecionada', [lista_informacoes], expand_x=True, visible=True, key='Info')      
    
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
    lista_outros.append(sg.Input(key="ed_outros_valor", size=(20,1), default_text='0,00', enable_events=True ))        
    frame_outros = sg.Frame('Outros Valores', [lista_outros], expand_x=True, key='frame_outros')    
    
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
    
    ## Itens da aba Principal
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
                [sg.Button('Calcular'), sg.Button('Fechar'), sg.Button('Testar'), sg.Text('Host: ' + maquina + '   |   IP: '+ ip +'   |   '+ internet +'   |   Versão: ' + versao_exe + ' - ' + versao_descricao, expand_x=True, justification='right', font=("Verdana",8))]      
             ]
    
    ## Itens da aba Importadas
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

    ## Itens da aba Parâmetros
    parametros = [    
                    [sg.Text(text='Parâmetros', text_color="black", font=("Arial",12), expand_x=True, justification='center')],                                        
                    [sg.Text(text='Parâmetros Padrão do Sistema', text_color="black", font=("Arial",9), expand_x=True, justification='left')],
                    [sg.Text(text='Atenção: Opções definidas aqui serão carregadas como padrão para todos ao inicializar uma nova importação.', text_color="red", font=("Verdana",7), expand_x=True, justification='left')],
                    [frame_indices_parametros],
                    [frame_multa_parametros],
                    [frame_honorarios_parametros],
                    [frame_outros_parametros],
                    [sg.Button('Definir Alçadas Padrão', key="btn_alcadas_padrao")],
                    [sg.Button('Salvar Parâmetros', key='btn_salvar_parametros_gerais'), sg.Button('Cancelar')]
                 ]

    ## Cria as abas
    tabgrp = [
                [sg.TabGroup
                    (
                        [
                            [sg.Tab('Principal', principal, title_color='Red',border_width =10,key='aba_principal'), sg.Tab('Importadas', importadas, key='aba_importadas'), sg.Tab('Parâmetros', parametros, key='aba_parametros')]                            
                        ], enable_events=True, key='-abas-'
                    )
                ]
             ]  
    

    ## Verifica se tem licença inserida no registro do windows
    busca_licena = f.verificaLicenca() 
    if ((busca_licena == None) or (busca_licena != 'THMPV-77D6F-94376-8HGKG-VRDRQ')):
        f.licenca()

    def atualizaTxtTitulo(texto, green=False):
        tela['txt_titulo'].update(texto)

        if green:
            tela['txt_titulo'].update(text_color="green")
            tela['btn_cancelar'].update(visible=True)
        else:
            tela['txt_titulo'].update(text_color="red")
            tela['btn_cancelar'].update(visible=True)

        if texto == "Vazio":
            tela['txt_titulo'].update(text_color="black")
            tela['ed_informacoes'].update("")
            tela['btn_cancelar'].update(visible=False)

    def atualizaInfo():
        ## Função para atualizar as informações do arquivo selecionado, que são exibidas no MEMO
        global informacoes
        informacoes = ''
        while True:                        
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

            if f.identificaVersao(arquivo_tmp) == 'cresol':
                ## Utiliza a mesma função de importar, porém, definido flag para True
                dados = cresol.importar_cabecalho(arquivo_tmp, True)                
            else:
                ## Utiliza a mesma função de importar, porém, definido flag para True
                dados = sicredi.importaFichaGrafica(arquivo_tmp, True)                

            for l in dados:
                informacoes = informacoes + ' ' + str(l) + '\n'

                if "Título" in str(l):
                    titulo = str(l).replace("Título:","")
                    atualizaTxtTitulo(titulo)

            tela['ed_informacoes'].update(informacoes)                
            break

    def carregaParametros(tela):
        ## Carrega os valores dos parametros do BD e alimenta os campos na tela

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
        
    def carregaAlcadasPadrao():
        global fichas_alcadas

        fichas_alcadas = []
        alcadas = f.carregarParametrosAlcadas()
        for item in alcadas:
            fichas_alcadas.append(item)
   
    ## Primeiro verifica a licença, depois abre tela do sistema
    busca_licena = f.verificaLicenca() 
    if (busca_licena == 'THMPV-77D6F-94376-8HGKG-VRDRQ'):                
        f.gravalog('Iniciado - Licença Ok')
        #Inicia o corretor
        tela         = sg.Window('Corretor', tabgrp, size=(None,None))
        progress_bar = tela['progressbar']

        carregou_parametros = False

        while True:                
            eventos, valores = tela.read(timeout=0.1)
            
            ## Carrega os parâmetros e alçadas padrão
            if carregou_parametros == False:
                carregou_parametros  = carregaParametros(tela)                
                carregaAlcadasPadrao()

            if eventos is None or eventos == "Fechar":
                break
            
            tela['ed_situacao'].update("Aguardando Operação")
            tela['ed_situacao'].update(text_color="Green")
            progress_bar.update(visible=False)
            progress_bar.UpdateBar(0)                        

            ## Chamar tela das alçadas
            if eventos == 'btn_alcadas':                               
                telaAlcadas()

            ## Chamar tela alçadas padrão(parâmetros)
            if eventos == 'btn_alcadas_padrao':                               
                telaAlcadas(True)

            if eventos == 'btn_cancelar':
                id_ficha_grafica = 0
                tela['-INPUT-'].update("")
                carregaParametros(tela)
                carregaAlcadasPadrao()
                atualizaTxtTitulo('Vazio')

            ## Clique duplo sobre item da lista
            if '+CLICKED+' in eventos:                
                try:
                    id_ficha_grafica = fichas_importadas[valores['-TABLE-'][0]][0]

                    ## Se for uma ficha já importada, carrega os valores utilizados nessa ficha
                    if id_ficha_grafica > 0:
                        parametros_ficha = f.carregaParametrosFichaGrafica(id_ficha_grafica)
                        dados_ficha      = f.carregaDadosCabecalhoFichaGrafica(id_ficha_grafica)
                        
                        alcadas_ficha    = f.carregarFichaAlcadas(id_ficha_grafica)
                        fichas_alcadas   = []
                        for item in alcadas_ficha:
                            fichas_alcadas.append(item)
                        #print(dados_ficha)

                        informacoes = ''
                        informacoes = informacoes + 'Coop: ' + dados_ficha[1] + '   ' + 'Modalidade: ' + dados_ficha[9]
                        informacoes = informacoes + '\n' + 'Associado: ' + dados_ficha[2] + '\n'
                        informacoes = informacoes + 'Data de Liberação: ' + dados_ficha[8].strftime("%d/%m/%Y") + '\n'
                        informacoes = informacoes + 'Número de Parcelas: ' + str(dados_ficha[3]) + '\n'
                        informacoes = informacoes + 'Parcela Atual: ' + str(dados_ficha[4]) + '\n'
                        informacoes = informacoes + 'Título: ' + dados_ficha[0] + '\n'
                        informacoes = informacoes + 'Taxa de Juros: ' + str(dados_ficha[6]) + '\n'
                        informacoes = informacoes + 'Valor Financiado: ' + str(dados_ficha[5])
                        tela['ed_informacoes'].update(informacoes)

                        if dados_ficha != []:
                            tela['ed_igpm'].update(parametros_ficha[0])
                            tela['ed_ipca'].update(parametros_ficha[1])
                            tela['ed_cdi'].update(parametros_ficha[2])
                            tela['ed_inpc'].update(parametros_ficha[3])
                            tela['ed_tr'].update(parametros_ficha[4])

                            tela['ed_multa_perc'].update(parametros_ficha[5])
                            tela['ed_multa_valor'].update(parametros_ficha[6])
                            tela['ed_multa_incidencia'].update(parametros_ficha[7])
                            
                            tela['ed_honorarios_perc'].update(parametros_ficha[8])
                            tela['ed_honorarios_valor'].update(parametros_ficha[9])
                            tela['ed_outros_valor'].update(parametros_ficha[10])

                            tela['-INPUT-'].update("")

                        atualizaTxtTitulo(str(dados_ficha[0]), True)
                        tela['aba_principal'].select()
                except:
                    print('Nenhum id selecionado')

            if eventos == '-INPUT-':
                id_ficha_grafica = 0
                carregaParametros(tela)
                atualizaInfo()     
                
            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_multa_perc': 
                if valores['ed_multa_perc'] != '':          
                    if valores['ed_multa_perc'][-1] not in ('0123456789,.'):                                                
                        tela['ed_multa_perc'].update(valores['ed_multa_perc'][:-1])

            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_multa_perc_param': 
                if valores['ed_multa_perc_param'] != '':          
                    if valores['ed_multa_perc_param'][-1] not in ('0123456789,.'):                                                
                        tela['ed_multa_perc_param'].update(valores['ed_multa_perc_param'][:-1])
            
            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_multa_valor': 
                if valores['ed_multa_valor'] != '':          
                    if valores['ed_multa_valor'][-1] not in ('0123456789,.'):                                                
                        tela['ed_multa_valor'].update(valores['ed_multa_valor'][:-1]) 

            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_multa_valor_param': 
                if valores['ed_multa_valor_param'] != '':          
                    if valores['ed_multa_valor_param'][-1] not in ('0123456789,.'):                                                
                        tela['ed_multa_valor_param'].update(valores['ed_multa_valor_param'][:-1])
                        
            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_honorarios_perc': 
                if valores['ed_honorarios_perc'] != '':          
                    if valores['ed_honorarios_perc'][-1] not in ('0123456789,.'):                                                
                        tela['ed_honorarios_perc'].update(valores['ed_honorarios_perc'][:-1])

            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_honorarios_perc_param': 
                if valores['ed_honorarios_perc_param'] != '':          
                    if valores['ed_honorarios_perc_param'][-1] not in ('0123456789,.'):                                                
                        tela['ed_honorarios_perc_param'].update(valores['ed_honorarios_perc_param'][:-1])

            ## Tratamento para campos de valor não aceitar letras            
            if eventos == 'ed_honorarios_valor': 
                if valores['ed_honorarios_valor'] != '':          
                    if valores['ed_honorarios_valor'][-1] not in ('0123456789,.'):                                                
                        tela['ed_honorarios_valor'].update(valores['ed_honorarios_valor'][:-1])

            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_honorarios_valor_param': 
                if valores['ed_honorarios_valor_param'] != '':          
                    if valores['ed_honorarios_valor_param'][-1] not in ('0123456789,.'):                                                
                        tela['ed_honorarios_valor_param'].update(valores['ed_honorarios_valor_param'][:-1])
                        
            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_outros_valor': 
                if valores['ed_outros_valor'] != '':          
                    if valores['ed_outros_valor'][-1] not in ('0123456789,.'):                                                
                        tela['ed_outros_valor'].update(valores['ed_outros_valor'][:-1])

            ## Tratamento para campos de valor não aceitar letras
            if eventos == 'ed_outros_valor_param': 
                if valores['ed_outros_valor_param'] != '':          
                    if valores['ed_outros_valor_param'][-1] not in ('0123456789,.'):                                                
                        tela['ed_outros_valor_param'].update(valores['ed_outros_valor_param'][:-1])

            if eventos == 'btn_atualizar_importadas':                                                 
                db     = f.conexao()
                cursor = db.cursor()
                
                ## Filtro Titulo
                if valores['ed_filtro_titulo'] != '':
                    v_titulo_filtro = ' AND titulo like "%' + valores['ed_filtro_titulo'] + '%" '
                else:
                    v_titulo_filtro = ''
                    
                ## Filtro Associado
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
                    
                ## Atualiza valores da lista
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
                    f.gravalog('Falha ao tentar salvar os parâmetros. ' + str(erro), True)                
            
            if eventos == 'Calcular':

                ## Se tentar calcular uma ficha já importada
                if id_ficha_grafica > 0:
                    sg.popup('Esta ficha já foi calculada. Para calcular novamente, selecione o arquivo.')
                    continue

                ## ALÇADAS

                ## As alçadas definidas e que serão utilizadas no cálculo da ficha estão carregadas em memória
                ## ou seja, estão no array global "fichas_alcadas", entao para sabermos se existe e quais são as alçadas
                ## terá que ser feito um laço no array. 
                # fichas_alcadas    


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
                
                ## Armazena os valores dos parâmetros utilizados para depois salvar no bd para ter o histórico
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
                
                ## Atualiza interface
                progress_bar.update(visible=True)
                progress_bar.UpdateBar(50)                
                tela['ed_situacao'].update("Processando Arquivo... Aguarde...")
                tela['ed_situacao'].update(text_color="Blue")
                eventos, valores = tela.read(timeout=1)
                                
                ## Quando é um pdf, converte para para txt, no final exclui o txt
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

                    is_pdf = False
                    if tipo_arquivo == 'pdf':                    
                        arquivo_importar = converterPDF(caminho)
                        is_pdf           = True
                        remove_arquivo   = True                    
                    else:
                        arquivo_importar = caminho

                    progress_bar.UpdateBar(34)

                    if arquivo_importar != None:
                        
                        progress_bar.UpdateBar(50)
                        if f.identificaVersao(arquivo_importar) == 'sicredi':
                            versao   = 'sicredi'
                            retorno  = sicredi.importarSicredi(arquivo_importar, parametros, is_pdf, tela)
                            titulo   = retorno[0]
                            ficha_id = retorno[1]
                            progress_bar.UpdateBar(74)                
                        else:
                            versao   = 'cresol'
                            retorno  = cresol.importarCresol(arquivo_importar, parametros, tela)
                            titulo   = retorno[0]
                            ficha_id = retorno[1]
                            progress_bar.UpdateBar(74)

                        f.salvaFichaAlcadas(ficha_id, fichas_alcadas)

                        db     = f.conexao()
                        cursor = db.cursor()
                        tipo = 'Correcao_Comum'

                        totalJurosAcumulado     = 0
                        totalMorasAcumulado     = 0
                        totalParcelasAcumuladas = 0

                        adicionalMulta = float(parametros['multa_valor'])
                        adicionalHonorarios = float(parametros['honorarios_valor'])
                        adicionalOutrosValores = float(parametros['outros_valor'])

                        ## Aqui Iniciaria a busca dos dados do banco e o calculo, adicionando as linhas calculadas para a confecção do relatório               
                        sql_consulta = 'select id, titulo, associado, modalidade_amortizacao, nro_parcelas, parcela, valor_financiado, tx_juro, multa, liberacao from ficha_grafica WHERE titulo = %s AND situacao = "ATIVO"'
                        cursor.execute(sql_consulta, (titulo,))
                        dados_cabecalho = cursor.fetchone()

                        sql_parcelas = 'select distinct(data), case when valor_credito = 0 then valor_debito else valor_credito end as valor, historico from ficha_detalhe WHERE id_ficha_grafica = %s AND (historico like %s or historico like %s or historico like %s)  order by data'
                        cursor.execute(sql_parcelas, [dados_cabecalho[0], "%AMORTI%", "%LIQUI%", "%LIBERA%"])

                        resultParcelas = cursor.fetchall()

                        progress_bar.UpdateBar(75)

                        # Caclula Juros Simples
                        f.geraPDFCalculo(cursor=cursor,
                                         parametros=parametros,
                                         parcelas=resultParcelas,
                                         tx_juros=dados_cabecalho[7],
                                         multa=dados_cabecalho[8],
                                         nomeAssociado=dados_cabecalho[2],
                                         tipoCorrecao=tipo,
                                         nroTitulo=dados_cabecalho[1],
                                         dataLiberacao=dados_cabecalho[9].strftime('%d/%m/%Y'),
                                         path_destino=path_destino,
                                         adicionalMulta=adicionalMulta,
                                         adicionalHonorarios=adicionalHonorarios,
                                         adicionalOutrosValores=adicionalOutrosValores,
                                         alcadas=fichas_alcadas,
                                         isCresol=versao == 'cresol')

                        progress_bar.UpdateBar(80)
                        if parametros['cdi']:
                            progress_bar.UpdateBar(82)
                            f.geraPDFCalculo(cursor=cursor,
                                             parametros=parametros,
                                             parcelas=resultParcelas,
                                             tx_juros=dados_cabecalho[7],
                                             multa=dados_cabecalho[8],
                                             nomeAssociado=dados_cabecalho[2],
                                             tipoCorrecao=tipo,
                                             nroTitulo=dados_cabecalho[1],
                                             dataLiberacao=dados_cabecalho[9].strftime('%d/%m/%Y'),
                                             path_destino=path_destino,
                                             adicionalMulta=adicionalMulta,
                                             adicionalHonorarios=adicionalHonorarios,
                                             adicionalOutrosValores=adicionalOutrosValores,
                                             aplicaCDI=True,
                                             alcadas=fichas_alcadas,
                                             isCresol=versao == 'cresol')

                        progress_bar.UpdateBar(95)

                        path_final = os.path.realpath(path_destino)
                        os.startfile(path_final)
                        
                        progress_bar.UpdateBar(100)                                                
                    try:
                        if remove_arquivo:
                            os.remove(arquivo_importar)
                    except Exception as erro:
                        sg.popup('Falha ao tentar remover o arquivo txt temporário.')    
                        f.gravalog('Falha ao tentar remover o arquivo txt temporário(Conversão PDF para Txt). ' + str(erro))

                    progress_bar.UpdateBar(100)
                    sg.popup('Processo Finalizado')
                    
            if eventos == sg.WINDOW_CLOSED:
                break
            if eventos == 'Fechar':
                break                                                                                                                                   
                
main()

