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

    if not os.path.isdir(vPath): # vemos de este diretorio ja existe        
        os.mkdir(vPath)
        os.mkdir(vPath + '/Processados')   

    sg.theme('Reddit')

    layout = [    
                [sg.Text(text='                         Corretor Monetário', text_color="Silver", font=("Arial",15))],      
                [sg.Text(text='___________________________________________________', text_color="Silver", font=("Arial",15))],      
                [sg.Text(text='Selecione a Ficha Gráfica', key='caminho_arquivo'), sg.FileBrowse(button_text='Procurar', key='arquivo')],    
                [sg.Text(text='______________________________________________________________________________', text_color="Silver", font=("Arial",10))],      
                [sg.Checkbox(text='IGPM'), sg.Checkbox(text='IPCA'), sg.Checkbox(text='CDI')],    
                [sg.Text(text='Aguardando Operação', key='ed_situacao', text_color="green")],
                [sg.ProgressBar(100, orientation='h', size=(100, 20), key='progressbar', visible=False)],
                [sg.Button('Calcular'), sg.Button('Fechar')]      
             ]
    

    busca_licena = verificaLicenca() 
    if ((busca_licena == None) or (busca_licena != 'THMPV-77D6F-94376-8HGKG-VRDRQ')):
        licenca()

    busca_licena = verificaLicenca() 
    if (busca_licena == 'THMPV-77D6F-94376-8HGKG-VRDRQ'):
        tela = sg.Window('Corretor', layout, size=(500,270))
        progress_bar = tela['progressbar']


        while True:                    
            eventos, valores = tela.read(timeout=0.1)
            tela['ed_situacao'].update("Aguardando Operação")
            tela['ed_situacao'].update(text_color="Green")
            progress_bar.update(visible=False)
            progress_bar.UpdateBar(0)
            

            if eventos == 'Calcular':
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

                    progress_bar.UpdateBar(40)

                    if arquivo_importar != None:
                        
                        progress_bar.UpdateBar(50)
                        if identificaVersao(arquivo_importar) == 'sicredi':
                            versao = 'sicredi'
                            titulo = sicredi.importarSicredi(arquivo_importar)    
                            progress_bar.UpdateBar(74)                
                        else:
                            versao = 'cresol'
                            titulo = cresol.importarCresol(arquivo_importar)
                            progress_bar.UpdateBar(74)

                        db     = f.conexao()
                        cursor = db.cursor()
                        tipo = 'Correcao_Comum'                    
                        
                        if versao == 'sicredi':                        
                            ##Busca dados do cabeçalho no BD
                            
                            #sql_consulta = 'select titulo,associado,nro_parcelas,parcela,valor_financiado,tx_juro,multa,liberacao from ficha_grafica WHERE titulo = %s AND situacao = "ATIVO"'
                            #cursor.execute(sql_consulta, (titulo,))                                                                        
                            sql_consulta = 'select titulo,associado,modalidade_amortizacao,nro_parcelas,parcela,valor_financiado,tx_juro,multa,liberacao from ficha_grafica WHERE situacao = "ATIVO"'
                            cursor.execute(sql_consulta)                                                                        
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

