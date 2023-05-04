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

vPath    = 'C:/Temp/Fichas_Graficas'
versao   = ''
lancamentos = []

def alimentaDetalhesRelatorio(lista, data, descricao, valor, correcao, corrigido, juros, total):
    lista.append({"data":data, "descricao":descricao, "valor":valor,"correcao":correcao,"corrigido":corrigido,"juros":juros,"total":total})

def janelaPrincipal():
    sg.theme('Reddit')

    layout = [    
        [sg.Text(text='Corretor Monetário')],      
        [sg.Text(text='Selecione a Ficha Gráfica', key='caminho_arquivo'), sg.FileBrowse(button_text='Procurar', key='arquivo')],    
        [sg.Checkbox(text='IGPM'), sg.Checkbox(text='IPCA'), sg.Checkbox(text='CDI')],    
        [sg.Button('Calcular'), sg.Button('Fechar')]      
    ]

    tela = sg.Window('Corretor', layout, size=(400,200))

    while True:                
        eventos, valores = tela.read()

        if eventos == 'Calcular':
            if valores['arquivo'] != '':
                caminho = valores['arquivo']
                sicredi.importarSicredi(caminho)

                sg.popup('Processo Finalizado')
                  
    
        if eventos == sg.WINDOW_CLOSED:
            break
        if eventos == 'Fechar':
            break

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
    
def main():            
    nome_arquivo       = ''
    tipo_arquivo       = ''
    versao             = '' 

    if not os.path.isdir(vPath): # vemos de este diretorio ja existe        
        os.mkdir(vPath)
        os.mkdir(vPath + '/Processados')   

    sg.theme('Reddit')

    layout = [    
                [sg.Text(text='Corretor Monetário')],      
                [sg.Text(text='Selecione a Ficha Gráfica', key='caminho_arquivo'), sg.FileBrowse(button_text='Procurar', key='arquivo')],    
                [sg.Checkbox(text='IGPM'), sg.Checkbox(text='IPCA'), sg.Checkbox(text='CDI')],    
                [sg.Text(text='Aguardando Operação', key='ed_situacao')],
                [sg.Button('Calcular'), sg.Button('Fechar')]      
             ]
    tela = sg.Window('Corretor', layout, size=(400,200))

    while True:                
        eventos, valores = tela.read()

        if eventos == 'Calcular':

            remove_arquivo = False

            if valores['arquivo'] != '':
                caminho = valores['arquivo']

                arquivo_selecionado = caminho.split('/')                

                nome_arquivo  = arquivo_selecionado[-1][:-4]
                formato       = arquivo_selecionado[-1].split('.')
                tipo_arquivo  = formato[-1]                

                path_destino  = vPath + '/Processados/' + nome_arquivo

                if not os.path.isdir(path_destino): 
                    os.mkdir(path_destino)

                if tipo_arquivo == 'pdf':                    
                    arquivo_importar = converterPDF(caminho)
                    remove_arquivo   = True
                else:
                    arquivo_importar = caminho

                if arquivo_importar != None:
                    if identificaVersao(arquivo_importar) == 'sicredi':
                        versao = 'sicredi'
                        titulo = sicredi.importarSicredi(arquivo_importar)                    
                    else:
                        versao = 'cresol'
                        titulo = cresol.importarCresol(arquivo_importar)

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

                    ## Gera o relatório e transforma em .pdf
                    template = DocxTemplate('C:/Temp/Fichas_Graficas/Template.docx')                    
                    template.render(context)
                    template.save(path_destino + '/' + tipo + '.docx')
                    convert(path_destino + '/' + tipo + '.docx', path_destino + '/' + tipo + '.pdf')
                    os.remove(path_destino + '/' + tipo + '.docx')
                    
                try:
                    if remove_arquivo:
                        os.remove(arquivo_importar)
                except:
                    sg.popup('Falha ao tentar remover o arquivo txt temporário.')    

                sg.popup('Processo Finalizado')
                
        if eventos == sg.WINDOW_CLOSED:
            break
        if eventos == 'Fechar':
            break                                                                                                                                      
                
main()

