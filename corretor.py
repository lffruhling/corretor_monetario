from datetime import datetime
import time
import os
import pdfplumber
import shutil

import funcoes   as f
import importar  as sicredi
import importarc as cresol

vPath  = 'C:/Temp/Fichas_Graficas'
versao = ''

def converterPDF(vPath, vNomeArquivo):    
    try:
        ## Carrega arquivo
        pdf = pdfplumber.open(vPath + '/' + vNomeArquivo + '.pdf')

        ## Converte PDF em TXT
        for pagina in pdf.pages:
            texto = pagina.extract_text(x_tolerance=1)
            
            with open(vPath + '/' + vNomeArquivo + '.txt', 'a') as arquivo_txt:
                arquivo_txt.write(str(texto))       
                
        pdf.close()         

        return True
    
    except Exception as erro:
        with open(vPath + '/ERRO_LOG.txt', 'a') as arquivo_txt:
            arquivo_txt.write(str(erro))
    
        return False
    
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
    vEncontrouArquivos = False
    ## Verifica se existe path em C:/Temp/Fichas Gráficas senão, criar
    ## Verifica dentro da pasta se existe e criar as pastas processados
    ## Verifica dentro da pasta se existe algum arquivo PDF ou txt e pegar o nome dele automaticamente
    ## Usar esse arquivo para realizar a conversão/importação, criando uma pasta para o titulo em processados, onde ficará o resultado.

    if not os.path.isdir(vPath): # vemos de este diretorio ja existe        
        os.mkdir(vPath)
        os.mkdir(vPath + '/Processados')        

    for arquivo in os.walk(vPath):
        if str(arquivo[0]) == vPath:
            for ficha in arquivo[2]:
                arquivo_capturado = ficha.split('.')
                nome_arquivo      = arquivo_capturado[0]
                tipo_arquivo      = arquivo_capturado[1].replace(" ","")
                
                if not os.path.isdir(vPath + '/Processados/' + nome_arquivo): 
                    os.mkdir(vPath + '/Processados/' + nome_arquivo)

                if tipo_arquivo == 'PRN':
                    vCaminhoTxt = vPath + '/' + nome_arquivo + '.PRN'    
                else:
                    vCaminhoTxt = vPath + '/' + nome_arquivo + '.txt'                    

                if not os.path.isdir(vPath + '/Processados/' + nome_arquivo): # vemos de este diretorio ja existe        
                    os.mkdir(vPath + '/Processados/' + nome_arquivo)           

                if tipo_arquivo in ('pdf', 'txt', 'PRN'):
                    vEncontrouArquivos = True
                    if tipo_arquivo == 'pdf':
                        if converterPDF(vPath, nome_arquivo):
                            shutil.move(vPath + '/' + nome_arquivo + '.pdf', vPath + '/Processados/' + nome_arquivo + '/' + nome_arquivo + '.pdf')
                    
                    if identificaVersao(vCaminhoTxt) == 'sicredi':
                        sicredi.importarSicredi(vCaminhoTxt)                    
                    else:
                        cresol.importarCresol(vCaminhoTxt)
                    
                    ## Aqui precisamos identificar o titulo para utilizar no select para buscar as informações e realizar os calculos
                        
                    shutil.move(vCaminhoTxt, vPath + '/Processados/' + nome_arquivo + '/' + nome_arquivo + '.' + tipo_arquivo)    
        
    if not vEncontrouArquivos:
        print('Não existem arquivos para processar.')
        time.sleep(10)                                            
                
main()
