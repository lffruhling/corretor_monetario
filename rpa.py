import funcoes as f
import time
import sys 
import decimal
 
vLocal           = 'C:/Temp/ficha.PRN'
versao           = ''
parcela          = ''
parcelaInt       = 0
data_parcela     = ''   
valor_financiado = ''
amortizacoes     = []
multas           = []
juros            = []
total_multas     = 0
total_juros      = 0
total_pagamentos = 0
valor_saldo      = 0
vContinua        = False
vIdentificouData = False
linha_inadimplente = 0
 
def analisaArquivo():    
    global parcela
    global parcelaInt
    global data_parcela
    global versao
    global valor_financiado
    global amortizacoes
    global multas
    global total_pagamentos
    global total_multas
    global total_juros
    global valor_saldo
    global vContinua
    global vIdentificouData
    global linha_inadimplente
    
    with open(vLocal, 'r') as reader:
        ficha_grafica = reader.readlines()  
        
        ## Identifica a Versão do Arquivo(Se é emitido pelo Sicredi ou da Cresol)
        vlinha = 1        
        for linha in ficha_grafica:                                                    
            if (vlinha == 5):
                if ("RAIZES" in linha):
                    versao = 'raizes'   
                else:
                    versao = 'cresol'
                
                break
                
            vlinha = vlinha + 1

        ## Se for Ficha Gráfica Sicredi        
        if versao == 'raizes':
            vlinha = 1
            for linha in ficha_grafica:                                    
                
                #Captura a parcela em que chegou a Ficha gráfica
                if (vlinha == 6):
                    parcela = linha[116:134].replace(" ","")
                    parcela = parcela.split('/')
                    
                    if parcela != '':
                        parcelaInt = int(parcela[0])                
                        
                #Captura o valor Financiado
                if (vlinha == 7):                    
                    valor_financiado = linha[116:134].replace(" ","")
                                     
                # Lógica destes trechos:
                    # Testa se a linh atual possui o texto esperado;
                    # Pega texto da linha das posições(colunas) e retira os espaços em branco
                    # Adiciona o valor em uma lista de amortizações                     

                if (("AMORTIZACAO DE PARCELA" in linha)or("LIQUIDACAO DE PARCELA" in linha)):                                        
                    valor = linha[82:107].replace(" ","") 
                    amortizacoes.append(valor.replace(".",""))    

                if ("MULTA INADIMPLENTE" in linha):                    
                    valor = linha[63:82].replace(" ","")
                    multas.append(valor.replace(".",""))                   

                if ("061  JUROS INADIMPLENTE" in linha):                    
                    valor = linha[63:82].replace(" ","")
                    juros.append(valor.replace(".",""))

                if parcelaInt > 0:
                    if existeTextoLinha(linha, '0' + str(parcelaInt) + ')'):
                        data_parcela = pegaDataVencimentoParcela(linha, parcelaInt)
                        
                    if data_parcela != '':
                        if ((data_parcela in linha)):
                            vIdentificouData = True
                            valor_saldo = pegaSaldoLinha(linha)
                            linha_inadimplente = vlinha
                            
                    if (vIdentificouData and not(data_parcela in linha)):
                        vContinua = True
                        
                    if vContinua:
                        #print('Aqui identificou o saldo do último lançamento da data de vencimento.') 
                        print(str(valor_saldo))
                        # A partir daqui temos o saldo na variavel valor_saldo para poder ser utilizado para o calculo.
                        
                        
                
                    
                vlinha = vlinha + 1    

        ## Soma totalizadores (valores adicionados nas listas)                    
        for pagamento in amortizacoes:
            valor            = float(pagamento.replace(",","."))
            total_pagamentos = total_pagamentos + valor
        
        for multa in multas:
            valor        = float(multa.replace(",","."))
            total_multas = total_multas + valor

        for juro in juros:
            valor        = float(juro.replace(",","."))
            total_juros  = total_juros + valor   
                

        count = 0
        valor_total = 72000
        while count <= 59:
            if count < 24:
                valor_total = valor_total + ((valor_total * 1.60)/100)
            else:
                valor_total = valor_total + ((valor_total * 1.60)/100)
                valor_total = valor_total + ((valor_total * 1)/100)
            count = count + 1

        print('Total juros: ' + str(valor_total))

        print('Parcela da Ficha Gráfica: ' + str(parcelaInt))
        print('Data de Vencimento:       ' + str(data_parcela))
        print('Total de Pagamentos: R$   ' + str(total_pagamentos))
        print('Total de Multas:     R$   ' + str(total_multas))
        print('Total de Juros:      R$   ' + str(total_juros))
        
def pegaParcelaLinha(linha):
    if linha != '':
        capturado = linha[56:63].replace(" ", "")
                
        try:
            parcela = int(capturado)
        except:
            parcela = 0
    
    return parcela

def pegaDataLinha(linha):
    if linha != '':
        capturado = linha[1:11].replace(" ", "")                        
    
    return capturado

def pegaSaldoLinha(linha):
    if linha != '':
        capturado = linha[120:134].replace(" ", "")
        capturado = capturado.replace(".", "")
        capturado = capturado.replace(",", ".")
    
    return capturado

def pegaTxJuroLinha(linha):
    if ("TX JR NORMAL" in linha):
        capturado = linha[17:35].replace(" ", "")
        capturado = capturado.replace("%", "")
        capturado = capturado.replace("a", "")
        capturado = capturado.replace("m", "")
        capturado = capturado.replace(".", "")
        
        valor       = float(capturado.replace(",","."))
        valor_final = "{:.2f}".format(valor)        
    
    return valor_final

def pegaDebitoLinha(linha):
    if linha != '':
        capturado = linha[63:82].replace(" ", "")
        capturado = capturado.replace(".", "")
        valor     = capturado.replace(",",".")
    
    return float(valor)

def pegaCreditoLinha(linha):
    if linha != '':
        capturado = linha[83:107].replace(" ", "")
        capturado = capturado.replace(".", "")
        valor     = capturado.replace(",",".")
    
    return float(valor)

def pegaDataVencimentoParcela(linha, parcela):
    if linha != '':
        if parcela <= 99:
            posicao_inicial = linha.find('0' + str(parcela) + ')') + 5
            posicao_final   = posicao_inicial +  12
            texto_capturado = linha[posicao_inicial:posicao_final].replace(" ","")

            data_capturada = texto_capturado
    
    return data_capturada

def existeTextoLinha(linha, texto):    
    if (texto in linha):
        return True
    else:
        return False

analisaArquivo()          
          
  
                        