import funcoes as f
from datetime import datetime, date, timedelta
import re
def pegaParcelaLinha(linha):
    if linha != '':
        capturado = linha[116:134].replace(" ", "")
        capturado = capturado.split('/')
             
        try:
            parcela = capturado
        except:
            parcela = 0
    
    return parcela

def pegaDataLinha(linha):
    if linha != '':
        capturado = linha[0:11].replace(" ", "")                        

        capturado = capturado.replace("/", "-")
                        
        data = capturado.split('-')
        dia = data[0]
        mes = data[1]
        ano = data[2]        
        
        data_formatada = ano + '-' + mes + '-' + dia
        data_final = datetime(int(ano),int(mes),int(dia))
    
    return data_final

def pegaSaldoLinha(linha):
    if linha != '':
        capturado = linha[120:134].replace(" ", "")
        
        if capturado != '':
            capturado = capturado.replace(".", "")
            valor = capturado.replace(",", ".")
        else:
            valor = 0
    
    return float(valor)

def pegaTxJuroLinha(linha):
    valor_final = 0
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
        if capturado != '':
            capturado = capturado.replace(".", "")
            valor     = capturado.replace(",",".")
        else:
            valor = 0
    
    return float(valor)

def pegaCreditoLinha(linha):
    if linha != '':
        capturado = linha[83:107].replace(" ", "")
        
        if capturado != '':
            capturado = capturado.replace(".", "")
            valor     = capturado.replace(",",".")
        else:
            valor = 0
    
    return float(valor)

def pegaDataVencimentoParcela(linha, parcela):
    # Alterar aqui para pegar a parcal e data
    data_capturada = None
    if linha != '':
        if (f"0{parcela + 1})" in linha):
            match=(re.search(f"0{parcela + 1}", linha))
            data_capturada = linha[match.end()+1:(match.end() + 12)].replace(" ", "")
            dataPrejuizo = datetime.strptime(data_capturada, '%d/%m/%Y')
            data_capturada = dataPrejuizo + timedelta(days=1)
            return data_capturada

def existeTextoLinha(linha, texto):    
    if (texto in linha):
        return True
    else:
        return False
    
def pegaCodigoLinha(linha):
    if linha != '':
        capturado = linha[12:17].replace(" ", "")
    
    return capturado

def pegaHistoricoLinha(linha):
    if linha != '':
        capturado = linha[17:59]
    
    return capturado

def pegaParcelaDetalheLinha(linha):
    if linha != '':
        capturado = linha[59:63].replace(" ", "")
    
    return capturado

def pegaAssociadoLinha(linha):
    capturado = ''
    if ("ASSOCIADO ....:" in linha):
        capturado = linha[24:57]                              
    
    return capturado

def pegaValorFinanciadoLinha(linha):
    if linha != '':
        capturado = linha[116:134].replace(" ", "")
        capturado = capturado.replace(".", "")
        capturado = capturado.replace(",", ".")
        capturado = capturado.rstrip('\n')
    
    return capturado

def pegaDataLiberacaoLinha(linha):
    data_final = None
    if linha != '':
        capturado = linha[70:81].replace(" ", "")
        capturado = capturado.replace("/", "-")
                        
        data = capturado.split('-')
        dia = data[0]
        mes = data[1]
        ano = data[2]        
        
        data_formatada = ano + '-' + mes + '-' + dia
        data_final = datetime(int(ano),int(mes),int(dia))
    
    return data_final

def importaFichaGrafica(vCaminhoTxt, informacoes=False):
    global valor_financiado        
        
    taxa_juro        = 0    
    associado        = ''
    nro_parcelas     = 0
    parcela          = 0
    valor_financiado = 0
    liberacao        = datetime.now()    
    titulo           = ''
    modalidade_amortizacao = ''
    dataEntradaPrejuizo = None

    with open(vCaminhoTxt, 'r') as reader:
        
        ficha_grafica = reader.readlines()  
                    
        vlinha = 1        
        for linha in ficha_grafica:         

            if ("TITULO:" in linha and vlinha <= 5):
                v_titulo = linha[122:134].replace(" ", "")
                titulo = v_titulo[:-1]

            if ("COMPOSICAO ...:" in linha and vlinha <= 10):
                if("PRICE" in linha):
                    modalidade_amortizacao = 'PRICE'

                if("SAC" in linha):
                    modalidade_amortizacao = 'SAC'

                if("SAV" in linha):
                    modalidade_amortizacao = 'SAV'

            if("TX JR NORMAL" in linha):
                taxa_juro = pegaTxJuroLinha(linha)
                
            if("ASSOCIADO ....:" in linha):
                associado = pegaAssociadoLinha(linha)
                
            if("NUMERO DE PARCELAS ...:" in linha):                    
                dados = pegaParcelaLinha(linha)
                
                nro_parcelas = int(dados[1])
                parcela      = int(dados[0])
                
            if("VALOR FINANCIADO .....:" in linha):
                valor_financiado = pegaValorFinanciadoLinha(linha)    
                
            if("VALOR FINANCIADO .....:" in linha):
                liberacao = pegaDataLiberacaoLinha(linha)

            if dataEntradaPrejuizo is None:
                dataEntradaPrejuizo = pegaDataVencimentoParcela(linha, parcela)

            vlinha = vlinha + 1

        print(dataEntradaPrejuizo) #Salvar a data de entrada no prejuizo na base

    if informacoes:
        return 'Coop: Sicredi   Modalidade: ' + str(modalidade_amortizacao), 'Associado: ' + str(associado), 'Data de Liberação: ' + str(liberacao.strftime('%d/%m/%Y')), 'Número de Parcelas: ' + str(nro_parcelas), 'Parcela atual: ' + str(parcela), 'Título: ' + titulo, 'Taxa de Juros: ' + taxa_juro, 'Valor Financiado: ' + valor_financiado, 'Data da Inadimplência: ' + dataEntradaPrejuizo.strftime("%m/%d/%Y")
    else:            
        db     = f.conexao()
        cursor = db.cursor()

        sql_update = 'UPDATE ficha_grafica SET situacao = "INATIVO" WHERE titulo = %s AND situacao = "ATIVO"'
        cursor.execute(sql_update, (titulo,))
        cursor.fetchall()
        db.commit()

        vsql = 'INSERT INTO ficha_grafica(versao, associado, liberacao, nro_parcelas, parcela, situacao, titulo, tx_juro, valor_financiado, modalidade_amortizacao, entrada_prejuizo)\
            VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
        
        parametros = ('sicredi', str(associado), liberacao, nro_parcelas, parcela, 'ATIVO', titulo, taxa_juro, valor_financiado, modalidade_amortizacao, dataEntradaPrejuizo)

        cursor.execute(vsql, parametros)
        resultado = cursor.fetchall()
        db.commit()

        if cursor.rowcount > 0:        
            print('Cabeçalho Ficha Gráfica Sicredi importado com sucesso!')
            return cursor.lastrowid

def importaFichaGraficaDetalhe(vArquivoTxt, id_ficha_grafica):
    global valor_financiado        
        
    data            = datetime.now()
    codigo          = 0    
    historico       = ''
    valor_debito    = 0
    valor_credito   = 0
    valor_saldo     = 0
    parcela         = 0
    situacao        = 'ATIVO'
    titulo          = ''

    array_datas = ['01/','02/','03/','04/','05/','06/','07/','08/','09/','10/','11/','12/',
                    '13/','14/','15/','16/','17/','18/','19/','20/','21/','22/','23/','24/',
                    '25/','26/','27/','28/','29/','30/','31/']

    with open(vArquivoTxt, 'r') as reader:
        db     = f.conexao()
        cursor = db.cursor()

        encontrou_titulo = False
        
        ficha_grafica = reader.readlines()  
                    
        vlinha = 1        
        for linha in ficha_grafica:         
            if ("TITULO:" in linha and vlinha <= 5 and encontrou_titulo == False):
                titulo = linha[122:134].replace(" ", "")
                titulo = titulo[:-1]

                sql_update = 'UPDATE ficha_detalhe SET situacao = "INATIVO" WHERE titulo = %s AND situacao = "ATIVO"'
                cursor.execute(sql_update, (titulo,))
                cursor.fetchall()
                db.commit()

                encontrou_titulo = True

            if(linha[0:3] in array_datas):
                data          = pegaDataLinha(linha)
                codigo        = pegaCodigoLinha(linha)
                historico     = pegaHistoricoLinha(linha)
                parcela       = pegaParcelaDetalheLinha(linha)
                valor_debito  = pegaDebitoLinha(linha)
                valor_credito = pegaCreditoLinha(linha)
                valor_saldo   = pegaSaldoLinha(linha)
                situacao      = "ATIVO"                                                                
                
                vsql = 'INSERT INTO ficha_detalhe(id_ficha_grafica, titulo, data, cod, historico, parcela, situacao, valor_credito, valor_debito, valor_saldo)\
                    VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
                
                parametros = (id_ficha_grafica, str(titulo), data, str(codigo), historico, str(parcela), situacao, valor_credito, valor_debito, valor_saldo)
                
                cursor.execute(vsql, parametros)
                resultado = cursor.fetchall()                   

                vlinha = vlinha + 1
    db.commit()
    if cursor.rowcount > 0:
        print('Detalhe Ficha Gráfica Sicredi importado com sucesso!')

    return titulo
            
def importarSicredi(vCaminhoTxt):
    print('Importando arquivo: ' + str(vCaminhoTxt))
    ficha_grafica = importaFichaGrafica(vCaminhoTxt)
    if ficha_grafica > 0:
        ficha_titulo = importaFichaGraficaDetalhe(vCaminhoTxt, ficha_grafica)        

    return ficha_titulo