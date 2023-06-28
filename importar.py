import funcoes as f
from datetime import datetime, date, timedelta
import re
import time

def pegaParcelaLinha(linha):    
    parcela = 0

    try:
        if linha != '':
            if('NUMERO DE PARCELAS ...:' in linha):
                ## Encontra as posições iniciais e finais da frase na linha
                pos_inicial = linha.find('NUMERO DE PARCELAS ...:')
                pos_final   = pos_inicial + len('NUMERO DE PARCELAS ...:')
                
                ## Captura o valor a partir da posição final da frase até o final da linha, remove os espaços em branco
                valor_capturado = linha[pos_final:].replace(' ','')       
                valor_capturado = valor_capturado.rstrip()                

                parcela = valor_capturado.split('/')                     

                return parcela        
        
    except Exception as erro:
        f.gravalog('Ocorreu um erro ao tentar identificar as informações de parcela. ' + str(erro) + '. LINHA DO ARQUIVO: ' + str(linha), True)
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

    try:
        if linha != '':
            if ("TX JR NORMAL" in linha):
                pos_inicial = linha.find("TX JR NORMAL .:")
                pos_final   = pos_inicial + len("TX JR NORMAL .:")

                pos_inicial_proximo_campo = linha.find('PERCENT CM NORMAL')

                capturado   = linha[pos_final:pos_inicial_proximo_campo-1].replace(' ','')      
                capturado = capturado.replace("%", "")
                capturado = capturado.replace("a", "")
                capturado = capturado.replace("m", "")
                capturado = capturado.replace(".", "")
                
                valor       = float(capturado.replace(",","."))
                valor_final = "{:.2f}".format(valor)                                          
    except Exception as erro:
        f.gravalog('Ocorreu um erro ao tentar identificar a taxa de Juros. ' + str(erro) + '. LINHA DO ARQUIVO: ' + str(linha))

    finally:
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
        pos_inicial = linha.find("ASSOCIADO ....:")
        pos_final   = pos_inicial + len("ASSOCIADO ....:")
        
        ## identifica onde começa o próximo campo para saber o intervalo para capturar a informação do nome
        pos_inicial_proximo_campo = linha.find("SITUACAO ..:")

        ## Remove o número da conta da frante do nome do associado
        remover = ['0','1','2','3','4','5','6','7','8','9','-']
        capturado = linha[pos_final:pos_inicial_proximo_campo]
        for item in remover:
            capturado = capturado.replace(item,'')

        ## Remove os espaços em branco do inicio e fim da string
        capturado = capturado.strip()
    
    return capturado

def pegaValorFinanciadoLinha(linha):
    if linha != '':
        if("VALOR FINANCIADO .....:" in linha):
            pos_inicial = linha.find("VALOR FINANCIADO .....:")
            pos_final   = pos_inicial + len("VALOR FINANCIADO .....:")

            capturado   = linha[pos_final:].replace(' ', '')        
            capturado = capturado.replace(".", "")
            capturado = capturado.replace(",", ".")
            capturado = capturado.rstrip()
    
    return capturado

def pegaDataLiberacaoLinha(linha):
    data_final = None

    try:
        if linha != '':
            if('LIBERACAO .:' in linha):
                pos_inicial = linha.find('LIBERACAO .:')
                pos_final   = pos_inicial + len('LIBERACAO .:')

                capturado = linha[pos_final:pos_final + 12].replace(' ', '')            
                capturado = capturado.replace("/", "-")            
                            
                data = capturado.split('-')
                dia = data[0]
                mes = data[1]
                ano = data[2]        
            
                data_formatada = ano + '-' + mes + '-' + dia
                data_final = datetime(int(ano),int(mes),int(dia))
    except Exception as erro:
        f.gravalog('Ocorreu um erro ao tentar capturar a data de liberação. ' + str(erro) + '. LINHA ARQUIVO: ' + str(linha))      

    finally:
        return data_final

def importaFichaGrafica(vCaminhoTxt, informacoes=False):

    if informacoes:
        f.gravalog('Inicia Importação das Informações Sicredi. Arquivo: ' + str(vCaminhoTxt))
    else:
        f.gravalog('Inicia Importação Cabeçalho Sicredi. Arquivo: ' + str(vCaminhoTxt))

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
                pos_inicial = linha.find("TITULO:")
                pos_final   = pos_inicial + len("TITULO:")

                v_titulo = linha[pos_final:].replace(' ', '')                
                titulo = v_titulo.rstrip()

            if ("COMPOSICAO ...:" in linha and vlinha <= 10):
                if("PRICE" in linha):
                    modalidade_amortizacao = 'PRICE'

                if("SAC" in linha):
                    modalidade_amortizacao = 'SAC'

                if("SAV" in linha):
                    modalidade_amortizacao = 'SAV'

                if("SPV" in linha):
                    modalidade_amortizacao = 'SPV'

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
                vparcela = parcela                
                if (parcela == nro_parcelas):
                    vparcela = parcela -1
                dataEntradaPrejuizo = pegaDataVencimentoParcela(linha, vparcela)

            vlinha = vlinha + 1

        print('Data Entrada Prejuizo: ' + str(dataEntradaPrejuizo)) #Salvar a data de entrada no prejuizo na base

    if informacoes:
        return 'Coop: Sicredi   Modalidade: ' + str(modalidade_amortizacao), 'Associado: ' + str(associado), 'Data de Liberação: ' + str(liberacao.strftime('%d/%m/%Y')), 'Número de Parcelas: ' + str(nro_parcelas), 'Parcela atual: ' + str(parcela), 'Título: ' + str(titulo), 'Taxa de Juros: ' + str(taxa_juro), 'Valor Financiado: ' + str(valor_financiado), 'Data da Inadimplência: ' + dataEntradaPrejuizo.strftime("%d/%m/%Y")
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
            f.gravalog('Cabeçalho Ficha Gráfica Sicredi importado com sucesso. Id: ' + str(cursor.lastrowid))
            return cursor.lastrowid

def importaFichaGraficaDetalhe(vArquivoTxt, id_ficha_grafica):

    f.gravalog('Inicia importação dos Detalhes da Ficha Gráfica Sicredi. Id: ' + str(id_ficha_grafica))

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
        f.gravalog('Detalhe Ficha Gráfica Sicredi importado com sucesso. Id: ' + str(id_ficha_grafica))        

    return titulo
            
def importarSicredi(vCaminhoTxt, parametros):
    print('Importando arquivo: ' + str(vCaminhoTxt))
    ficha_grafica = importaFichaGrafica(vCaminhoTxt)
    if ficha_grafica > 0:
        f.salvaParametrosImportacao(ficha_grafica, parametros['igpm'], parametros['ipca'], parametros['cdi'], parametros['inpc'], parametros['tr'], parametros['multa_perc'],\
                                    parametros['multa_valor'], parametros['multa_incidencia'], parametros['honorarios_perc'], parametros['honorarios_valor'], parametros['outros_valor'])
        ficha_titulo = importaFichaGraficaDetalhe(vCaminhoTxt, ficha_grafica)        

    return ficha_titulo, ficha_grafica