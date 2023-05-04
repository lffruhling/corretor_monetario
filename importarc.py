from datetime import datetime
import funcoes as f

vLinhaLancamento = False
vParcelas        = 0
array_datas      = ['01/','02/','03/','04/','05/','06/','07/','08/','09/','10/','11/','12/',
                    '13/','14/','15/','16/','17/','18/','19/','20/','21/','22/','23/','24/',
                    '25/','26/','27/','28/','29/','30/','31/']

## Variaveis cabeçalho
associado        = ''
parcelas         = 0
parcela          = 0
titulo           = ''
versao           = 'cresol'
valor_financiado = 0
tx_juro          = 0
multa            = 0
vCET             = 0
data_liberacao   = datetime.now()
modalidade_amortizacao = ''

## Variaveis Detalhe
parcela          = 0
data_vencimento  = None
data_movimento   = None
operacao         = 0
descricao        = ''
valor            = 0
saldo            = 0
tipo             = ''
valor_credito    = 0
valor_debito     = 0

def importar_cabecalho(vArquivoTxt):    
    global titulo
    global modalidade_amortizacao
    with open(vArquivoTxt, 'r') as arquivo:        
        vlinha = 1
        for linha in arquivo:                                    
            if vlinha <= 50:
                if("Parcelas:" in linha):                
                    parcelas = int(linha[10:15].replace(" ",""))

                if("Forma de Amortização:" in linha):
                    if("SAC" in linha):
                        modalidade_amortizacao = 'SAC'

                    if("SPC" in linha):
                        modalidade_amortizacao = 'SAC'

                    if("SAV" in linha):
                        modalidade_amortizacao = 'SAV'

                    if("PRICE" in linha):
                        modalidade_amortizacao = 'PRICE'
                    
                if("Nome:" in linha):   
                    vnome     = linha.split(":")
                    associado = vnome[1][1:]
                    
                if("Juros ao Mês:" in linha):   
                    vjuros  = linha.split(":")
                    vjuros  = vjuros[1].replace(" ","")
                    vjuros  = vjuros.replace("%","")
                    vjuros  = vjuros.replace(",",".")
                    tx_juro = float(vjuros)                
                    
                if("Multa:" in linha):   
                    vmulta  = linha.split(":")
                    vmulta  = vmulta[1].replace(" ","")
                    vmulta  = vmulta.replace("%","")
                    vmulta  = vmulta.replace(",",".")
                    multa   = float(vmulta)                                            
                    
                if("CET ao Mês:" in linha):   
                    vcet = linha.split(":")
                    vcet = vcet[1].replace(" ","")
                    vcet = vcet.replace("%","")
                    vcet = vcet.replace(",",".")
                    vCET = float(vcet)                
                
                if("Valor do Contrato:" in linha):   
                    vvalor = linha.split(":")
                    vvalor = vvalor[1].replace(" ","")
                    vvalor = vvalor.replace(".","")
                    vvalor = vvalor.replace(",",".")
                    valor_financiado = float(vvalor)
                    
                if("Data Liberação:" in linha):
                    vdata  = linha.split(':')
                    vdata  = vdata[1].replace(" ","")
                    data_l = vdata.split('/')
                    dia_l  = data_l[0]
                    mes_l  = data_l[1]
                    ano_l  = data_l[2]                                                
                    data_liberacao = datetime(int(ano_l),int(mes_l),int(dia_l))
                
                ## Titulo    
                if("Contrato Agrupado:" in linha):   
                    vcontrato  = linha.split(":")
                    vcontrato  = vcontrato[1].replace(" ", "")
                    titulo     = vcontrato
                
                ## Parcela Atual
                #if("Nome:" in linha):   
                #    vnome     = linha.split(":")
                #    associado = vnome[1][1:]
            else:
                break

            vlinha = vlinha + 1

    ## Executa a Inserção do cabeçalho no BD
    db     = f.conexao()
    cursor = db.cursor()

    sql_update = 'UPDATE ficha_grafica SET situacao = "INATIVO" WHERE titulo = %s AND situacao = "ATIVO"'
    cursor.execute(sql_update, (titulo,))
    cursor.fetchall()
    db.commit()

    vsql = 'INSERT INTO ficha_grafica(versao, associado, liberacao, nro_parcelas, parcela, situacao, titulo, tx_juro, multa, valor_financiado, modalidade_amortizacao)\
        VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
    
    parametros = ('cresol', str(associado), data_liberacao, parcelas, parcela, 'ATIVO', titulo, tx_juro, multa, valor_financiado, modalidade_amortizacao)

    cursor.execute(vsql, parametros)
    resultado = cursor.fetchall()
    db.commit()

    if cursor.rowcount > 0:
        print('Cabeçalho importado com sucesso!')
        return True

def importar_detalhes(vArquivoTxt, titulo):    
    global vLinhaLancamento        
    global parcela
    global data_vencimento
    global data_movimento
    global operacao
    global descricao
    global valor
    global saldo
    global tipo
    global valor_credito
    global valor_debito

    db     = f.conexao()
    cursor = db.cursor()

    ## Inativa uma possível importação anterior
    sql_update = 'UPDATE ficha_detalhe SET situacao = "INATIVO" WHERE titulo = %s AND situacao = "ATIVO"'
    cursor.execute(sql_update, (titulo,))
    cursor.fetchall()
    db.commit()
    
    ## Reliza leitura do TXT
    with open(vArquivoTxt, 'r') as arquivo:        
        vlinha = 1
        for linha in arquivo:                                    
                                            
            ## Divide a linha em um array de 4 partes
            linha_atual = linha.split(" ", 4)
            
            ## identifica se é uma linha de lançamento de movimentação
            if(linha_atual[1][0:3] in array_datas):
                vLinhaLancamento = True                    

            ## Se for linha de lançamento, captura as informações 
            if vLinhaLancamento:                    
                descricao = ''
                parcela   = int(linha_atual[0]) 
                                                            
                data_v = linha_atual[1].split('/')
                dia_v  = data_v[0]
                mes_v  = data_v[1]
                ano_v  = data_v[2]                                                
                data_vencimento = datetime(int(ano_v),int(mes_v),int(dia_v))

                data_m = linha_atual[2].split('/')
                dia_m  = data_m[0]
                mes_m  = data_m[1]
                ano_m  = data_m[2]                                                                                  
                data_movimento  = datetime(int(ano_m),int(mes_m),int(dia_m))

                operacao = int(linha_atual[3])

                ## Fatia o restante da string para montar a descrição, removendo os valores
                texto = linha_atual[4].split(' ')
                #print(str(texto))
                    
                for posicao in texto:                        
                    try:
                        texto_          = posicao.replace('.','')                            
                        valor_capturado = float(texto_.replace(',','.'))
                    except:                                                        
                        if descricao != '':
                            descricao = descricao + ' ' + str(posicao)
                        else:
                            descricao = str(posicao)
                
                descricao = descricao[:-2]
                
                ## Cria variavel removendo o texto da descrição para sobrar apenas os valores para fatiar
                string_valores = linha_atual[4].replace(descricao, "").split(' ')
                #print(str(string_valores))
                
                ## Captura valor, substitui virgulas por ponto, converte em float
                str_valor = string_valores[0].replace('.','')
                str_valor = str_valor.replace(',','.')
                valor     = float(str_valor)
                
                ## Captura saldo, substitui virgulas por ponto, converte em float
                str_saldo = string_valores[2].replace('.','')
                str_saldo = str_saldo.replace(',','.')
                saldo     = float(str_saldo)
                
                ## Captura o tipo crédito ou débito
                tipo_capturado = string_valores[3]
                tipo = tipo_capturado[0:1]
                
                if str(tipo) == 'D':
                    valor_credito = 0
                    valor_debito  = valor                    

                if str(tipo) == 'C':
                    valor_credito = valor
                    valor_debito  = 0                     

                ## Realiza a inserção no BD
                vsql = 'INSERT INTO ficha_detalhe(titulo, data, cod, historico, parcela, situacao, valor_credito, valor_debito, valor_saldo)\
                        VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s)'                
                    
                parametros = (str(titulo), data_movimento, str(operacao), descricao, str(parcela), "ATIVO", valor_credito, valor_debito, saldo)                
                cursor.execute(vsql, parametros)
                resultado = cursor.fetchall()                                               

            vlinha = vlinha + 1
            vLinhaLancamento = False

    db.commit()
    if cursor.rowcount > 0:
        print('Detalhe Ficha Gráfica Cresol importado com sucesso!')
    
def importarCresol(vArquivoTxt):
    print('Importando arquivo: ' + str(vArquivoTxt))
    if importar_cabecalho(vArquivoTxt):
        importar_detalhes(vArquivoTxt, titulo)

    return titulo