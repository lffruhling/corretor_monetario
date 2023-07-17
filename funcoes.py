import MySQLdb
import constantes as c
import locale
from datetime import datetime
import socket
import requests
import os
import winreg as wrg
import PySimpleGUI as sg
import wget

def conexao():
    #return MySQLdb.connect(host="10.4.21.24", user='root', passwd='*Sicred1',db='db_teste')
    return MySQLdb.connect(host="mysql.edersondallabrida.com", user=c.USUARIO_DB1, passwd=c.SENHA_DB1,db=c.NOME_DB1)
    
def download():
    file_url = 'https://drive.google.com/file/d/19gNbvs-7WYSD6FzgJ3tImE995A6tUCfD/view?usp=sharing'
    file = 'arquivolocal.rar'
    wget.download(file_url , file )

def salvaFichaAlcadas(id_ficha, alcadas=[]):
    db     = conexao()
    cursor = db.cursor()

    vsql = 'UPDATE ficha_alcadas set situacao=%s WHERE id_ficha_grafica=%s'
    parametros = ('INATIVO', id_ficha)
    cursor.execute(vsql, parametros)    
    db.commit()    
    
    for alcada in alcadas:
        vsql = 'INSERT INTO ficha_alcadas(id_ficha_grafica, alcada, valor, situacao)\
                VALUES(%s,%s,%s,%s)'
        
        parametros = (id_ficha, alcada[0], float(str(alcada[1]).replace(",",".")), 'ATIVO')
    
        cursor.execute(vsql, parametros)
        resultado = cursor.fetchall()
        db.commit()

        if cursor.rowcount > 0:
            print('Alçada salva com sucesso!')

    cursor.close()

def salvaParametrosAlcadas(alcadas=[]):
    db     = conexao()
    cursor = db.cursor()

    vsql = 'UPDATE parametros_alcadas set situacao="INATIVO"'    
    cursor.execute(vsql,)    
    db.commit()    
    
    for alcada in alcadas:        
        vsql = 'INSERT INTO parametros_alcadas(alcada, valor, situacao)\
                VALUES(%s,%s,%s)'
        
        #valor = alcada[1].replace(",",".")
        parametros = (alcada[0], float(alcada[1]), 'ATIVO')
    
        cursor.execute(vsql, parametros)
        resultado = cursor.fetchall()
        db.commit()

        if cursor.rowcount > 0:
            print('Parâmetros definidos com sucesso!')

    cursor.close()

def carregarParametrosAlcadas():
    db     = conexao()
    cursor = db.cursor()
        
    vsql = 'SELECT alcada, valor from parametros_alcadas WHERE situacao="ATIVO"'
    cursor.execute(vsql,)
    resultado = cursor.fetchall()

    cursor.close()
    
    if len(resultado) > 0:        
        return resultado
    else:
        return []    
    
def carregarFichaAlcadas(id_ficha):
    db     = conexao()
    cursor = db.cursor()
        
    vsql = 'SELECT alcada, valor from ficha_alcadas WHERE situacao="ATIVO" and id_ficha_grafica=%s'
    parametro = (id_ficha,)
    cursor.execute(vsql,parametro)
    resultado = cursor.fetchall()

    cursor.close()
    
    if len(resultado) > 0:        
        return resultado
    else:
        return []

def salvaParametrosImportacao(id_ficha, igpm, ipca, cdi, inpc, tr, multa_perc, multa_valor_fixo, multa_incidencia, honorarios_perc, honorarios_valor_fixo, outros_valor):
    db     = conexao()
    cursor = db.cursor()
    
    vsql = 'INSERT INTO ficha_parametros(id_ficha_grafica, igpm, ipca, cdi, inpc, tr, multa_perc, multa_valor_fixo, multa_incidencia, honorarios_perc, honorarios_valor_fixo, outros_valor)\
            VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
        
    parametros = (id_ficha, igpm, ipca, cdi, inpc, tr, multa_perc, multa_valor_fixo, multa_incidencia, honorarios_perc, honorarios_valor_fixo, outros_valor)
    
    cursor.execute(vsql, parametros)
    resultado = cursor.fetchall()
    db.commit()

    if cursor.rowcount > 0:
        print('Parâmetros de importação salvos com sucesso!')

    cursor.close()

def salvarParametrosGerais(igpm, ipca, cdi, inpc, tr, multa_perc, multa_valor_fixo, multa_incidencia, honorarios_perc, honorarios_valor_fixo, outros_valor):
    db     = conexao()
    cursor = db.cursor()

    vsql_update = 'UPDATE parametros_gerais SET situacao="INATIVO"'
    cursor.execute(vsql_update,)
    db.commit()

    vsql = 'INSERT INTO parametros_gerais(igpm, ipca, cdi, inpc, tr, multa_perc, multa_valor_fixo, multa_incidencia, honorarios_perc, honorarios_valor_fixo, outros_valor, situacao, data_inclusao)\
            VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'

    parametros = (igpm, ipca, cdi, inpc, tr, multa_perc, multa_valor_fixo, multa_incidencia, honorarios_perc, honorarios_valor_fixo, outros_valor,'ATIVO',datetime.now())

    cursor.execute(vsql, parametros)
    resultado = cursor.fetchall()
    db.commit()

    if cursor.rowcount > 0:
        v_retorno = True
    else:
        v_retorno = False

    cursor.close()

    return v_retorno

def carregaParametrosGerais():
    db     = conexao()
    cursor = db.cursor()

    vsql = 'SELECT igpm, ipca, cdi, inpc, tr, multa_perc, multa_valor_fixo, multa_incidencia, honorarios_perc, honorarios_valor_fixo, outros_valor FROM parametros_gerais WHERE situacao="ATIVO"'
    cursor.execute(vsql,)
    resultado = cursor.fetchall()

    if len(resultado) > 0:
        return resultado[0]
    else:
        return None

def carregaParametrosFichaGrafica(ficha_id):
    db     = conexao()
    cursor = db.cursor()

    vsql = 'SELECT igpm, ipca, cdi, inpc, tr, multa_perc, multa_valor_fixo, multa_incidencia, honorarios_perc, honorarios_valor_fixo, outros_valor FROM ficha_parametros WHERE id_ficha_grafica=' + str(ficha_id)
    cursor.execute(vsql,)
    resultado = cursor.fetchall()

    if len(resultado) > 0:
        return resultado[0]
    else:
        return None

def carregaDadosCabecalhoFichaGrafica(ficha_id):
    db     = conexao()
    cursor = db.cursor()

    vsql = 'SELECT titulo,versao,associado,nro_parcelas,parcela,valor_financiado,tx_juro,multa,liberacao,modalidade_amortizacao FROM ficha_grafica WHERE id=' + str(ficha_id)
    cursor.execute(vsql,)
    resultado = cursor.fetchall()

    if len(resultado) > 0:
        return resultado[0]
    else:
        return None

def carregaIndice(tabela, ano, mes):
    p_mes = ''
    
    if mes == 1:
        p_mes = 'jan'
    if mes == 2:
        p_mes = 'fev'
    if mes == 3:
        p_mes = 'mar'
    if mes == 4:
        p_mes = 'abr'
    if mes == 5:
        p_mes = 'mai'
    if mes == 6:
        p_mes = 'jun'
    if mes == 7:
        p_mes = 'jul'
    if mes == 8:
        p_mes = 'ago'
    if mes == 9:
        p_mes = 'set'
    if mes == 10:
        p_mes = 'out'
    if mes == 11:
        p_mes = 'nov'
    if mes == 12:
        p_mes = 'dez'

    try:
        db = conexao()
        cursor = db.cursor()
        cursor.execute("SELECT " + p_mes + " FROM " + tabela + " WHERE ano = " + str(ano))
        resultado = cursor.fetchall()

        if len(resultado) > 0:
            valor = str(resultado[0][0])
            return round(float(valor),2)
        else:            
            return 0
    except Exception as erro:
        gravalog('Ocorreu um erro ao tentar carregar o indice ' + str(tabela) + '. ' + str(erro), True)        

def calculaPrice(valorEmprestimo, nroParcelas, taxaJuros):
        
    # valorEmprestimo=72000 #PV
    # nroParcelas=60 #N
    # taxaJuros=1.6
    
    valorParcela            = 0 #P        
    vTotalJurosAcumulado    = 0
    vTotalCapitalAmortizado = 0

    i = (taxaJuros/100)         # I
    tx = (1 + i) ** nroParcelas # I Calculado
    txA = tx * i                # Parte de cima da formula
    txB = tx - 1                # Parte de baixo da formula
    txR = txA / txB             # Formula Calculada

    valorParcela = valorEmprestimo * txR # PV * Formula Calculada = P

    print(f'Valor da Parcela {valorParcela}')

    for nroP in range(nroParcelas):
        vJuros = valorEmprestimo * i                            #72000 * 0,016 (Cada Aoperação Diminuir/Somar desse valor aqui)
        vTotalJurosAcumulado = vTotalJurosAcumulado + vJuros    #Saber o total Pago de Juros no final do Caluclo
        vCapitalAmortizado = valorParcela - vJuros              #Saber quanto do capital foi amortizado Valor da Parcela menos o Juros
        vTotalCapitalAmortizado = vTotalCapitalAmortizado + vCapitalAmortizado #Saber o quanto foi pago do capital no final do calculo
        valorEmprestimo = valorEmprestimo - vCapitalAmortizado  #Saber o valor do saldo devedor para calcular o Juros novamente


        print('#########-------############')
        print(f'Juros: {vJuros}')
        print(f'Total Parcela: {valorParcela}')
        print(f'Total Capital Amortizado: {vCapitalAmortizado}')
        print(f'Valor da Dívida atualizado {valorEmprestimo}')
        print('############################')

    print('\n\n------------Totais--------------------')
    print(f'Total Juros {vTotalJurosAcumulado}')
    print(f'Total Amortizado {vTotalCapitalAmortizado}')
    print(f'Total Devedor Emprestimo {valorEmprestimo}')
    print(f'Total pago {vTotalJurosAcumulado + vTotalCapitalAmortizado}')

def pegarIp():
    try:
        return socket.gethostbyname(socket.gethostname())
    except:
        return '0.0.0.0'
    
def pegarNomeMaquina():
    try:
        return socket.gethostname()
    except:
        return 'Máquina não identificada'
    
def testarInternet():
    #Checar conexão de internet
    url = 'http://www.google.com.br'
    timeout = 5
    try:
        requests.get(url, timeout=timeout)
        return 'Conectado'
    except:
        return 'Sem Conexão com a Internet'

def pastaExiste(vLocal, vCriarDiretorio=False):
    if vCriarDiretorio:
        try:
            if os.path.isdir(vLocal):
                return True
            else:
                os.makedirs(vLocal)
                return True
        except Exception as e:            
            return False
    else:
        return os.path.isdir(vLocal)

def gravalog(msg, is_erro=False):
    maquina = pegarNomeMaquina()
    ip      = pegarIp()

    try:
        if is_erro:
            vtipo = 'ERRO'
            vmensagem = f'\n{datetime.now().strftime("%m/%d/%Y %H:%M:%S")}    ERRO: {str(msg)}'
        else:
            vtipo = 'INFO'
            vmensagem = f'\n{datetime.now().strftime("%m/%d/%Y %H:%M:%S")}    INFO: {str(msg)}'
            
        vNome    = datetime.now().strftime("%d-%m-%Y") + '_log.txt'
        vLocal   = 'C:\Temp\Corretor_Monetario_Logs'
        pastaExiste(vLocal, True)
        vArquivo = f'{vLocal}\{vNome}'

        if os.path.isfile(vArquivo):
            with open(vArquivo, 'a') as log:
                log.write(vmensagem)
        else:
            with open(vArquivo, 'w') as log:
                log.write(vmensagem)
                        
        try:
            vconexao = conexao()
            cursor  = vconexao.cursor()  
                                                            
            vparametros = (str(datetime.now().strftime("%Y-%m-%d %H:%M:%S")), str(vtipo), 'Corretor Monetário', str(msg), str(maquina), str(ip))
            
            cursor.execute('INSERT INTO logs(data_hora,tipo,sistema,descricao,host,ip) VALUES(%s,%s,%s,%s,%s,%s)',vparametros)    
            vconexao.commit()

            cursor.close()
            vconexao.close()  
        except Exception as erro:
            print('Falha ao tentar gravar log no Db. ' + str(erro))
    
        print(vmensagem)
    except Exception as erro:
        log('Falha ao tentar gravar log na pasta Temp. ' + str(erro)) 
    
def BuscaUltimaVersao():
    try:
        db = conexao()
        cursor = db.cursor()
        cursor.execute("SELECT versao FROM versao order by id desc limit 1")
        resultado = cursor.fetchall()

        if len(resultado) > 0:
            valor = str(resultado[0][0])
            return valor
        else:            
            return None
    except Exception as erro:
        gravalog('Ocorreu um erro ao tentar carregar a última versão gerada. '  + str(erro), True)   

def verificaLicenca():
    try:
        registry_key = wrg.OpenKey(wrg.HKEY_CURRENT_USER, r"SOFTWARE\\CorretorMonetario", 0,
                                       wrg.KEY_READ)
        value, regtype = wrg.QueryValueEx(registry_key, "Chave")
        wrg.CloseKey(registry_key)
        return value
    except WindowsError:
        gravalog('Falha ao tentar verificar licença do Registro do Windows', True)
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

def identificaVersao(caminho_txt):
    versao = 'Não Identificada'
    with open(caminho_txt, 'r') as arquivo_txt:        
        vlinha = 1
        for linha in arquivo_txt:
            if vlinha <= 20:
                if("CRESOL" in linha):
                    return 'cresol'                    

                if("COOP CRED POUP E INVEST" in linha):
                    return 'sicredi'                    

                if("COOP.CRED.POUP.INVESTIMENTO CONEXAO" in linha):
                    return 'sicredi'                    

                if("C.C.P.I. DA REGIAO DA PRODUCAO" in linha):
                    return 'sicredi'  

                if("COOP.CRED.POUP.INVESTIMENTO ALTO URUGUAI" in linha):
                    return 'sicredi'                                  

            else:
                break 
            
            vlinha = vlinha + 1
    
    return versao

#dados = carregaIndice('igpm', 2022, 3)

def calcularJurosPrice(valorParcela, txJuro, totalMeses, composto=True):
    valorParcelaAcumulado = 0

    if composto:
        valorParcelaAcumulado   = valorParcela

    totalJurosAcumulado     = 0
    txJuroCalculada         = txJuro/100

    for i in range(totalMeses + 1):
        if composto:
            vTotalJuros = valorParcelaAcumulado * txJuroCalculada
            valorParcelaAcumulado = valorParcelaAcumulado + vTotalJuros
        else:
            vTotalJuros = valorParcela * txJuroCalculada

        totalJurosAcumulado = totalJurosAcumulado + vTotalJuros

    if composto:
        return {
            'totalJurosAcumuladoPeriodo': totalJurosAcumulado,
            'valorParcelaAtualizada'    : valorParcelaAcumulado
        }
    else:
        return totalJurosAcumulado

def moeda(valor):
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    valor = locale.currency(valor, grouping=True, symbol=None)
    return ('R$ %s' % valor)
