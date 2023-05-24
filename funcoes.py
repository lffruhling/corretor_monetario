import json
import MySQLdb
import constantes as c

def conexao():
    return MySQLdb.connect(host="10.4.21.24", user='root', passwd='*Sicred1',db='db_teste')
    # return MySQLdb.connect(host="mysql.edersondallabrida.com", user=c.USUARIO_DB1, passwd=c.SENHA_DB1,db=c.NOME_DB1)
    
def salvaParametrosImportacao(id_ficha, igpm, ipca, cdi, inpc, tr, multa_perc, multa_valor_fixo, multa_incidencia, honorarios_perc, honorarios_valor_fixo, outros_valor):
    db     = conexao()
    cursor = db.cursor()
    
    vsql = 'INSERT INTO ficha_parametros(id_ficha_grafica, igpm, ipca, cdi, inpc, tr, multa_perc, multa_valor_fixo, multa_incidencia, honorarios_perc, honorarios_valor_fixo, outros_valor)\
            VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
        
    parametros = (id_ficha, igpm, ipca, cdi, inpc, tr, multa_perc, multa_valor_fixo, multa_incidencia, honorarios_perc, honorarios_valor_fixo, outros_valor)
    
    cursor.execute(vsql, parametros)
    resultado = cursor.fetchall()
    db.commit()

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
            valor = str(resultado[0][0]*100)
            return round(float(valor),2)
        else:
            print('Sem resultados')
            return 0
    except Exception as erro:
        print('Ocorreu um erro ao tentar carregar o indice ' + str(tabela) + '. ' + str(erro))

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


#dados = carregaIndice('igpm', 2022, 3)

def calcularJuros(valorParcela, txJuro, totalMeses, composto=True):
    valorParcelaAcumulado = 0

    if composto:
        valorParcelaAcumulado   = valorParcela

    totalJurosAcumulado     = 0
    txJuroCalculada = txJuro/100

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
