# -*- coding: utf-8 -*-

"""
******************************************************************************
Lê e escreve dados em arquivos binários do tipo 'vazoes.dat'.
Este tipo de arquivo é utilizado nos modelos 'Newave', 'Decomp', 'Gevazp' e 
'Dessem'.
              

Autor   : Nelson Rossi Bittencourt
Versão  : 0.1
Licença : MIT
Dependências: numpy e openpyxl (se desejar ler dados do Excel).
******************************************************************************
"""

from os import scandir
import numpy as np
from openpyxl import load_workbook


# Classe que conterá um histórico de vazões.
# anoInicial : deverá, obrigatoriamente, ser fornecido pelo usuário e corresponderá ao primeiro ano do histórico;
# anoFinal : será calculado com base em anoInicial e no número de registros do arquivo;
# numPostos : deverá ser fornecido pelo usuário;
# valores : dicionário que conterá o número do posto como chave e um lista com todos os valores lidos.
class historicoVazoes:
    def __init__(self):
        self.anoInicial = 0
        self.anoFinal = 0
        self.numPostos = 0
        self.valores = {}


def lerVazoesExcel(nomeArquivoExcel, linIni, colIni, linFim,colFim):
    """
    Lê valores de vazão de uma planilha Excel (xlsx) para atualizar arquivo binário de vazões.
    A plalinha deverá conter:

        Primeira linha (linIni) - meses/anos das vazões a serem atualizadas/inseridas;
        Primeira coluna (colIni) - os números dos postos a terem valores atualizados/inseridos;
        Demais intervalos: dados das vazões a alterar.

        Exemplo:

        Posto   Jan/2018    Fev/2018    Mar/2018
            1        100         200         300
          320        500         600         700
        
        A primeira célula (linIni, colIni) será ignorada.
    
    Argumentos
    ----------

    nomeArquivoExcel : Caminho completo para um arquivo Excel (xlsx) com estrutura de dados compatível;

    linIni e linFim : Linhas inicial e final do intervalo de dados a ser lido do Excel;

    colIni e colFim : Colunas inicial e final do intervalo de dados a ser lido do Excel;


    Retorno
    -------

    Dicionário contendo como chave o número do posto e como valor um lista com sub-lista na forma
    [mes ano valor].

    """
    
    outPut = {}                         # Dicionário de saída
    meses = []                          # Lista de meses (lidos da coluna de cabeçalho do Excel)
    anos = []                           # Lista de anos (lidos da coluna de cabeçalho do Excel)
    numCols = colFim - colIni           # Número de colunas do intervalo

    # Abre o arquivo Excel especificado, como somente leitura e obtendo apenas valores.
    wb = load_workbook(filename=nomeArquivoExcel, read_only=True, data_only=True)
    
    # Seleciona a aba de dados (Worksheet) 'Dados'.
    ws = wb['Dados']

    # Lê cabelahaço com os meses e anos.
    for coluna in range(colIni+1,colFim+1):
        data = ws.cell(linIni, coluna).value
        meses.append(data.month)
        anos.append(data.year)

    # Lê os dados e aloca na variáve de saída.
    for linha in range(linIni+1,linFim+1):
        posto = ws.cell(linha,colIni).value            
        listaDados = []
        
        for coluna in range(0, numCols):
            valor = int(ws.cell(linha,coluna+colIni+1).value)                
            listaDados.append([meses[coluna], anos[coluna], valor])
        
        outPut[posto] = listaDados           
    
    # Fecha o Excel.
    wb.close()

    return outPut


def leVazoes(nomeArquivo, anoInicial=1931, numPostos=320):
    """
    Lê todas as vazões de um arquivo binário.

    Argumentos
    ----------

    nomeArquivo : nome do arquivo binário de vazões no formato ONS;

    anoInicial : (Opcional) ano inicial do histórico de vazões. Default: 1931.

    numeroPostos : (Opcional) número de postos contidos no histórico de vazões. Default: 320.
        O ONS utiliza 320 postos para o horizonte de operação e 600 postos para o horizonte de planejamento.

    
    Retorno
    -------

    Lista de objetos tipo 'historicoVazoes'.

    """

    # Contador do número de registros para 
    numRegistros = 0        

    # Cria instância da classe 'historicoVazoes'.
    localVazoesLidas = historicoVazoes()
    
    localVazoesLidas.anoInicial = anoInicial
    localVazoesLidas.numPostos = numPostos

    # Cria listas vazias para conter as vazões.
    for i in range(1,numPostos+1):        
        localVazoesLidas.valores[i] = []

    # Contador do número de postos.
    posto = 1

    # Abre o arquivo e aloca seus dados na lista.
    try:        
        with open(nomeArquivo, 'rb') as f:
            while (byte1:=f.read(4)):                                
                numRegistros = numRegistros + 1
                tmp = int.from_bytes(byte1,'little')                
                localVazoesLidas.valores[posto].append(tmp)
                posto = posto + 1
                if (posto==(numPostos+1)):
                    posto = 1             
    except:
        print("Erro ao abrir arquivo!")

    # Calcula o ano final do arquivo.
    anoFinal = int((anoInicial + (numRegistros/(12*numPostos)))-1)
    localVazoesLidas.anoFinal = anoFinal
    
    return localVazoesLidas


def salvaArquivo(nomeArquivo,vazoes, tipo='binario'):
    """
    Salva os dados binários de vazão no arquivo especificado.

    Argumentos
    ----------
    
    nomeArquivo : nome do arquivo a salvar;

    vazoes: objeto do tipo 'historicoVazoes' com os dados a serem salvos;

    tipo : (Opcional) especifica o tipo de arquivo a salvar. Default:'binario'
         Existem três tipos possíveis:

            'binário' - arquivo de vazões binário no formato dos modelos do setor elétrico;

            'csv' - arquivo separado por vírgulas para uso no Excel e

            'vazEdit' - formato idêntico ao produzido pelo aplicativo 'VazEdit' do ONS.


    Retorno
    -------

    Nenhum.

    """

    # Número de registros para cada posto de vazão.
    n = len(vazoes.valores[1])

    if (tipo=='binario'):
        #  Salva as vazões no arquivo binário.
        try:        
            with open(nomeArquivo, 'wb') as f:
                for n1 in range(0,n):
                    for posto in range(1,vazoes.numPostos+1):
                        f.write(vazoes.valores[posto][n1].to_bytes(4,'little'))               
                
        except:
            print("Erro ao salvar arquivo!")


    elif (tipo=='vazEdit' or 'tipo=csv'):
        try:

            if (tipo == 'csv'):
                sep = ','
                adjusts = [0,0,0]
            else:
                sep = ''
                adjusts = [3,6,5]

            with open(nomeArquivo, 'w') as f:
                for posto in range(1,vazoes.numPostos+1):
                    ano = vazoes.anoInicial
                    sPosto = str(posto).rjust(adjusts[0])                    
                    for n1 in range(0,n,12):
                        sVazoes = ''
                        valores = vazoes.valores[posto][n1:n1+12]
                        for m in range(0,12):                            
                            sVazoes = sVazoes + str(valores[m]).rjust(adjusts[1]) + sep
                        
                        sAno = str(ano).rjust(adjusts[2])
                        saida = sPosto + sep +  sAno + sep + sVazoes
                        f.write(saida+'\n')               
                        ano = ano + 1
                
        except:
            print("Erro ao salvar arquivo!")


def mudaVazao(Vazoes, posto,mes,ano,valor):
    """
    Altera ou inclui valores de um objeto 'historicoVazoes' para posterior uso/salvamento.
    
    Caso o ano especificado não faça parte do horizonte, serão incluídos vetores com valor zero
    de modo a manter o arquivo compatível.
    
    Argumentos
    ----------

    """
    # Ano inicial inferior ao mínimo do histórico.
    if ano<Vazoes.anoInicial:
        raise NameError("Você não pode alterar vazões de anos anteriores a {}.".format(Vazoes.anoInicial))

    # Converte o valor para inteiro.
    valorI = int(valor)

    # Posição do arquivo a escrever o(s) valor(es).
    pos = (mes-1)+(ano-Vazoes.anoInicial)*12

    # Altera valor existente no histórico.
    if ano>=Vazoes.anoInicial and ano<=Vazoes.anoFinal:
        if ((mes-1) in range(0,12)):            
            Vazoes.valores[posto][pos] = valorI
            
    # Insere novos anos no histórico.
    if ano>Vazoes.anoFinal:
        
        anosInserir = int(ano - Vazoes.anoFinal)              # número de anos a inserir
        valoresAno = [0] * (12 * anosInserir)                 # vetor de 12 meses x número de anos a inserir
        Vazoes.anoFinal = Vazoes.anoFinal + anosInserir       # altera o ano final do histórico
        
        for p in range(1,Vazoes.numPostos+1):
                Vazoes.valores[p].extend(valoresAno)        
        
        Vazoes.valores[posto][pos] = valorI



if __name__ == '__main__':

    # Lê o arquivo binário para um objeto tipo 'historicoVazoes' (ver definição da classe).
    vazoesLidas = leVazoes(nomeArquivo='VAZOES.DAT', anoInicial=1931, numPostos=320)
    
    # Teste para 'clonar classe'
    # vazoesEspelho = copy.deepcopy(vazoesLidas)
   
    # Se desejar, pode utilizar algumas informações do objeto 'historicoVazoes' ou
    # até mesmo, utilizar seu algoritmo para alterar os valores do dicionário de vazões.
    # Exemplos de uso direto:
    # a) ano_inicial = vazoesLidas.anoFinal
    # b) ano_final = vazoesLidas.anoFinal
    # c) numero_postos = vazoesLidas.numPostos
    
    # d) no caso de acesso direto ao dicionário, o valores são sequencias por posto são sequenciais.
    # Para calcular a posição de um mês/ano use a expressão para cada posto, use a expressão:
    # pos = (mes-1)+12*(ano-ano inicial do histórico)*12
    # Exemplos:
    # vazaoCamargosJan1931 = vazoesLidas.valores[1][0] 
    # vazaoCamargosDez1931 = vazoesLidas.valores[1][11]
    # vazaoCamargosJan1932 = vazoesLidas.valores[1][12]
        
    # Exemplo 1: Alterando o valore de Jan/1931 do posto Camargos.
    # Utilizando a rotina 'mudaVazao', a posição sequencial é calculada automaticamente.
    mudaVazao(vazoesLidas,posto=1,mes=1,ano=1931,valor=180)

    # Exemplo 2: Alterando os dados passados de Furnas para valores de teste.
    for ano in range(vazoesLidas.anoInicial, vazoesLidas.anoFinal+1):
        for mes in range(1,13):
            mudaVazao(vazoesLidas,6,mes, ano, mes+ano)

    # Salva arquivos de vazões alterado.
    salvaArquivo('VAZOES2.txt', vazoesLidas, tipo='vazEdit') 
    salvaArquivo('VAZOES2.csv', vazoesLidas, tipo='csv') 

    # Exemplo 3: Lendo dados do Excel para alterar o histórico de vazões.
    vazoesNovas = lerVazoesExcel('pyVazEdit_Excel.xlsx',3,2,13,14)

    for key in vazoesNovas:
        tmpList = vazoesNovas[key]
        for sl in tmpList:
            mudaVazao(vazoesLidas, key, sl[0], sl[1], sl[2]) 

    salvaArquivo('VAZOES3.txt', vazoesLidas,'vazEdit') 
    
   
   