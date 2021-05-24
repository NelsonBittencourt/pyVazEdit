# -*- coding: utf-8 -*-

"""
******************************************************************************
Exemplos de uso do 'pyVazEdit'.

Autor   : Nelson Rossi Bittencourt
Versão  : 0.1
Licença : MIT
Dependências: pyVazEdit
******************************************************************************
"""

import pyVazEdit as pVE

if __name__ == '__main__':

    # Lê os dados de um arquivo binário para um objeto tipo 'historicoVazoes' (ver definição da classe
    # no código fonte do 'pyVazEdit').
    meuHistVazoes = pVE.leVazoes(nomeArquivo='tests/vazoes_original_ONS.dat', anoInicial=1931, numPostos=320)

    # Se desejar, pode utilizar algumas informações do objeto 'historicoVazoes' ou
    # até mesmo, utilizar códigos diferentes alterar os valores do dicionário de vazões.
    # Exemplos de uso direto:
    # a) ano_inicial = meuHistVazoes.anoFinal
    # b) ano_final = meuHistVazoes.anoFinal
    # c) numero_postos = meuHistVazoes.numPostos    
    # d) no caso de acesso direto ao dicionário, as vazões sequencias por posto.
    # A posição de um determinado mês/ano para cada posto, segue a seguinte expressão:
    # pos = (mes-1)+12*(ano-ano inicial do histórico)*12
    # Exemplos:
    # vazaoCamargosJan1931 = meuHistVazoes.valores[1][0] 
    # vazaoCamargosDez1931 = meuHistVazoes.valores[1][11]
    # vazaoCamargosJan1932 = meuHistVazoes.valores[1][12]
    # vazaoCamargosJanPosto320 = meuHistVazoes.valores[320][0]
        
    # Exemplo 1: Alterando o valore de Jan/1931 do posto Camargos.
    # Utilizando a rotina 'mudaVazao' a posição sequencial é calculada automaticamente.
    pVE.mudaVazao(meuHistVazoes,posto=1,mes=1,ano=1931,novaVazao=180)
    pVE.salvaArquivo(nomeArquivo='tests/vazoes_ex_01.dat',vazoesHist=meuHistVazoes)
    

    # Exemplo 2: Alterando todos os dados passados de Furnas (posto número 6) para valores de teste.
    for ano in range(meuHistVazoes.anoInicial, meuHistVazoes.anoFinal+1):
        for mes in range(1,13):
            pVE.mudaVazao(meuHistVazoes,6,mes, ano, mes+ano)

    # Salva arquivos de vazão com valores alterados.
    # Utilize o argumento 'tipo' para especificar o tipo de arquivo a ser salvo.
    # Se omitir o 'tipo', será considerado como 'binario'.
    pVE.salvaArquivo(nomeArquivo='tests/vazoes_ex_02.bin', vazoesHist=meuHistVazoes, tipoArquivo='binario')   # Formato binário    
    pVE.salvaArquivo(nomeArquivo='tests/vazoes_ex_02.txt', vazoesHist=meuHistVazoes, tipoArquivo='vazEdit')   # Formato texto compatível com o VazEdit
    pVE.salvaArquivo(nomeArquivo='tests/vazoes_ex_02.csv', vazoesHist=meuHistVazoes, tipoArquivo='csv')       # Formato separado por vírgulas


    # Exemplo 3: Lendo dados do Excel para alterar o histórico de vazões.
    vazoesNovasExcel = pVE.lerVazoesExcel('tests/pyVazEdit_Excel.xlsx',3,2,13,14)
       
    for key in vazoesNovasExcel:
        tmpList = vazoesNovasExcel[key]
        for sl in tmpList:
            pVE.mudaVazao(meuHistVazoes, key, sl[0], sl[1], sl[2]) 

    pVE.salvaArquivo('tests/vazoes_ex_03.txt', meuHistVazoes,'vazEdit') 


    # Exemplo 4: Reabre o arquivo original do ONS, calcula o valor médio das vazões 
    # de Camargos e lança esse valor para 2022.
    meuHistVazoes = pVE.leVazoes(nomeArquivo='tests/vazoes_original_ONS.dat')

    aux = meuHistVazoes.valores[1]
    denominador = sum(x > 0 for x in aux)
    mltCamargos = int(sum(aux)/denominador)

    for m in range(1,13):
        pVE.mudaVazao(meuHistVazoes,1,m,2022,mltCamargos)
    
    pVE.salvaArquivo(
                    nomeArquivo="tests/vazoes_camargos.txt",
                    vazoesHist=meuHistVazoes,
                    tipoArquivo='vazEdit')