# pyVazEdit versão 0.111
Código em Python para ler e escrever arquivos binários de vazão utilizados nos modelos Newave, Decomp, Gevazp e Dessem.

Dados inseridos pelo pyVazEdit para o posto Camargos (ano 2021):

<img src="figs/pyVazEdit_Exemplo4_Camargos.png" width="700"> 

Modelo de arquivo Excel que o pyVazEdit lê para atualizar um arquivo de vazões binários:

<img src="figs/pyVazEdit_Exemplo_Excel.png" width="700"> 


## Utilização:

Considerando que o pyVazEdit foi importando para seu projeto como
```Python

import pyVazEdit as pVE

```
a sintaxe de cada função é:


1) Para ler os dados básicos dos postos de vazão (arquivo 'postos.dat'):
```Python

meusPostos = pVE.lePostos(nomeArquivo='tests/POSTOS.DAT')

```

2) Para ler os valores das MLTS (arquivo 'mlt.dat'):
```Python

minhasMLTs = pVE.leMLTS(nomeArquivo='tests/MLT.DAT',numPostos=numPostos)

```

3) Obter os valores mensais de vazões do histórico (arquivo 'vazoes.dat'):
```Python

meuHistVazoes = pVE.leVazoes(nomeArquivo='tests/vazoes_original_ONS.dat', anoInicial=1931, numPostos=numPostos)

```

4) Alterar/Inserir valores em um histórico lido:
```Python

pVE.mudaVazao(meuHistVazoes,posto=1,mes=1,ano=1931,novaVazao=180)

```
Você também poderá alterar ou inserir valores acessando diretamente o objeto 'meuHistVazoes'.

5) Salva o histórico de vazões alterado:
```Python

# Formato binário padrão do ONS:
pVE.salvaArquivo(nomeArquivo='tests/vazoes_ex_02.bin', vazoesHist=meuHistVazoes, tipoArquivo='binario')   

# Formato texto padrão do ONS (software VazEdit):
pVE.salvaArquivo(nomeArquivo='tests/vazoes_ex_02.txt', vazoesHist=meuHistVazoes, tipoArquivo='vazEdit')   

# Formato csv para abertura no Excel:
pVE.salvaArquivo(nomeArquivo='tests/vazoes_ex_02.csv', vazoesHist=meuHistVazoes, tipoArquivo='csv')       


```

## Funções já implementadas:

### lePostos:
Obtêm os dados básicos (nome, ano inicial e ano final) dos postos de um arquivo binário padrão do ONS ('postos.dat').

### leMLTS:
Lê as médias de longo termo das vazões mensais de uma arquivo binário padrão do ONS ('mlt.dat').

### leVazoes:
Lê todas as vazões mensais de um arquivo binário no padrão ONS ('vazoes.dat').

### salvaArquivo:
Salva os dados binários de vazão no arquivo especificado, utilizando um dos formatos válidos.

### mudaVazao:
Altera/inclui valores de/em um objeto 'historicoVazoes' para posterior uso/salvamento.

### lerVazoesExcel:
Lê valores de vazão de uma planilha Excel (xlsx) para atualizar um arquivo binário de vazões.



## Dependências:

struct

Se desejar utilizar a função de leitura de dados de vazão do Excel: [openpyxl](https://openpyxl.readthedocs.io/en/stable/)



## Licença:

[Ver licença](LICENSE)


## Projeto relacionado:

[NVazEdit C#](http://nrbenergia.somee.com/SoftDev/NVazEdit/NVazEdit)


## Sobre o autor:

[Meu LinkedIn](http://www.linkedin.com/in/nelsonrossibittencourt)

[Minha página de projetos](http://www.nrbenergia.somee.com)



