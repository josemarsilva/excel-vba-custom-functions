# excel-vba-custom-functions

## 1. Introdução ##

Este repositório contém uma biblioteca com conjunto de funções em **VBA (Visual Basic) para o Excel** que precisei customizar no Excel para atender minhas necessidades e acabei reutilizando com bastante frequencia.

Biblioteca:
* _**customSplit**(strString As String, strSeparator As String, nIdxElement As Integer)_: retorna o nIdxElement elemento do string strString separado pelo string strSeparator
* _**customSplitCount**(strString As String, strSeparator As String)_: retorna a quantidade de elementos do string strString separado pelo string strSeparator
* _**customToUTF7**(strString As String)_: retorna o string sem as acentuações da língua portuguesa
* _**customStrReptCount**(strString As strElement)_: retorna a quantidade de vezes em que o string strElement se repete dentro do string strString

PS:Na seção "3.6. Guia para Demonstração" tem o uso e explicação do uso de cada uma.

### 2. Biblioteca de funções excel-vba-custom-functions ###

### 2.1. Function customSplit(strString As String, strSeparator As String, nIdxElement As Integer) ###
```vba
Function customSplit(strString As String, strSeparator As String, nIdxElement As Integer)
    ' 2018-05-29 - https://github.com/josemarsilva/excel-vba-custom-functions - customSplit() return nIdxElement (starting with 1) of string strString using separator (strSeparator)
    Dim splitReturn As String
    ' Default
    splitReturn = ""
    
    i = 1 ' First char of string
    j = 1 ' First element index
    strElement = ""
    Do While i < Len(strString) And j <= nIdxElement
    
        ' split element between separators ...
        nElementEndsAt = InStr(i, strString, strSeparator)
        If nElementEndsAt = 0 Then
            nElementEndsAt = Len(strString) + 1
        End If
        strElement = Mid(strString, i, nElementEndsAt - i)
        
        ' Was nIdxElement found?
        If j = nIdxElement Then
            splitReturn = strElement
            ' nIdxElement found! Exit Do
            Exit Do
        End If
        
        ' Next ...
        i = nElementEndsAt + Len(strSeparator)
        j = j + 1
    
    Loop
    
    ' Return
    customSplit = splitReturn
    
End Function
```


### 2.2. Function customSplitCount(strString As String, strSeparator As String) ###
```vba
Function customSplitCount(strString As String, strSeparator As String)
    ' 2018-06-10 - https://github.com/josemarsilva/excel-vba-custom-functions - customSplitCount() return the count of elements of string (strString) using separator (strSeparator)
    Dim splitCountReturn As Integer
    ' Default
    splitCountReturn = 0
    
    i = 1 ' Start with first char of string
    Do While i < Len(strString)
    
        ' split element between separators ...
        nElementEndsAt = InStr(i, strString, strSeparator)
        
        ' Any element found starting on i position?
        If nElementEndsAt = 0 Then
            ' No element found! Exit Do
            Exit Do
        End If
        
        ' Next ...
        splitCountReturn = splitCountReturn + 1
        i = nElementEndsAt + Len(strSeparator)
    
    Loop
    
    ' Return
    If splitCountReturn > 0 Then
      splitCountReturn = splitCountReturn + 1
    End If
    customSplitCount = splitCountReturn
    
End Function
```


### 2.3. Function customToUTF7(strString As String) ###
```vba
Function customToUTF7(strString As String)
    ' 2018-05-29 - https://github.com/josemarsilva/excel-vba-custom-functions - Take accents off
    customToUTF7 = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(strString, "á", "a"), "Á", "A"), "à", "a"), "À", "A"), "ã", "a"), "Ã", "A"), "â", "a"), "Â", "A"), "é", "e"), "É", "E"), "ê", "e"), "Ê", "E"), "í", "i"), "Í", "I"), "ó", "o"), "Ó", "O"), "õ", "o"), "Õ", "O"), "ô", "o"), "Ô", "O"), "ú", "u"), "Ú", "U"), "ç", "c"), "Ç", "C")
End Function
```



### 2.4. Function customStrReptCount(strString As strElement) ###
```vba
Function customStrReptCount(strString As String, strElement As String)
    ' 2018-11-25 - https://github.com/josemarsilva/excel-vba-custom-functions - customStrReptCount() return the count of (strElements) elements inside string (strString)
    Dim nReptCountReturn As Integer
    ' Default
    nReptCountReturn = customSplitCount(strString, strElement) - 1
    ' Special cases
    If nReptCountReturn <= 0 Then
      nReptCountReturn = 0
    End If
    customStrReptCount = nReptCountReturn
    
End Function
```


## 3. Projeto ##

### 3.1. Pré-requisitos ###

* Microsoft Excel (infelizmente não funciona com OpenOffice)
* Planilha precisa _obrigatoriamente_ ser salva no formato "Pasta de trabalho habilitada para macro do Excel (.xlsm)" porque senão não pode conter macros

### 3.2. Guia para Desenvolvimento ###

* Simples. No Excel, menu "Desenvolvedor" em seguida botão de opção "Visual Basic"


### 3.3. Guia para Configuração ###

* n/a


### 3.4. Guia para Teste ###

* n/a


### 3.5. Guia para Implantação ###

* Você precisa salvar sua planilha **obrigatoriamente** no formato "Pasta de trabalho habilitada para macro do Excel (.xlsm)"

![PrintScreen-01](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-01.PNG) 

* Clique no botão "Visual Basic" sob o menu "Desenvolvedor". Se a opção de menu "Desenvolvedor" não estiver habilitado você precisa habilitá-la. Siga os seguintes passos: 1) Escolha o item de menu "Opções" no menu "Arquivo"; 2) Escolha a opção de menu lateral "Personalizar faixa de opções"; 3) Habilite a opção "Desenvolvedor" na personalização da faixa de opções
  

![PrintScreen-02](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-02.PNG) 

![PrintScreen-03](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-03.PNG) 

![PrintScreen-04](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-04.PNG) 

![PrintScreen-05](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-05.PNG) 

![PrintScreen-06](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-06.PNG) 


### 3.6. Guia para Demonstração ###

#### 3.6.1. Exemplo com customSplit() e customSplitCount() ####
* Suponha uma situação de extração de extrato de conta corrente pelo .PDF
```pdf
TEC Depósito Dinheiro 650,00
DA NET SERVIÇOS 2607613 29,80-
```
* Suponha que você quer extrair a informação de valor que está no final do string. O problema é que não há uma posição fixa. O que sabemos é que o valor é o último elemento separado por um espaço em branco " ". Você até poderia usar a formula PROCURAR() pelo espaço em branco, porém conforme pode ver no exemplo, ele pode se repetir e o pior de tudo não tem quantidade de repetições fixas. Neste problema as funções _customSplit()_ e _customSplitCount()_ podem ajudar. Com o _customSplitCount()_ conseguimos saber quantos elementos possui o string separado pelo espaço em branco e com _customSplit()_ pegamos o último elemento. Exemplo

```excel
  |            A                 |           B               | B(*) |                      C                        | C(*) |
1 |TEC Depósito Dinheiro 650,00  | =customSplitCount(A1;" ") |   4  | =customSplit(A1;" ";customSplitCount(A1;" ")) |650,00|
2 |DA NET SERVIÇOS 2607613 29,80-| =customSplitCount(A2;" ") |   5  | =customSplit(A2;" ";customSplitCount(A2;" ")) |29,80-|
```
(\*) Conteúdo da célula

#### 3.6.2. Exemplo com customToUTF7() ####
* Suponha uma situação onde você precise fazer comparação entre células extraídas de lugares diferentes. Mas um dos lugares aceita acentuação e caracteres símbolos e o outro não. Como vamos conseguir comparar "diferenciação" com "Diferenciacao"

```excel
  |      A      |      B      |               C              |   C(*)  |                        D                               | D(*)|
1 |Sistema#1    |Sistema#2    |    Sistema#1 vs Sistema#2    | #1 vs #2|             Comparacao com customToUTF7()              |     |
2 |diferenciação|DIFERENCIACAO|=SE(A2=B2;"Igual";"Diferente")|Diferente| =SE(MAIÚSCULA(customToUTF7(A2))=B2;"Igual";"Diferente")|Igual|
```


#### 3.6.3. Exemplo com file save informação da planilha ####
* Suponha uma situação onde você precise salvar uma parte especifica da planilha. Suponha que você deseja correr a planilha enquanto a coluna que indica se tem dado ou não estiver preenchida, caso afirmativo jogar o conteúdo em um arquivo. Evidente que você poderia salvar direto como texto, mas aqui  o objetivo é ensinar.

* ![excel-vba-custom-functions.xlsm](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/src/excel-vba-custom-functions.xlsm) 


#### 3.6.4. Exemplo parser de arquivo posicional ####
* Suponha uma situação onde precisa abrir um arquivo posicional para analisar seu conteúdo.

* ![file-parser.xlsm](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/src/file-parser.xlsm) 


## Referências ##

* http://excelevba.com.br/cores-no-vba/
* https://github.com/OfficeDev/Excel-Custom-Functions
