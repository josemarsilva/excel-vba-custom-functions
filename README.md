# excel-vba-custom-functions

## 1. Introdução ##

Este repositório contém uma biblioteca com conjunto de funções em **VBA (Visual Basic) para o Excel** que precisei customizar no Excel para atender minhas necessidades e acabei reutilizando com bastante frequencia.

Biblioteca:
* _**split**(strString As String, strSeparator As String, nIdxElement As Integer)_: retorna o nIdxElement elemento do string strString separado pelo string strSeparator
* _**splitCount**(strString As String, strSeparator As String)_: retorna a quantidade de elementos do string strString separado pelo string strSeparator
* _**toUTF7**(strString As String)_: retorna o string sem as acentuações da língua portuguesa

Na seção "3.6. Guia para Demonstração" tem o uso e explicação do uso de cada uma.

### 2. Biblioteca de funções excel-vba-custom-functions ###

### 2.1. Function split(strString As String, strSeparator As String, nIdxElement As Integer) ###
```vba
Function split(strString As String, strSeparator As String, nIdxElement As Integer)
    ' 2018-05-29 - github.com/josemarsilva - split() return nIdxElement (starting with 1) of string strString using separator (strSeparator)
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
    split = splitReturn
    
End Function
```


### 2.2. Function splitCount(strString As String, strSeparator As String) ###
```vba
Function splitCount(strString As String, strSeparator As String)
    ' 2018-06-10 - github.com/josemarsilva - splitCount() return the count of elements of string (strString) using separator (strSeparator)
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
    splitCount = splitCountReturn
    
End Function
```


### 2.3. Function toUTF7(strString As String) ###
```vba
Function toUTF7(strString As String)
    ' 2018-05-29 - github.com/josemarsilva - Take accents off
    toUTF7 = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(strString, "á", "a"), "Á", "A"), "à", "a"), "À", "A"), "ã", "a"), "Ã", "A"), "â", "a"), "Â", "A"), "é", "e"), "É", "E"), "ê", "e"), "Ê", "E"), "í", "i"), "Í", "I"), "ó", "o"), "Ó", "O"), "õ", "o"), "Õ", "O"), "ô", "o"), "Ô", "O"), "ú", "u"), "Ú", "U"), "ç", "c"), "Ç", "C")
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

* "Pasta de trabalho habilitada para macro do Excel (.xlsm)"

![PrintScreen-01](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-01.PNG) 

![PrintScreen-02](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-02.PNG) 

![PrintScreen-03](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-03.PNG) 

![PrintScreen-04](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-04.PNG) 

![PrintScreen-05](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-05.PNG) 

![PrintScreen-06](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-06.PNG) 


### 3.6. Guia para Demonstração ###

#### a. Exemplo com split() e splitCount() ####
* Suponha uma situação de extração de extrato de conta corrente pelo .PDF
```pdf
TEC Depósito Dinheiro 650,00
DA NET SERVIÇOS 2607613 29,80-
```
* Suponha que você quer extrair a informação de valor que está no final do string. O problema é que não há uma posição fixa. O que sabemos é que o valor é o último elemento separado por um espaço em branco " ". Você até poderia usar a formula PROCURAR() pelo espaço em branco, porém conforme pode ver no exemplo, ele pode se repetir e o pior de tudo não tem quantidade de repetições fixas. Neste problema as funções _split()_ e _splitCount()_ podem ajudar. Com o _splitCount()_ conseguimos saber quantos elementos possui o string separado pelo espaço em branco e com _split()_ pegamos o último elemento. Exemplo

```excel
  |            A                 |        B            | B(*) |                 C                 | C(*) |
1 |TEC Depósito Dinheiro 650,00|   =splitCount(A1;" ") |   4  | =split(A1;" ";splitCount(A1;" ")) |650,00|
2 |DA NET SERVIÇOS 2607613 29,80-| =splitCount(A2;" ") |   5  | =split(A2;" ";splitCount(A2;" ")) |29,80-|
```
(\*) Conteúdo da célula

#### b. Exemplo com toUTF7() ####
* Suponha uma situação onde você precise fazer comparação entre células extraídas de lugares diferentes. Mas um dos lugares aceita acentuação e caracteres símbolos e o outro não. Como vamos conseguir comparar "diferenciação" com "Diferenciacao"

```excel
  |      A      |      B      |               C              |   C(*)  |                        D                         | D(*)|
1 |Sistema#1    |SISTEMA#2    |    Sistema#1 vs Sistema#2    | #1 vs #2|             Comparacao com toUTF7()              |     |
2 |diferenciação|DIFERENCIACAO|=SE(A2=B2;"Igual";"Diferente")|Diferente| =SE(MAIÚSCULA(toUTF7(A2))=B2;"Igual";"Diferente")|Igual|
```


## Referências ##

* n/a
