# excel-vba-custom-functions

## 1. Introdução ##

Este repositório contém uma biblioteca com conjunto de funções em **VBA (Visual Basic) para o Excel** que precisei customizar no Excel para atender minhas necessidades e acabei reutilizando com bastante frequencia.

Biblioteca:
* _**split**(strString As String, strSeparator As String, nIdxElement As Integer)_: retorna o nIdxElement elemento do string strString separado pelo string strSeparator
* _**splitCount**(strString As String, strSeparator As String)_: retorna a quantidade de elementos do string strString separado pelo string strSeparator
* _**toUTF7**(strString As String)_: retorna o string sem as acentuações da língua portuguesa
* _**strReptCount**(strString As strElement)_: retorna a quantidade de vezes em que o string strElement se repete dentro do string strString
* _**procvList**(strKeyList, strRange, keyIndex, valueIndex, strItemPrefix, strSeparator)_: retorna os valores correspondentes a lista de chaves passadas como parâmetro

Na seção [3.6. Guia para Demonstração](#36-guia-para-demonstração) tem mais explicações sobre cada uma, mas se preferir baixe os exemplos e veja você mesmo:

* [excel-vba-custom-functions.xlsm](#./src/excel-vba-custom-functions.xlsm)
* [file-parser.xlsm](#./src/file-parser.xlsm)

### 2. Biblioteca de funções excel-vba-custom-functions ###

### 2.1. Function split(strString As String, strSeparator As String, nIdxElement As Integer) ###
```vba
Function split(strString As String, strSeparator As String, nIdxElement As Integer)
    ' 2018-05-29 - https://github.com/josemarsilva/excel-vba-custom-functions - split() return nIdxElement (starting with 1) of string strString using separator (strSeparator)
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
    ' 2020-05-28 - https://github.com/josemarsilva/excel-vba-custom-functions - splitCount() return the count of elements of string (strString) using separator (strSeparator)
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
    If splitCountReturn > 0 Or Len(strString) > 0 Then
      splitCountReturn = splitCountReturn + 1
    End If
    splitCount = splitCountReturn
    
End Function
```


### 2.3. Function toUTF7(strString As String) ###
```vba
Function toUTF7(strString As String)
    ' 2018-05-29 - https://github.com/josemarsilva/excel-vba-custom-functions - toUTF7() - Take accents off
    toUTF7 = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(strString, "á", "a"), "Á", "A"), "à", "a"), "À", "A"), "ã", "a"), "Ã", "A"), "â", "a"), "Â", "A"), "é", "e"), "É", "E"), "ê", "e"), "Ê", "E"), "í", "i"), "Í", "I"), "ó", "o"), "Ó", "O"), "õ", "o"), "Õ", "O"), "ô", "o"), "Ô", "O"), "ú", "u"), "Ú", "U"), "ç", "c"), "Ç", "C")
End Function
```



### 2.4. Function strReptCount(strString As strElement) ###
```vba
Function strReptCount(strString As String, strElement As String)
    ' 2018-11-25 - https://github.com/josemarsilva/excel-vba-custom-functions - strReptCount() return the count of (strElements) elements inside string (strString)
    Dim nReptCountReturn As Integer
    ' Default
    nReptCountReturn = splitCount(strString, strElement) - 1
    ' Special cases
    If nReptCountReturn <= 0 Then
      nReptCountReturn = 0
    End If
    strReptCount = nReptCountReturn
    
End Function
```

### 2.5. Function procvList(strKeyList, strRange, keyIndex, valueIndex, strItemPrefix, strSeparator)
```vba
Function procvList(strKeyList As String, strRange As String, keyIndex As Integer, valueIndex As Integer, strItemPrefix As String, strSeparator As String)
    ' 2020-05-21 - https://github.com/josemarsilva/excel-vba-custom-functions - procvList()
    ' Initialize ...
    Dim strReturn As String
    strReturn = ""
    Set cellMatrix = range(strRange)
    ' Loop matrix
    Dim nIndex As Integer
    ' Dim strBuffer As String
    nIndex = 1
    Do While cellMatrix(nIndex, keyIndex).Value <> ""
        ' Is key ?
        If InStr((", " & strKeyList & ", "), (", " & cellMatrix(nIndex, keyIndex).Value & ", ")) Then
            If strReturn = "" Then
                strReturn = strItemPrefix & cellMatrix(nIndex, valueIndex).Value
            Else
                strReturn = strReturn & strSeparator & strItemPrefix & cellMatrix(nIndex, valueIndex).Value
            End If
        End If
        ' Next ...
        nIndex = nIndex + 1
    Loop
    ' Return
    procvList = strReturn
    
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

* Clique no botão "Visual Basic" sob o menu "Desenvolvedor". 
* Se a opção de menu "Desenvolvedor" não estiver habilitado Então você precisa habilitá-la. Siga os seguintes passos:
  * Escolha o item de menu "Opções" no menu "Arquivo"; 
  * Escolha a opção de menu lateral "Personalizar faixa de opções"; 
  * Habilite a opção "Desenvolvedor" na personalização da faixa de opções
  

![PrintScreen-02](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-02.PNG) 

![PrintScreen-03](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-03.PNG) 

![PrintScreen-04](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-04.PNG) 

![PrintScreen-05](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-05.PNG) 

![PrintScreen-06](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/doc/PrintScreen-06.PNG) 


### 3.6. Guia para Demonstração ###

#### 3.6.1. Exemplo com split() e splitCount() ####
* Suponha uma situação de extração de extrato de conta corrente pelo .PDF

![PrintScreen-07](./doc/PrintScreen-07.PNG) 


* Suponha que você quer extrair a informação de valor que está no final do string. O problema é que não há uma posição fixa. O que sabemos é que o valor é o último elemento separado por um espaço em branco " ". Você até poderia usar a formula PROCURAR() pelo espaço em branco, porém conforme pode ver no exemplo, ele pode se repetir e o pior de tudo não tem quantidade de repetições fixas. Neste problema as funções _split()_ e _splitCount()_ podem ajudar. Com o _splitCount()_ conseguimos saber quantos elementos possui o string separado pelo espaço em branco e com _split()_ pegamos o último elemento. Exemplo

#### 3.6.2. Exemplo com toUTF7() ####
* Suponha uma situação onde você precise fazer comparação entre células extraídas de lugares diferentes. Mas um dos lugares aceita acentuação e caracteres símbolos e o outro não. Como vamos conseguir comparar "diferenciação" com "Diferenciacao"

![PrintScreen-08](./doc/PrintScreen-08.PNG) 

#### 3.6.3. Exemplo com `file Save ...` e `file Open ...` dialog boxes informação da planilha ####
* Suponha uma situação onde você precise salvar uma parte especifica da planilha. Suponha que você deseja correr a planilha enquanto a coluna que indica se tem dado ou não estiver preenchida, caso afirmativo jogar o conteúdo em um arquivo. Evidente que você poderia salvar direto como texto, mas aqui  o objetivo é ensinar.

![PrintScreen-09](./doc/PrintScreen-09.PNG) 


#### 3.6.4. Exemplo com PROCVLIST()
* Suponha que você tenha uma lista de chaves que precise recuperar o valor

![PrintScreen-10](./doc/PrintScreen-10.PNG) 


#### 3.6.5. Exemplo parser de arquivo posicional ####
* Suponha uma situação onde precisa abrir um arquivo posicional para analisar seu conteúdo.

* ![file-parser.xlsm](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/src/file-parser.xlsm) 


## Referências ##

* http://excelevba.com.br/cores-no-vba/
* https://github.com/OfficeDev/Excel-Custom-Functions
