# excel-vba-custom-functions

## 1. Introdução ##

Este repositório contém uma biblioteca com conjunto de funções em **VBA (Visual Basic) para o Excel** que precisei customizar no Excel para atender minhas necessidades e acabei reutilizando com bastante frequencia.

Na seção "3.6. Guia para Demonstração" tem o uso e explicação do uso de cada uma.

### 2. Biblioteca de funções excel-vba-custom-functions ###

### 2.1. Function split(strString As String, strSeparator As String, nIdxElement As Integer) ###
```vba
Function split(strString As String, strSeparator As String, nIdxElement As Integer)
    ' 2018-05-29 - Josemar Silva - split() return nIdxElement (starting with 1) of string strString using separator (strSeparator)
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
    ' 2018-06-10 - Josemar Silva - splitCount() return the count of elements of string (strString) using separator (strSeparator)
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

![PrintScreen-01](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/PrintScreen-01.PNG) 

![PrintScreen-02](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/PrintScreen-02.PNG) 

![PrintScreen-03](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/PrintScreen-03.PNG) 

![PrintScreen-04](https://github.com/josemarsilva/excel-vba-custom-functions/blob/master/PrintScreen-04.PNG) 


### 3.6. Guia para Demonstração ###

* n/a


## Referências ##

* n/a
