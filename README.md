# excel-vba-custom-functions

## 1. Introdução ##

Este repositório contém um conjunto de funções que podem se transformar em fórmulas em **VBA (Visual Basic) para o Excel** muito usadas por mim e que podem ser úteis e reaproveitáveis  

### 2. Documentação ###

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


### 3.6. Guia para Demonstração ###

* n/a


## Referências ##

* n/a
