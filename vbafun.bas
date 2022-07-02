Attribute VB_Name = "funciones_varias"
Function mr(rangodir) 'As String
    'hace matriz a partir de range.value, aun range es una sola celda
    'entrega una matriz (1 to n, 1 to 1)
    
    On Error GoTo simal
    
    If Range(rangodir).Count = 1 Then
         ReDim mtx(1 To 1, 1 To 1)
         mtx(1, 1) = Range(rangodir).Value
    Else
        mtx = Range(rangodir).Value
    End If
    
    mr = mtx
    
    If nada = 1 Then 'candado pal goto
simal:
    MsgBox ("no es un rango valido")
    Exit Function
    End If
End Function

Function buscador(dirOr, dirBus, dirPaResult)  'As String As String As String
    'busca en matrices hechas a partir de range.value, aun si range es una sola celda
    'entrega una matriz (1 to n, 1 to 1)
    
    mtxOr = mr(dirOr)
    mtxBus = mr(dirBus)
    mtxPaResult = mr(dirPaResult)
    
    If Not UBound(mtxBus) = UBound(mtxPaResult) Then
        MsgBox ("La region de busqueda y para resultado deben ser del mismo tamaño")
        Exit Function
    End If
    
    szOr = UBound(mtxOr)
    szDos = UBound(mtxBus)
    
    ReDim mtxResult(1 To szOr, 1 To 1) '(1 to 1, 1 to 1) nomás pa mantener las dimensiones esperadas
    
    For i = 1 To szOr
        For ii = 1 To szDos
            If mtxOr(i, 1) = mtxBus(ii, 1) Then
                mtxResult(i, 1) = mtxPaResult(ii, 1)
            End If
        Next
    Next
    
    buscador = mtxResult
End Function

Function filtrador(dirPaResult, dirCrit, criterio) 'as string, as string, as variant
    'filtra en matrices hechas a partir de range.value, aun si range es una sola celda
    'solo filtra pa un criterio
    'entrega una matriz (1 to n) con una fila en blanco
    
    If Not Range(criterio).Count = 1 Then
        MsgBox ("El criterio debe referenciar una sola celda")
        Exit Function
    End If
    
    mtxPaResult = mr(dirPaResult)
    mtxCrit = mr(dirCrit)
    crit = Range(criterio).Value
    
    If Not UBound(mtxCrit) = UBound(mtxPaResult) Then
        MsgBox ("La region 'criterio' y 'para resultado' deben ser del mismo tamaño")
        Exit Function
    End If
    
    'mtx es una matriz hotizontal
    ReDim mtx(1 To 1, 1 To 1)
    For i = 1 To UBound(mtxCrit)
        If mtxCrit(i, 1) = crit Then
            mtx(1, UBound(mtx, 2)) = mtxPaResult(i, 1)
            ReDim Preserve mtx(1 To 1, 1 To UBound(mtx, 2) + 1)
        End If
    Next
    
    'transpone mtx
    ReDim mtxt(1 To UBound(mtx, 2), 1 To 1)
    For i = 1 To UBound(mtxt)
        mtxt(i, 1) = mtx(1, i)
    Next
    
    filtrador = mtxt

End Function

Function arrayToStr(mtx As Variant)
    'convierte un array (1 to 1, 1 to 1) en un lista separada por comas -string-
    For i = 1 To UBound(mtx)
        lista = lista & mtx(i, 1) & ","
    Next
    'quita la ultima coma
    lista = Left(lista, Len(lista) - 1)
    
    arrayToStr = lista
End Function

Function valDesdeLista(StrRango, lista) 'ambas podrían as string
    With Range(StrRango).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=lista
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    valDesdeLista = 1
End Function

Function valMaximo(StrRango, valor) 'de cero a max
    With Range(StrRango).Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween, Formula1:="0", Formula2:=valor
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Error"
        .ErrorMessage = "Excede las existencias"
        .ShowInput = True
        .ShowError = True
    End With
End Function

Function dentroRango(dirTarget, dirRango) 'as string, dirtarget debe referenciar una sola celda
    
    dentroRango = 0
    
    If Not Range(dirTarget).Count = 1 Then 'si dirtarget referencia muchas celdas, abandona la función
        Exit Function
    End If
        
    'reconoce por las coordenadas, independientemente de la hoja de trabajo
    coltar = Range(dirTarget).Column
    filtar = Range(dirTarget).Row
    
    If Range(dirRango).Count = 1 Then 'si se trata de una sola celda, nomás compara fila y columna
        col = Range(dirRango).Column
        fil = Range(dirRango).Row
        
        If fil = filtar And col = coltar Then
            dentroRango = 1
        End If
    Else 'si más de una celda, primera fila y columna y última fila y columna
        dira = Mid(dirRango, 1, InStr(1, dirRango, ":") - 1)
        dirb = Mid(dirRango, InStr(1, dirRango, ":") + 1, 30)
        
        coli = Range(dira).Column
        fili = Range(dira).Row
        colu = Range(dirb).Column
        filu = Range(dirb).Row
        
        If colu >= coltar And coltar >= coli And filu >= filtar And filtar >= fili Then
            dentroRango = 1
        End If
    End If
End Function

Function idConsecutivo(dirIds) 'as string
    ids = mr(dirIds)
    For i = 1 To UBound(ids)
        If novid <= ids(i, 1) Then
            novid = ids(i, 1) + 1
        End If
    Next
    
    idConsecutivo = novid
End Function

