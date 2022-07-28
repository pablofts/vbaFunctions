Attribute VB_Name = "vbafun"
Function impMtx(mtx, dirimp)
    
    Set cel = Range(dirimp)
    
    For i = 1 To UBound(mtx)
        For ii = 1 To UBound(mtx, 2)
            cel.Offset(i - 1, ii - 1).Value = mtx(i, ii)
        Next
    Next
    
    impMtx = 1
    
End Function

Function transponeRang(dirbas, filst, icolt)

    'dirbas = "Table001__Page_1"
    'filst = 2 'how many rows are going to be considered as headings
    'icolt = 4 'after which column are the headings going to be considered
    
    'dir rango a transponer
    mbas = mr(dirbas)
    Filu = UBound(mbas)
    Colu = UBound(mbas, 2)
    marea = seccionaMtx(mbas, filst + 1, Filu, icolt, Colu)
    menct = transponeMtx(seccionaMtx(mbas, 1, filst, icolt, Colu))
    mcolsrep = seccionaMtx(mbas, filst + 1, Filu, 1, icolt - 1)
    
    'numero de columnas trnaspuestas
    upboundcols = UBound(menct, 2) + UBound(mcolsrep, 2) + 1 'las no transpuestas, el encabezado transpuesto y los valores -estos últimos en una sola col-
    
    'cuántas veces se repetira mcolsrep -según la combinacion única de campos unic hallados-
    
    'matriz transpuesta
    ReDim mtrans(1 To UBound(menct) * UBound(mcolsrep), 1 To upboundcols)
    
    'repite columnas no transpuestas por cada combinación unica de encabezado
    For i = 1 To UBound(menct)
        For ii = 1 To UBound(mcolsrep, 2) 'cols
            For iii = 1 To UBound(mcolsrep) 'filas
                mtrans(iii + UBound(mcolsrep) * (i - 1), ii) = mcolsrep(iii, ii)
            Next
        Next
    Next
    
    'encabezado como etiqueta
    off = UBound(mcolsrep, 2) 'recorrido por las cols no transpuestas
    For i = 1 To UBound(menct)
        For ii = 1 To UBound(menct, 2)
            For iii = 1 To UBound(mcolsrep)
                mtrans(iii + UBound(mcolsrep) * (i - 1), off + ii) = menct(i, ii)
            Next
        Next
    Next
    
    'contenido
    off = off + UBound(menct, 2) + 1 'recorrido por las cols no transpuestas y por las columnas etiqueta del encabezado transpuesto
    For i = 1 To UBound(menct)
        For ii = 1 To UBound(marea)
            mtrans(ii + UBound(marea) * (i - 1), off) = marea(ii, i)
        Next
    Next
    
    transponeRang = mtrans
    
End Function
Function concatFilsMtx(mtx As Variant) '(1 to n, 1 to n)
    
    ReDim mconc(1 To 1, 1 To UBound(mtx))
    
    For i = 1 To UBound(mtx)
        strconc = ""
        For ii = 1 To UBound(mtx, 2)
            strconc = strconc & mtx(i, ii)
        Next
        mconc(1, i) = strconc
    Next
    
    mconct = transponeMtx(mconc)
    
    concatFilsMtx = mconct
    
End Function
Function seccionaMtx(mtx As Variant, Fili, Filu, Coli, Colu) '(1 to n, 1 to n)
    
    fildif = Fili - 1
    coldif = Coli - 1
    
    ReDim novmtx(Fili - fildif To Filu - fildif, Coli - coldif To Colu - coldif)
    
    For i = Fili To Filu
        For ii = Coli To Colu
            novmtx(i - fildif, ii - coldif) = mtx(i, ii)
        Next
    Next
    
    seccionaMtx = novmtx
End Function

Function transponeMtx(mtx As Variant) '(1 to n, 1 to n)
    Filu = UBound(mtx)
    Colu = UBound(mtx, 2)
    
    ReDim novmtx(1 To Colu, 1 To Filu)
    
    For i = 1 To Colu
        For ii = 1 To Filu
            novmtx(i, ii) = mtx(ii, i)
        Next
    Next
    
    transponeMtx = novmtx
End Function

Function unic(mtx As Variant) '(1 to n, 1 to 1)

    'marca cuáles valores son numericos
    ReDim mtip(1 To UBound(mtx), 1 To 1)
    For i = 1 To UBound(mtx)
        If IsNumeric(mtx(i, 1)) Then
            mtip(i, 1) = 1
        End If
    Next
    
    'convierte el contenido de mtx a string
    For i = 1 To UBound(mtx)
        mtx(i, 1) = CStr(mtx(i, 1))
    Next
    
    'encuentra unicos
    ReDim munic(1 To 1, 1 To 1)
    cont = 1
    For i = LBound(mtx) To UBound(mtx)
        For ii = LBound(munic, 2) To UBound(munic, 2)
            If mtx(i, 1) = munic(1, ii) Then
                pres = 1
                Exit For
            End If
        Next
        If pres = 0 Then
            ReDim Preserve munic(1 To 1, 1 To cont)
            munic(1, cont) = mtx(i, 1)
            cont = cont + 1
        End If
        pres = 0
    Next
    
    unict = transponeMtx(munic)
    
    'si era numérico, devuelve a numérico
    For i = 1 To UBound(mtx)
        For ii = 1 To UBound(unict)
            If unict(ii, 1) = mtx(i, 1) And mtip(i, 1) = 1 Then
                unict(ii, 1) = CLng(unict(ii, 1))
                Exit For
            End If
        Next
    Next
    
    unic = unict
End Function


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

Function filtrador(dirPaResult, dirCrit, criterio) 'as string, as string, as string
    'filtra en matrices hechas a partir de range.value, aun si range es una sola celda
    'solo filtra pa un criterio
    'entrega una matriz (1 to n)
    
    If Not Range(criterio).Count = 1 Then
        MsgBox ("El criterio debe referenciar una sola celda")
        Exit Function
    End If
    
    mtxPaResult = mr(dirPaResult)
    mtxcrit = mr(dirCrit)
    Crit = Range(criterio).Value
    
    If Not UBound(mtxcrit) = UBound(mtxPaResult) Then
        MsgBox ("La region 'criterio' y 'para resultado' deben ser del mismo tamaño")
        Exit Function
    End If
    
    'mtx es una matriz hotizontal
    cont = 1
    ReDim mtx(1 To 1, 1 To 1)
    For i = 1 To UBound(mtxcrit)
        If mtxcrit(i, 1) = Crit Then
            ReDim Preserve mtx(1 To 1, 1 To cont)
            mtx(1, cont) = mtxPaResult(i, 1)
            cont = cont + 1
        End If
    Next
    
    'transpone mtx
'    ReDim mtxt(1 To UBound(mtx, 2), 1 To 1)
'    For i = 1 To UBound(mtxt)
'        mtxt(i, 1) = mtx(1, i)
'    Next
    mtxt = transponeMtx(mtx)
    
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
        
        Coli = Range(dira).Column
        Fili = Range(dira).Row
        Colu = Range(dirb).Column
        Filu = Range(dirb).Row
        
        If Colu >= coltar And coltar >= Coli And Filu >= filtar And filtar >= Fili Then
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

Function nombreCol_r(dirref)
    
    Set ini = Range(dirref)
    Set ul = Range(dirref).Offset(999999).End(xlUp)
    If ul.Row = 1 Then
         Set ul = ul.Offset(1)
    End If
    
    Range(ul, ini).Name = ini.Offset(-1).Value & "_r"
    
    nombreCol_r = 1
    
End Function

Function sumaMtx(mtx As Variant) '(1 to n, 1 to 1)
    For i = 1 To UBound(mtx)
        su = su + mtx(i, 1)
    Next
    sumaMtx = su
End Function

Function buscadorMtx(mtxo, mtxb, mtxe) '(1 to n, 1 to 1)

    If Not UBound(mtxe) = UBound(mtxb) Then
        MsgBox ("Las matrices de indice y coincidir deben ser del mismo tamaño")
    End If
    
    ReDim mtxr(1 To UBound(mtxo), 1 To 1)
    
    For i = 1 To UBound(mtxo)
        For ii = 1 To UBound(mtxb)
            If mtxo(i, 1) = mtxb(ii, 1) Then
                mtxr(i, 1) = mtxe(ii, 1)
                Exit For
            End If
        Next
    Next
    
    buscadorMtx = mtxr
    
End Function

Function filtMayorMenor(dirPaResult, dirCrit, dirLinf, dirLsup) 'as string, as string, as string, as string
    'filtra en matrices hechas a partir de range.value, aun si range es una sola celda
    'filtra mayor que y menor que
    'entrega una matriz (1 to n,1 to 1)
    
    Linf = Range(dirLinf).Value * 1
    Lsup = Range(dirLsup).Value * 1
    
    'acomda límites, si hace falta
    If IsEmpty(Linf) Then
        Linf = 0
    End If
    If IsEmpty(Lsup) Then
        Lsup = 1E+21
    End If
    
    mtxPaResult = mr(dirPaResult)
    mtxcrit = mr(dirCrit)
    
    If Not UBound(mtxcrit) = UBound(mtxPaResult) Then
        MsgBox ("La region 'criterio' y 'para resultado' deben ser del mismo tamaño")
        Exit Function
    End If
    
    'mtx es una matriz hotizontal
    ReDim mtx(1 To 1, 1 To 1)
    cont = 1
    For i = 1 To UBound(mtxcrit)
        If mtxcrit(i, 1) >= Linf And mtxcrit(i, 1) <= Lsup Then
            ReDim Preserve mtx(1 To 1, 1 To cont)
            mtx(1, cont) = mtxPaResult(i, 1)
            cont = cont + 1
        End If
    Next
    
    mtxt = transponeMtx(mtx)
    
    filtMayorMenor = mtxt

End Function

Function fullAddress(rang) 'as string
    fullnm = rang.Parent.Name & "!" & rang.Address(External:=False)
    fullAddress = fullnm
End Function

Function encMax(mtx As Variant) '(1 to n, 1 to 1), entrega escalar
    
    vmax = 0
    For i = 1 To UBound(mtx)
        If mtx(i, 1) > vmax Then
            vmax = mtx(i, 1)
        End If
    Next
    
    encMax = vmax
    
End Function

Function agSum(mtx As Variant)  '(1 to n, 1 to 2)
    
    mcrit = seccionaMtx(mtx, 1, UBound(mtx), 1, 1)
    umcrit = unic(mtx)
    
    For i = 1 To UBound(umcrit)
        For ii = 1 To UBound(mcrit)
            If umcrit(i, 1) = mcrit(ii, 1) Then
                su = su + mtx(ii, 2)
            End If
        Next
        umcrit(i, 1) = su
        su = 0
    Next
    
    agSum = umcrit
    
End Function

Function mochaFilas(mtx, mfils) 'quiza arreglos pendientes

    ReDim nmtx(1 To UBound(mtx) - UBound(mfils), 1 To 1)
    
    j = 1
    For i = 1 To UBound(mtx)
        pres = 0
        For ii = 1 To UBound(mfils)
            If mfils(ii, 1) = i Then
                pres = 1
                Exit For
            End If
        Next
        If pres = 0 Then
            nmtx(j, 1) = mtx(i, 1)
            j = j + 1
        End If
    Next
    
    mochaFilas = nmtx
    
End Function

'pendiente filas colores alternados, está en min gallo

Function dirPrimFil(dir)
    
    dirfulladd = fullAddress(Range(dir))
    mbas = mr(dir)
    cols = UBound(mbas, 2)
    fils = UBound(mbas)
    
    PrimFil = Mid( _
                dirfulladd, _
                InStr(InStr(1, dirfulladd, "$") + 1, dirfulladd, "$") + 1, _
                1 _
                )
    dirPrimFil = Left(dirfulladd, Len(dirfulladd) - 1) & PrimFil
    
End Function

Function DosColores(dirbas, dircol) 'as string, as string

    mcol = mr(dircol)
    
    dirIniFil = dirPrimFil(dirbas)
    
    Set rPrimFil = Range(dirIniFil)
    
    For i = 2 To UBound(mcol)
        If mcol(i, 1) = mcol(i - 1, 1) Then
            rPrimFil.Offset(i - 1).Interior.ColorIndex = _
                rPrimFil.Offset(i - 2).Interior.ColorIndex
        Else
            If rPrimFil.Offset(i - 2).Interior.ColorIndex = 37 Then
                rPrimFil.Offset(i - 1).Interior.ColorIndex = 2
            Else
                rPrimFil.Offset(i - 1).Interior.ColorIndex = 37
            End If
        End If
    Next
End Function
