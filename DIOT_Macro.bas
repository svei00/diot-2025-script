Attribute VB_Name = "DIOT_Macro"
Option Explicit

' ============================================================================
' Generador DIOT — macro VBA nativa (sin Python, sin instalar nada)
' Corre DENTRO del libro Registro de Facturación.
'
' Misma logica que diot_generator.py, verificada contra Resumen meses 1-7:
'   Bucket A = PUE bancarizadas (RecibidasXML, Bancarizado="Si")
'   Bucket B = Efectivo <= limite (RecibidasXML, FormaDePago="01", Total<=limite)
'   Bucket C = REP pagados en el mes (PagosRecibidasXML, BancarizadoP="Si", RelUUID="OK")
'   Redondeo Art. 20 CFF + invariante SAT: (base*16)\100 >= iva (via enteros Long)
'
' USO: Alt+F8 -> GenerarDIOT  (o asigna un boton a esta macro)
' ============================================================================

Private Const DEFAULT_OP As String = "85"
Private Const OPMAP_SHEET As String = "diot_rfc_op"   ' hoja opcional: RFC | Op

Public Sub GenerarDIOT()
    Dim mes As Long, anio As Long, tipo As String
    Dim limite As Double
    Dim usosDed As Object

    If Not PedirParametros(mes, tipo) Then Exit Sub

    anio = LeerAnioControl()
    limite = LeerLimiteControl()
    Set usosDed = LeerUsosControl()

    Dim agg As Object   ' Scripting.Dictionary  RFC -> array(base, iva, ret)
    Set agg = CreateObject("Scripting.Dictionary")

    Dim totA_base As Double, totA_iva As Double, totA_n As Long
    Dim totB_base As Double, totB_iva As Double, totB_n As Long
    Dim totC_base As Double, totC_iva As Double, totC_n As Long
    Dim descartados As String

    ProcesarRecibidas ThisWorkbook.Sheets("RecibidasXML"), mes, anio, limite, usosDed, agg, _
        totA_base, totA_iva, totA_n, totB_base, totB_iva, totB_n

    ProcesarPagos ThisWorkbook.Sheets("PagosRecibidasXML"), mes, anio, agg, _
        totC_base, totC_iva, totC_n, descartados

    Dim opMap As Object
    Set opMap = LeerOpMap()

    Dim lineas() As String, nLineas As Long
    Dim baseTot As Long, ivaTot As Long, retTot As Long, bumped As Long
    ConstruirRenglones agg, opMap, lineas, nLineas, baseTot, ivaTot, retTot, bumped

    If nLineas = 0 Then
        MsgBox "No hay proveedores con datos para " & NombreMes(mes) & " " & anio & ".", vbExclamation
        Exit Sub
    End If

    Dim nombreSugerido As String
    nombreSugerido = NombreArchivoDIOT(mes, anio, tipo)

    Dim ruta As Variant
    ruta = Application.GetSaveAsFilename( _
        InitialFileName:=nombreSugerido, _
        FileFilter:="Archivo DIOT (*.txt), *.txt", _
        Title:="Guardar DIOT como...")
    If ruta = False Then Exit Sub   ' cancelado

    EscribirArchivoUTF8 CStr(ruta), lineas, nLineas

    Dim resumen As String
    resumen = ConciliarConResumen(mes, totA_iva, totA_n, totB_iva, totB_n, totC_iva, totC_n, limite)

    Dim rfc As String
    rfc = DetectarRFC()

    Dim msg As String
    msg = "DIOT " & rfc & "  -  " & NombreMes(mes) & " " & anio & "   tipo: " & tipo & vbCrLf & _
          String(60, "=") & vbCrLf & _
          "Proveedores en TXT: " & nLineas & vbCrLf & _
          "Archivo: " & ruta & vbCrLf & _
          String(60, "-") & vbCrLf & resumen & String(60, "-") & vbCrLf & _
          "TXT TOTAL (CFF redon.)   base " & Format(baseTot, "#,##0") & "   IVA " & Format(ivaTot, "#,##0") & vbCrLf
    If retTot > 0 Then msg = msg & "IVA retenido: " & retTot & " (verificar campo 48)" & vbCrLf
    If bumped > 0 Then msg = msg & "Ajuste base invariante SAT: +" & bumped & " peso(s)" & vbCrLf
    If descartados <> "" Then msg = msg & vbCrLf & "REP descartado(s):" & vbCrLf & descartados

    MsgBox msg, vbInformation, "DIOT generado"
End Sub


' ---------- parametros ----------
Private Function PedirParametros(ByRef mes As Long, ByRef tipo As String) As Boolean
    Dim s As String
    s = InputBox("Mes a generar (1-12):", "DIOT - Mes", Month(Date))
    If s = "" Then PedirParametros = False: Exit Function
    If Not IsNumeric(s) Then
        MsgBox "Mes invalido.", vbExclamation: PedirParametros = False: Exit Function
    End If
    mes = CLng(s)
    If mes < 1 Or mes > 12 Then
        MsgBox "Mes invalido (1-12).", vbExclamation: PedirParametros = False: Exit Function
    End If

    tipo = InputBox("Tipo (N=normal, C1, C2, ...):", "DIOT - Tipo", "N")
    tipo = UCase(Trim(tipo))
    If tipo = "" Then tipo = "N"
    If tipo <> "N" And Not (Left(tipo, 1) = "C" And IsNumeric(Mid(tipo, 2))) Then
        MsgBox "Tipo invalido. Usa N o C1, C2, ...", vbExclamation: PedirParametros = False: Exit Function
    End If

    PedirParametros = True
End Function


' ---------- Control: anio, limite efectivo, usos CFDI deducibles ----------
Private Function LeerAnioControl() As Long
    On Error GoTo falla
    Dim ws As Worksheet, c As Range
    Set ws = ThisWorkbook.Sheets("Control")
    Set c = ws.Columns(8).Find("Año fiscal", LookIn:=xlValues, LookAt:=xlPart)
    If Not c Is Nothing Then
        LeerAnioControl = CLng(ws.Cells(c.Row, 9).Value)
        Exit Function
    End If
falla:
    LeerAnioControl = Year(Date)
End Function

Private Function LeerLimiteControl() As Double
    On Error GoTo falla
    Dim ws As Worksheet, c As Range
    Set ws = ThisWorkbook.Sheets("Control")
    Set c = ws.Columns(8).Find("Límite pago en efectivo", LookIn:=xlValues, LookAt:=xlPart)
    If Not c Is Nothing Then
        LeerLimiteControl = CDbl(ws.Cells(c.Row, 9).Value)
        Exit Function
    End If
falla:
    LeerLimiteControl = 2000
End Function

Private Function LeerUsosControl() As Object
    Dim usos As Object
    Set usos = CreateObject("Scripting.Dictionary")
    On Error GoTo falla
    Dim ws As Worksheet, c As Range, i As Long
    Set ws = ThisWorkbook.Sheets("Control")
    Set c = ws.Columns(8).Find("Usos CFDI deducibles", LookIn:=xlValues, LookAt:=xlPart)
    If Not c Is Nothing Then
        i = c.Row + 1
        Do While ws.Cells(i, 8).Value <> ""
            usos(CStr(ws.Cells(i, 8).Value)) = True
            i = i + 1
        Loop
    End If
falla:
    If usos.Count = 0 Then
        usos("G01") = True
        usos("G03") = True
    End If
    Set LeerUsosControl = usos
End Function

Private Function LeerOpMap() As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(OPMAP_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Set LeerOpMap = m: Exit Function

    Dim r As Long
    r = 2
    Do While ws.Cells(r, 1).Value <> ""
        m(CStr(ws.Cells(r, 1).Value)) = CStr(ws.Cells(r, 2).Value)
        r = r + 1
    Loop
    Set LeerOpMap = m
End Function

Private Function DetectarRFC() As String
    Dim m As Object, re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "[A-ZÑ&]{3,4}\d{6}[A-Z\d]{3}"
    re.Global = False
    Dim s As String
    s = ThisWorkbook.FullName
    If re.Test(s) Then
        DetectarRFC = re.Execute(s)(0).Value
    Else
        DetectarRFC = "RFC"
    End If
End Function


' ---------- helpers de fecha / texto ----------
Private Function MesDe(ByVal fecha As Variant) As Long
    If IsEmpty(fecha) Or fecha = "" Then MesDe = 0: Exit Function
    If IsDate(fecha) Then
        MesDe = Month(CDate(fecha))
    Else
        Dim s As String
        s = Trim(CStr(fecha))
        If Len(s) >= 5 Then
            On Error Resume Next
            MesDe = CLng(Mid(s, 4, 2))
            On Error GoTo 0
        End If
    End If
End Function

Private Function AnioDe(ByVal fecha As Variant) As Long
    If IsEmpty(fecha) Or fecha = "" Then AnioDe = 0: Exit Function
    If IsDate(fecha) Then
        AnioDe = Year(CDate(fecha))
    Else
        Dim s As String
        s = Trim(CStr(fecha))
        If Len(s) >= 10 Then
            On Error Resume Next
            AnioDe = CLng(Mid(s, 7, 4))
            On Error GoTo 0
        End If
    End If
End Function

Private Function EsSi(ByVal v As Variant) As Boolean
    Dim s As String
    s = LCase(Trim(CStr(v)))
    EsSi = (s = "si" Or s = "sí")
End Function

Private Function Cod2(ByVal v As Variant) As String
    Cod2 = Left(Trim(CStr(v)), 2)
End Function

Private Function Cod3(ByVal v As Variant) As String
    Cod3 = Left(Trim(CStr(v)), 3)
End Function

Private Function Num(ByVal v As Variant) As Double
    If IsNumeric(v) Then Num = CDbl(v) Else Num = 0
End Function

' ---------- encuentra el indice de columna por encabezado (con alternativas) ----------
Private Function ColPorNombre(ws As Worksheet, filaEncabezado As Long, ParamArray nombres() As Variant) As Long
    Dim ultCol As Long, i As Long, j As Long
    ultCol = ws.Cells(filaEncabezado, ws.Columns.Count).End(xlToLeft).Column
    For j = LBound(nombres) To UBound(nombres)
        For i = 1 To ultCol
            If Trim(CStr(ws.Cells(filaEncabezado, i).Value)) = CStr(nombres(j)) Then
                ColPorNombre = i
                Exit Function
            End If
        Next i
    Next j
    Err.Raise vbObjectError + 1, , "Columna no encontrada: " & Join(nombres, " / ")
End Function


' ---------- Bucket A + B: RecibidasXML ----------
Private Sub ProcesarRecibidas(ws As Worksheet, ByVal mes As Long, ByVal anio As Long, _
        ByVal limite As Double, usosDed As Object, agg As Object, _
        ByRef totA_base As Double, ByRef totA_iva As Double, ByRef totA_n As Long, _
        ByRef totB_base As Double, ByRef totB_iva As Double, ByRef totB_n As Long)

    Dim ultFila As Long
    ultFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultFila < 2 Then Exit Sub

    Dim cFecha As Long, cTipo As Long, cEstado As Long, cMetodo As Long
    Dim cBanc As Long, cUso As Long, cIva As Long, cForma As Long
    Dim cComb As Long, cRfc As Long, cRet As Long, cTotal As Long

    cFecha = ColPorNombre(ws, 1, "Fecha Emision")
    cTipo = ColPorNombre(ws, 1, "Tipo")
    cEstado = ColPorNombre(ws, 1, "Estado")
    cMetodo = ColPorNombre(ws, 1, "Metodo de Pago")
    cBanc = ColPorNombre(ws, 1, "Bancarizado")
    cUso = ColPorNombre(ws, 1, "UsoCFDI")
    cIva = ColPorNombre(ws, 1, "IVA 16%")
    cForma = ColPorNombre(ws, 1, "FormaDePago")
    cComb = ColPorNombre(ws, 1, "Combustible")
    cRfc = ColPorNombre(ws, 1, "RFC Emisor")
    cRet = ColPorNombre(ws, 1, "Retenido IVA")
    cTotal = ColPorNombre(ws, 1, "Total")

    Dim r As Long, rfc As String
    Dim iva As Double, base As Double, ret As Double
    Dim arr As Variant

    For r = 2 To ultFila
        rfc = Trim(CStr(ws.Cells(r, cRfc).Value))
        If rfc <> "" Then
            If MesDe(ws.Cells(r, cFecha).Value) = mes And AnioDe(ws.Cells(r, cFecha).Value) = anio Then
                If Trim(CStr(ws.Cells(r, cTipo).Value)) = "Factura" Then
                    If Trim(CStr(ws.Cells(r, cEstado).Value)) <> "Cancelado" Then
                        If Left(Trim(CStr(ws.Cells(r, cMetodo).Value)), 3) = "PUE" Then
                            If usosDed.Exists(Cod3(ws.Cells(r, cUso).Value)) Then

                                iva = Num(ws.Cells(r, cIva).Value)
                                If iva <> 0 Then base = iva / 0.16 Else base = 0
                                ret = Num(ws.Cells(r, cRet).Value)

                                If EsSi(ws.Cells(r, cBanc).Value) Then
                                    AcumularAgg agg, rfc, base, iva, ret
                                    totA_base = totA_base + base: totA_iva = totA_iva + iva: totA_n = totA_n + 1
                                ElseIf Cod2(ws.Cells(r, cForma).Value) = "01" _
                                        And Num(ws.Cells(r, cTotal).Value) <= limite _
                                        And Not EsSi(ws.Cells(r, cComb).Value) Then
                                    AcumularAgg agg, rfc, base, iva, ret
                                    totB_base = totB_base + base: totB_iva = totB_iva + iva: totB_n = totB_n + 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Sub


' ---------- Bucket C: PagosRecibidasXML (REP) ----------
Private Sub ProcesarPagos(ws As Worksheet, ByVal mes As Long, ByVal anio As Long, agg As Object, _
        ByRef totC_base As Double, ByRef totC_iva As Double, ByRef totC_n As Long, _
        ByRef descartados As String)

    Dim ultFila As Long
    ultFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultFila < 2 Then Exit Sub

    Dim cFPago As Long, cEstado As Long, cBancP As Long, cBase As Long, cIva As Long, cRfc As Long
    Dim cRel As Long, cDup As Long
    cFPago = ColPorNombre(ws, 1, "FechaPago")
    cEstado = ColPorNombre(ws, 1, "Estado")
    cBancP = ColPorNombre(ws, 1, "BancarizadoP")
    cBase = ColPorNombre(ws, 1, "IVA 16 Base")
    cIva = ColPorNombre(ws, 1, "IVA 16 Importe")
    cRfc = ColPorNombre(ws, 1, "RFC Emisor CFDI")

    On Error Resume Next
    cRel = ColPorNombre(ws, 1, "RelUUID")
    cDup = ColPorNombre(ws, 1, "EsDuplicado")
    On Error GoTo 0

    Dim r As Long, rfc As String, base As Double, iva As Double

    For r = 2 To ultFila
        rfc = Trim(CStr(ws.Cells(r, cRfc).Value))
        If rfc <> "" Then
            If MesDe(ws.Cells(r, cFPago).Value) = mes And AnioDe(ws.Cells(r, cFPago).Value) = anio Then
                If Trim(CStr(ws.Cells(r, cEstado).Value)) <> "Cancelado" Then
                    If EsSi(ws.Cells(r, cBancP).Value) Then
                        If cDup > 0 And Num(ws.Cells(r, cDup).Value) = 1 Then
                            ' duplicado, se descarta silenciosamente
                        ElseIf cRel > 0 And Trim(CStr(ws.Cells(r, cRel).Value)) <> "OK" Then
                            iva = Num(ws.Cells(r, cIva).Value)
                            descartados = descartados & "  " & rfc & "  RelUUID=" & _
                                ws.Cells(r, cRel).Value & "  IVA " & Format(iva, "#,##0.00") & vbCrLf
                        Else
                            base = Num(ws.Cells(r, cBase).Value)
                            iva = Num(ws.Cells(r, cIva).Value)
                            AcumularAgg agg, rfc, base, iva, 0
                            totC_base = totC_base + base: totC_iva = totC_iva + iva: totC_n = totC_n + 1
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Sub


Private Sub AcumularAgg(agg As Object, ByVal rfc As String, ByVal base As Double, ByVal iva As Double, ByVal ret As Double)
    Dim arr As Variant
    If agg.Exists(rfc) Then
        arr = agg(rfc)
        arr(0) = arr(0) + base
        arr(1) = arr(1) + iva
        arr(2) = arr(2) + ret
        agg(rfc) = arr
    Else
        agg(rfc) = Array(base, iva, ret)
    End If
End Sub


' ---------- redondeo Art. 20 CFF + invariante SAT ----------
Private Function CffRound(ByVal x As Double) As Long
    If x < 0 Then CffRound = -CffRound(-x): Exit Function
    Dim n As Long, cents As Long
    n = Int(x)
    cents = CLng(Round((x - n) * 100, 0))
    If cents >= 51 Then CffRound = n + 1 Else CffRound = n
End Function

' minimo entero base >= baseInt tal que (base*16)\100 >= ivaInt   (aritmetica entera)
Private Function AsegurarInvariante(ByVal baseInt As Long, ByVal ivaInt As Long) As Long
    If ivaInt <= 0 Then AsegurarInvariante = baseInt: Exit Function
    Dim minimo As Long
    minimo = -Int(-(ivaInt * 25) / 4)   ' ceil(iva*25/4)
    If baseInt > minimo Then AsegurarInvariante = baseInt Else AsegurarInvariante = minimo
End Function


' ---------- construir los 54 campos por RFC ----------
Private Sub ConstruirRenglones(agg As Object, opMap As Object, ByRef lineas() As String, ByRef nLineas As Long, _
        ByRef baseTot As Long, ByRef ivaTot As Long, ByRef retTot As Long, ByRef bumped As Long)

    Dim claves As Variant, rfc As Variant
    Dim n As Long
    n = agg.Count
    ReDim lineas(1 To IIf(n = 0, 1, n))
    nLineas = 0
    baseTot = 0: ivaTot = 0: retTot = 0: bumped = 0

    claves = agg.Keys
    ' orden alfabetico (como sorted() en Python)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(claves) To UBound(claves) - 1
        For j = i + 1 To UBound(claves)
            If claves(i) > claves(j) Then
                tmp = claves(i): claves(i) = claves(j): claves(j) = tmp
            End If
        Next j
    Next i

    Dim arr As Variant, baseInt As Long, ivaInt As Long, retInt As Long, ajustada As Long
    Dim op As String
    Dim campos(1 To 54) As String

    For i = LBound(claves) To UBound(claves)
        rfc = claves(i)
        arr = agg(rfc)
        baseInt = CffRound(arr(0))
        ivaInt = CffRound(arr(1))
        retInt = CffRound(arr(2))
        If Not (baseInt = 0 And ivaInt = 0 And retInt = 0) Then
            ajustada = AsegurarInvariante(baseInt, ivaInt)
            If ajustada <> baseInt Then
                bumped = bumped + (ajustada - baseInt)
                baseInt = ajustada
            End If

            If opMap.Exists(CStr(rfc)) Then op = opMap(CStr(rfc)) Else op = DEFAULT_OP

            For j = 1 To 54: campos(j) = "": Next j
            campos(1) = "04"
            campos(2) = op
            campos(3) = CStr(rfc)
            campos(12) = CStr(baseInt)
            campos(22) = CStr(ivaInt)
            If retInt <> 0 Then campos(48) = CStr(retInt)

            nLineas = nLineas + 1
            lineas(nLineas) = Join(campos, "|")

            baseTot = baseTot + baseInt
            ivaTot = ivaTot + ivaInt
            retTot = retTot + retInt
        End If
    Next i
End Sub


' ---------- nombre de archivo (misma convencion que la macro vieja / Python) ----------
Private Function NombreMes(ByVal mes As Long) As String
    Dim m() As String
    m = Split(",Enero,Febrero,Marzo,Abril,Mayo,Junio,Julio,Agosto,Septiembre,Octubre,Noviembre,Diciembre", ",")
    NombreMes = m(mes)
End Function

Private Function AbrevMes(ByVal mes As Long) As String
    Dim m() As String
    m = Split(",Ene,Feb,Mar,Abr,May,Jun,Jul,Ago,Sep,Oct,Nov,Dic", ",")
    AbrevMes = m(mes)
End Function

Private Function NombreArchivoDIOT(ByVal mes As Long, ByVal anio As Long, ByVal tipo As String) As String
    If UCase(tipo) = "N" Then
        NombreArchivoDIOT = Format(mes, "00") & ". " & AbrevMes(mes) & " " & anio & "  N DIOT Declaración.txt"
    Else
        NombreArchivoDIOT = Format(mes, "00") & ". " & AbrevMes(mes) & " " & anio & " " & UCase(tipo) & " DIOT Declaración.txt"
    End If
End Function


' ---------- escribir .txt en UTF-8 con CRLF (Print # nativo es ANSI, por eso ADODB.Stream) ----------
Private Sub EscribirArchivoUTF8(ByVal ruta As String, lineas() As String, ByVal n As Long)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2            ' texto
    stm.Charset = "utf-8"
    stm.Open

    Dim i As Long
    For i = 1 To n
        stm.WriteText lineas(i) & vbCrLf
    Next i

    ' quitar el BOM que ADODB.Stream antepone en UTF-8
    Dim bin As Object
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1
    bin.Open
    stm.Position = 0
    stm.Type = 2
    stm.CopyTo bin
    stm.Close

    Dim bytes() As Byte
    bin.Position = 0
    bytes = bin.Read
    bin.Close

    ' BOM UTF-8 = EF BB BF (3 bytes) -> se descarta
    Dim outStream As Object
    Set outStream = CreateObject("ADODB.Stream")
    outStream.Type = 1
    outStream.Open
    outStream.Write bytes
    outStream.Position = 0
    outStream.Type = 1

    Dim finalStream As Object
    Set finalStream = CreateObject("ADODB.Stream")
    finalStream.Type = 1
    finalStream.Open
    outStream.Position = 3   ' saltar BOM
    outStream.CopyTo finalStream
    outStream.Close

    finalStream.SaveToFile ruta, 2   ' adSaveCreateOverWrite
    finalStream.Close
End Sub


' ---------- conciliacion contra Resumen (igual que summary() en Python) ----------
Private Function ConciliarConResumen(ByVal mes As Long, ByVal aIva As Double, ByVal aN As Long, _
        ByVal bIva As Double, ByVal bN As Long, ByVal cIva As Double, ByVal cN As Long, _
        ByVal limite As Double) As String

    Dim ws As Worksheet
    On Error GoTo sinResumen
    Set ws = ThisWorkbook.Sheets("Resumen")

    Dim r As Long, fila As Long
    fila = 0
    For r = 3 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(r, 1).Value = mes Then fila = r: Exit For
    Next r
    If fila = 0 Then GoTo sinResumen

    Dim expA As Double, expB As Double, expC As Double
    expA = Num(ws.Cells(fila, 18).Value)   ' col R (index 18 = R en 1-based)
    expB = Num(ws.Cells(fila, 23).Value)   ' col W
    expC = Num(ws.Cells(fila, 26).Value)   ' col Z

    Dim s As String
    s = "Bucket A (PUE banc)      IVA " & Format(aIva, "#,##0.00") & _
        "   " & IIf(Abs(aIva - expA) < 0.01, "OK", "DIF " & Format(aIva - expA, "0.00")) & vbCrLf
    s = s & "Bucket B (efectivo<=" & Format(limite, "0") & ")  IVA " & Format(bIva, "#,##0.00") & _
        "   " & IIf(Abs(bIva - expB) < 0.01, "OK", "DIF " & Format(bIva - expB, "0.00")) & vbCrLf
    s = s & "Bucket C (REP/PPD)        IVA " & Format(cIva, "#,##0.00") & _
        "   " & IIf(Abs(cIva - expC) < 0.01, "OK", "DIF " & Format(cIva - expC, "0.00")) & vbCrLf
    ConciliarConResumen = s
    Exit Function

sinResumen:
    ConciliarConResumen = "Bucket A  IVA " & Format(aIva, "#,##0.00") & vbCrLf & _
                           "Bucket B  IVA " & Format(bIva, "#,##0.00") & vbCrLf & _
                           "Bucket C  IVA " & Format(cIva, "#,##0.00") & vbCrLf
End Function
