Attribute VB_Name = "DIOT_Macro"
Option Explicit

' ============================================================================
' Generador DIOT - macro VBA nativa (sin Python, sin instalar nada)
' Corre DENTRO del libro Registro de Facturacion.
'
' Esta macro ESPEJA el paquete Python  diot/  (esa es la fuente de verdad).
' Si una regla cambia, se cambia primero en Python y despues aqui.
'
'   Bucket A = PUE bancarizadas (RecibidasXML, Bancarizado="Si", Total > 0)
'   Bucket B = Efectivo <= limite (RecibidasXML, FormaDePago="01", Total<=limite)
'   Bucket C = REP pagados en el mes (PagosRecibidasXML, BancarizadoP="Si",
'              RelUUID="OK", pago en MXN). Los importes del documento se
'              dividen entre EquivalenciaDR (vacia o 0 = 1).
'   Redondeo por proveedor Art. 20 CFF + invariante SAT para 16% y 8%:
'              (base*tasa)\100 >= iva   (aritmetica entera, Long)
'
' Campos del TXT que llena (54 campos separados por "|", UTF-8, CRLF):
'    1  Tipo de tercero ("04")          2  Tipo de operacion (default "85")
'    3  RFC del proveedor
'    8  Valor de actos region fronteriza NORTE 8%
'   12  Valor de actos tasa 16%
'   18  IVA acreditable region fronteriza NORTE 8%
'   22  IVA acreditable tasa 16%
'   48  IVA retenido (facturas + retenciones de REP)
'   50  Actos exentos (colegiaturas de facturas A + exento de REP)
'   51  Actos tasa 0% (SOLO de REP; el 0% de facturas se captura a mano)
'
' PENDIENTE: no probada en Excel tras la actualizacion de nov-2026; correr
' octubre 2026 y comparar el TXT contra el de Python.
'
' USO: Alt+F8 -> GenerarDIOT  (o asigna un boton a esta macro)
' ============================================================================

Private Const DEFAULT_OP As String = "85"
Private Const OPMAP_SHEET As String = "diot_rfc_op"   ' hoja opcional: RFC | Op

' Posiciones dentro de las "bolsas" de importes (arreglos de Double / Long).
' Una bolsa guarda los importes de un documento, de un proveedor o de un bucket.
Private Const IDX_BASE16 As Long = 0
Private Const IDX_IVA16 As Long = 1
Private Const IDX_BASE8 As Long = 2
Private Const IDX_IVA8 As Long = 3
Private Const IDX_EXENTO As Long = 4
Private Const IDX_TASA0 As Long = 5
Private Const IDX_RET As Long = 6
Private Const IDX_N As Long = 7          ' solo en totales de bucket: # de documentos
Private Const ULT_IDX_IMPORTE As Long = 6

' Columnas (numero de columna, 0 = la columna opcional no existe) de RecibidasXML
Private Type ColumnasFactura
    fecha As Long
    tipo As Long
    estado As Long
    metodo As Long
    banc As Long
    uso As Long
    iva16 As Long
    forma As Long
    comb As Long
    rfc As Long
    ret As Long
    total As Long
    subtotal As Long            ' opcional
    descuento As Long           ' opcional
    iva8 As Long                ' opcional
    totalTrasladados As Long    ' opcional
    nivelEducativo As Long      ' opcional
    conceptos As Long           ' opcional
    lugar As Long               ' opcional (LugarDeExpedicion)
    uuid As Long                ' opcional
End Type

' Columnas de PagosRecibidasXML
Private Type ColumnasRep
    fechaPago As Long
    estado As Long
    bancP As Long
    base16 As Long
    iva16 As Long
    rfc As Long
    relUuid As Long             ' opcional
    esDuplicado As Long         ' opcional
    monedaP As Long             ' opcional
    equivalencia As Long        ' opcional (EquivalenciaDR)
    impPagado As Long           ' opcional
    docRelacionado As Long      ' opcional (IdDocumento Relacionado)
    metodoDR As Long            ' opcional (MetodoDePagoDR)
    base8 As Long               ' opcional
    iva8 As Long                ' opcional
    baseExento As Long          ' opcional
    baseCero As Long            ' opcional
    lugar As Long               ' opcional (Lugar de Expedicion CFDI)
    retenciones As Collection   ' columnas "Ret IVA * Importe" (puede estar vacia)
End Type


' ============================================================================
' PROGRAMA PRINCIPAL
' ============================================================================
Public Sub GenerarDIOT()
    Dim mes As Long, anio As Long, tipo As String
    Dim limite As Double
    Dim usosDed As Object

    On Error GoTo ManejarError

    If Not PedirParametros(mes, tipo) Then Exit Sub

    anio = LeerAnioControl()
    limite = LeerLimiteControl()
    Set usosDed = LeerUsosControl()

    Dim agg As Object   ' Scripting.Dictionary  RFC -> arreglo de 7 importes (IDX_*)
    Set agg = CreateObject("Scripting.Dictionary")

    Dim avisos As Object   ' Dictionary: texto del aviso -> True (sin repetidos)
    Set avisos = CreateObject("Scripting.Dictionary")

    ' Totales por bucket (para conciliar contra Resumen): 7 importes + # documentos
    Dim totA(0 To 7) As Double, totB(0 To 7) As Double, totC(0 To 7) As Double
    Dim descartados As String

    Dim wsRec As Worksheet, wsPag As Worksheet
    Set wsRec = ThisWorkbook.Sheets("RecibidasXML")
    Set wsPag = ThisWorkbook.Sheets("PagosRecibidasXML")

    Dim colF As ColumnasFactura, colR As ColumnasRep
    LeerColumnasFactura wsRec, colF
    LeerColumnasRep wsPag, colR

    ' UUID -> metodo de pago de las facturas (para detectar REP que paga una PUE)
    Dim metodoPorUuid As Object
    Set metodoPorUuid = ConstruirMetodoPorUuid(wsRec, colF)

    ProcesarRecibidas wsRec, colF, mes, anio, limite, usosDed, agg, totA, totB, avisos
    ProcesarPagos wsPag, colR, mes, anio, metodoPorUuid, agg, totC, descartados, avisos

    Dim opMap As Object
    Set opMap = LeerOpMap()

    Dim lineas() As String, nLineas As Long
    Dim totTxt(0 To 6) As Long          ' suma de cada campo del TXT (pesos enteros)
    Dim ajuste16 As Long, ajuste8 As Long
    ConstruirRenglones agg, opMap, lineas, nLineas, totTxt, ajuste16, ajuste8

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
    resumen = ConciliarConResumen(mes, totA, totB, totC, limite)

    Dim rfc As String
    rfc = DetectarRFC()

    Dim msg As String
    msg = "DIOT " & rfc & "  -  " & NombreMes(mes) & " " & anio & "   tipo: " & tipo & vbCrLf & _
          String(60, "=") & vbCrLf & _
          "Proveedores en TXT: " & nLineas & vbCrLf & _
          "Archivo: " & ruta & vbCrLf & _
          String(60, "-") & vbCrLf & resumen & String(60, "-") & vbCrLf & _
          TextoTotalesTxt(totTxt)
    If ajuste16 > 0 Then msg = msg & "Ajuste base 16% invariante SAT: +" & ajuste16 & " peso(s)" & vbCrLf
    If ajuste8 > 0 Then msg = msg & "Ajuste base 8% invariante SAT: +" & ajuste8 & " peso(s)" & vbCrLf
    If descartados <> "" Then msg = msg & vbCrLf & "REP descartado(s):" & vbCrLf & descartados
    msg = msg & TextoDeAvisos(avisos)
    msg = msg & vbCrLf & "NOTA: la tasa 0% de FACTURAS no se deriva del export; el campo 51 solo trae " & _
          "la de REP. La tasa 0% de facturas se captura a mano en el portal." & vbCrLf

    MostrarMensajeLargo msg, "DIOT generado"
    Exit Sub

ManejarError:
    MsgBox "Error en la macro DIOT: " & Err.Description, vbCritical, "DIOT"
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


' ============================================================================
' HELPERS DE FECHA / TEXTO / NUMERO
' ============================================================================

' Las fechas del libro son TEXTO dd/mm/aaaa. NUNCA se usa DateValue/CDate sobre
' texto (depende de la configuracion regional y confunde dia con mes sin error).
' Se separa con Mid/Left. Solo un valor que ya es fecha real o serial se
' interpreta como fecha.
Private Sub LeerAnioMes(ByVal fecha As Variant, ByRef anio As Long, ByRef mes As Long)
    anio = 0
    mes = 0
    If IsError(fecha) Or IsNull(fecha) Or IsEmpty(fecha) Then Exit Sub
    If VarType(fecha) = vbBoolean Then Exit Sub

    If VarType(fecha) = vbDate Then                   ' fecha real de Excel
        anio = Year(fecha)
        mes = Month(fecha)
    ElseIf VarType(fecha) = vbString Then             ' texto dd/mm/aaaa (lo normal)
        LeerAnioMesDeTexto CStr(fecha), anio, mes
    ElseIf IsNumeric(fecha) Then                      ' numero serial de Excel
        anio = Year(CDate(fecha))
        mes = Month(CDate(fecha))
    End If
End Sub

Private Sub LeerAnioMesDeTexto(ByVal texto As String, ByRef anio As Long, ByRef mes As Long)
    Dim s As String
    s = Trim(texto)
    If Len(s) < 10 Then Exit Sub
    If (Mid(s, 3, 1) = "/" Or Mid(s, 3, 1) = "-") Then
        ' dd/mm/aaaa : dia = 1-2, mes = 4-5, anio = 7-10
        mes = CLng(Val(Mid(s, 4, 2)))
        anio = CLng(Val(Mid(s, 7, 4)))
    ElseIf Mid(s, 5, 1) = "-" Then
        ' aaaa-mm-dd (ISO) : anio = 1-4, mes = 6-7
        anio = CLng(Val(Left(s, 4)))
        mes = CLng(Val(Mid(s, 6, 2)))
    End If
End Sub

Private Function MesDe(ByVal fecha As Variant) As Long
    Dim a As Long, m As Long
    LeerAnioMes fecha, a, m
    MesDe = m
End Function

Private Function AnioDe(ByVal fecha As Variant) As Long
    Dim a As Long, m As Long
    LeerAnioMes fecha, a, m
    AnioDe = a
End Function

Private Function EsSi(ByVal v As Variant) As Boolean
    If VarType(v) = vbBoolean Then EsSi = v: Exit Function
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

' Texto de una celda sin espacios; vacio si la columna no existe (col = 0).
Private Function TextoCelda(ws As Worksheet, ByVal fila As Long, ByVal col As Long) As String
    If col = 0 Then TextoCelda = "": Exit Function
    TextoCelda = Trim(CStr(ws.Cells(fila, col).Value))
End Function

' Numero de una celda; 0 si la columna no existe (col = 0).
Private Function NumCelda(ws As Worksheet, ByVal fila As Long, ByVal col As Long) As Double
    If col = 0 Then NumCelda = 0: Exit Function
    NumCelda = Num(ws.Cells(fila, col).Value)
End Function

' Valor crudo de una celda; Empty si la columna no existe (col = 0).
Private Function ValorCelda(ws As Worksheet, ByVal fila As Long, ByVal col As Long) As Variant
    If col = 0 Then ValorCelda = Empty: Exit Function
    ValorCelda = ws.Cells(fila, col).Value
End Function


' ============================================================================
' UBICAR COLUMNAS POR ENCABEZADO
' ============================================================================

' Columna OBLIGATORIA: prueba cada nombre (sin importar mayusculas). Si ninguno
' existe, detiene la macro con un error claro (nunca lee una columna al azar).
Private Function ColPorNombre(ws As Worksheet, filaEncabezado As Long, ParamArray nombres() As Variant) As Long
    Dim ultCol As Long, i As Long, j As Long
    ultCol = ws.Cells(filaEncabezado, ws.Columns.Count).End(xlToLeft).Column
    For j = LBound(nombres) To UBound(nombres)
        For i = 1 To ultCol
            If StrComp(Trim(CStr(ws.Cells(filaEncabezado, i).Value)), CStr(nombres(j)), vbTextCompare) = 0 Then
                ColPorNombre = i
                Exit Function
            End If
        Next i
    Next j
    Err.Raise vbObjectError + 1, , "Columna no encontrada en '" & ws.Name & "': " & Join(nombres, " / ")
End Function

' Columna OPCIONAL: nombre exacto (sin importar mayusculas). Regresa 0 si no
' existe, y entonces su importe vale 0. No hay coincidencia parcial a proposito.
Private Function ColOpcional(ws As Worksheet, ByVal filaEncabezado As Long, ByVal nombre As String) As Long
    Dim ultCol As Long, i As Long
    ultCol = ws.Cells(filaEncabezado, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To ultCol
        If StrComp(Trim(CStr(ws.Cells(filaEncabezado, i).Value)), nombre, vbTextCompare) = 0 Then
            ColOpcional = i
            Exit Function
        End If
    Next i
    ColOpcional = 0
End Function

' Todas las columnas "Ret IVA <tasa> Importe" del REP (4, 10.667, 2, 16...).
Private Function ColumnasRetencionIva(ws As Worksheet) As Collection
    Dim cols As Collection
    Set cols = New Collection
    Dim ultCol As Long, i As Long, enc As String
    ultCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To ultCol
        enc = LCase(Trim(CStr(ws.Cells(1, i).Value)))
        If Left(enc, 8) = "ret iva " And Right(enc, 8) = " importe" Then cols.Add i
    Next i
    Set ColumnasRetencionIva = cols
End Function

Private Sub LeerColumnasFactura(ws As Worksheet, c As ColumnasFactura)
    c.fecha = ColPorNombre(ws, 1, "Fecha Emision", "Fecha")
    c.tipo = ColPorNombre(ws, 1, "Tipo")
    c.estado = ColPorNombre(ws, 1, "Estado")
    c.metodo = ColPorNombre(ws, 1, "Metodo de Pago", "Metodo")
    c.banc = ColPorNombre(ws, 1, "Bancarizado")
    c.uso = ColPorNombre(ws, 1, "UsoCFDI")
    c.iva16 = ColPorNombre(ws, 1, "IVA 16%")
    c.forma = ColPorNombre(ws, 1, "FormaDePago")
    c.comb = ColPorNombre(ws, 1, "Combustible")
    c.rfc = ColPorNombre(ws, 1, "RFC Emisor", "RFC")
    c.ret = ColPorNombre(ws, 1, "Retenido IVA")
    c.total = ColPorNombre(ws, 1, "Total", "Importe Neto")

    c.subtotal = ColOpcional(ws, 1, "SubTotal")
    c.descuento = ColOpcional(ws, 1, "Descuento")
    c.iva8 = ColOpcional(ws, 1, "IVA 8%")
    c.totalTrasladados = ColOpcional(ws, 1, "Total Trasladados")
    c.nivelEducativo = ColOpcional(ws, 1, "Nivel Educativo")
    c.conceptos = ColOpcional(ws, 1, "Conceptos")
    c.lugar = ColOpcional(ws, 1, "LugarDeExpedicion")
    c.uuid = ColOpcional(ws, 1, "UUID")
End Sub

Private Sub LeerColumnasRep(ws As Worksheet, c As ColumnasRep)
    c.fechaPago = ColPorNombre(ws, 1, "FechaPago")
    c.estado = ColPorNombre(ws, 1, "Estado")
    c.bancP = ColPorNombre(ws, 1, "BancarizadoP", "Bancarizado")
    c.base16 = ColPorNombre(ws, 1, "IVA 16 Base", "TotalTrasladosBaseIVA16")
    c.iva16 = ColPorNombre(ws, 1, "IVA 16 Importe", "TotalTrasladosImpuestoIVA16")
    c.rfc = ColPorNombre(ws, 1, "RFC Emisor CFDI", "RFC Emisor")

    c.relUuid = ColOpcional(ws, 1, "RelUUID")
    c.esDuplicado = ColOpcional(ws, 1, "EsDuplicado")
    c.monedaP = ColOpcional(ws, 1, "MonedaP")
    c.equivalencia = ColOpcional(ws, 1, "EquivalenciaDR")
    c.impPagado = ColOpcional(ws, 1, "ImpPagado")
    c.docRelacionado = ColOpcional(ws, 1, "IdDocumento Relacionado")
    c.metodoDR = ColOpcional(ws, 1, "MetodoDePagoDR")
    c.base8 = ColOpcional(ws, 1, "IVA 8 Base")
    c.iva8 = ColOpcional(ws, 1, "IVA 8 Importe")
    c.baseExento = ColOpcional(ws, 1, "IVA Excento Base")
    c.baseCero = ColOpcional(ws, 1, "IVA Cero Base")
    c.lugar = ColOpcional(ws, 1, "Lugar de Expedicion CFDI")
    Set c.retenciones = ColumnasRetencionIva(ws)
End Sub


' ============================================================================
' BOLSAS DE IMPORTES (arreglos 0..6 con los IDX_*)
' ============================================================================
Private Sub LimpiarImportes(imp() As Double)
    Dim i As Long
    For i = 0 To ULT_IDX_IMPORTE
        imp(i) = 0
    Next i
End Sub

' destino = destino + origen (solo los 7 importes)
Private Sub SumarImportes(destino() As Double, origen() As Double)
    Dim i As Long
    For i = 0 To ULT_IDX_IMPORTE
        destino(i) = destino(i) + origen(i)
    Next i
End Sub

' Suma los importes de un documento al total de su bucket y cuenta el documento.
Private Sub AgregarABucket(tot() As Double, imp() As Double)
    SumarImportes tot, imp
    tot(IDX_N) = tot(IDX_N) + 1
End Sub

' Suma los importes de un documento a lo acumulado de su proveedor (RFC).
Private Sub AcumularAgg(agg As Object, ByVal rfc As String, imp() As Double)
    Dim arr As Variant, i As Long
    If agg.Exists(rfc) Then
        arr = agg(rfc)
        For i = 0 To ULT_IDX_IMPORTE
            arr(i) = arr(i) + imp(i)
        Next i
        agg(rfc) = arr
    Else
        Dim nuevo(0 To 6) As Double
        For i = 0 To ULT_IDX_IMPORTE
            nuevo(i) = imp(i)
        Next i
        agg(rfc) = nuevo
    End If
End Sub

' Guarda un aviso una sola vez (el contador lo lee al final).
Private Sub Avisar(avisos As Object, ByVal texto As String)
    If Not avisos.Exists(texto) Then avisos(texto) = True
End Sub

' Aviso si el CP de expedicion de un documento con IVA 8% parece frontera SUR.
' Todo el 8% se reporta como frontera NORTE (decision del despacho); si el CP
' empieza con 24, 29, 30, 77 u 86 se avisa para revisar.
Private Sub AvisarSiFronteraSur(avisos As Object, ByVal rfc As String, ByVal cp As Variant)
    Dim prefijo As String
    prefijo = Left(Trim(CStr(cp)), 2)
    Select Case prefijo
        Case "24", "29", "30", "77", "86"
            Avisar avisos, rfc & " trae IVA 8% y expide desde CP " & Trim(CStr(cp)) & _
                " (frontera SUR?). El TXT lo reporta como frontera NORTE: revisar."
    End Select
End Sub


' ============================================================================
' Bucket A + B: RecibidasXML
' ============================================================================

' Universo del mes: del periodo, Factura, no cancelada, PUE, uso deducible, Total > 0.
Private Function FacturaEsDeducibleDelMes(ws As Worksheet, ByVal r As Long, c As ColumnasFactura, _
        ByVal mes As Long, ByVal anio As Long, usosDed As Object) As Boolean
    FacturaEsDeducibleDelMes = False
    If TextoCelda(ws, r, c.rfc) = "" Then Exit Function
    Dim fecha As Variant
    fecha = ws.Cells(r, c.fecha).Value
    If MesDe(fecha) <> mes Or AnioDe(fecha) <> anio Then Exit Function
    If TextoCelda(ws, r, c.tipo) <> "Factura" Then Exit Function
    If TextoCelda(ws, r, c.estado) = "Cancelado" Then Exit Function
    If Left(TextoCelda(ws, r, c.metodo), 3) <> "PUE" Then Exit Function
    If Not usosDed.Exists(Cod3(ws.Cells(r, c.uso).Value)) Then Exit Function
    If NumCelda(ws, r, c.total) <= 0 Then Exit Function    ' el SAT excluye facturas de $0
    FacturaEsDeducibleDelMes = True
End Function

' "A" bancarizada, "B" efectivo deducible, "" si no se acredita.
Private Function BucketDeFactura(ws As Worksheet, ByVal r As Long, c As ColumnasFactura, _
        ByVal limite As Double) As String
    BucketDeFactura = ""
    If EsSi(ws.Cells(r, c.banc).Value) Then
        BucketDeFactura = "A"
    ElseIf Cod2(ws.Cells(r, c.forma).Value) = "01" _
            And NumCelda(ws, r, c.total) <= limite _
            And Not EsSi(ws.Cells(r, c.comb).Value) Then
        BucketDeFactura = "B"
    End If
End Function

' Una factura con IVA 16% e IVA 8% a la vez rompe la regla SubTotal - Descuento.
Private Function FacturaEsMixta16y8(ws As Worksheet, ByVal r As Long, c As ColumnasFactura) As Boolean
    FacturaEsMixta16y8 = (NumCelda(ws, r, c.iva16) > 0 And NumCelda(ws, r, c.iva8) > 0)
End Function

' Heuristica de acto exento (misma que Resumen!AT): sin IVA 16 ni 8, sin
' Total Trasladados, y es colegiatura (Nivel Educativo lleno) o el concepto
' dice "alumno". Una factura sin impuestos NO es exenta por si sola.
Private Function FacturaEsExenta(ws As Worksheet, ByVal r As Long, c As ColumnasFactura) As Boolean
    FacturaEsExenta = False
    If NumCelda(ws, r, c.iva16) <> 0 Or NumCelda(ws, r, c.iva8) <> 0 Then Exit Function
    If NumCelda(ws, r, c.totalTrasladados) <> 0 Then Exit Function
    If TextoCelda(ws, r, c.nivelEducativo) <> "" Then FacturaEsExenta = True: Exit Function
    If InStr(1, TextoCelda(ws, r, c.conceptos), "alumno", vbTextCompare) > 0 Then FacturaEsExenta = True
End Function

' Llena la bolsa de importes de UNA factura ya clasificada en su bucket.
Private Sub ImportesDeFactura(ws As Worksheet, ByVal r As Long, c As ColumnasFactura, _
        ByVal bucket As String, ByVal rfc As String, imp() As Double, avisos As Object)
    Dim iva16 As Double, iva8 As Double, neto As Double
    LimpiarImportes imp

    ' --- 16%: base = IVA / 0.16 (absorbe IEPS y conceptos exentos/0% de facturas mixtas)
    iva16 = NumCelda(ws, r, c.iva16)
    imp(IDX_IVA16) = iva16
    If iva16 <> 0 Then imp(IDX_BASE16) = iva16 / 0.16

    ' --- 8% region fronteriza norte (campos 8 y 18)
    iva8 = NumCelda(ws, r, c.iva8)
    neto = NumCelda(ws, r, c.subtotal) - NumCelda(ws, r, c.descuento)
    If iva8 <> 0 Then
        imp(IDX_IVA8) = iva8
        If FacturaEsMixta16y8(ws, r, c) Then
            imp(IDX_BASE8) = iva8 / 0.08
            Avisar avisos, rfc & " tiene una factura con IVA 16% e IVA 8% a la vez: la base 8% " & _
                "se tomo como IVA / 0.08; verificar contra el XML."
        ElseIf bucket = "A" Then
            imp(IDX_BASE8) = neto          ' bancarizadas: SubTotal - Descuento
        Else
            imp(IDX_BASE8) = iva8 / 0.08   ' efectivo: el neto podria traer IEPS o exentos
        End If
        AvisarSiFronteraSur avisos, rfc, ValorCelda(ws, r, c.lugar)
    End If

    ' --- exento (campo 50): solo bucket A
    If bucket = "A" Then
        If FacturaEsExenta(ws, r, c) Then imp(IDX_EXENTO) = neto
    End If

    ' --- la tasa 0% de facturas queda en 0 a proposito (se captura a mano)

    ' --- IVA retenido (campo 48)
    imp(IDX_RET) = NumCelda(ws, r, c.ret)
End Sub

Private Sub ProcesarRecibidas(ws As Worksheet, c As ColumnasFactura, ByVal mes As Long, ByVal anio As Long, _
        ByVal limite As Double, usosDed As Object, agg As Object, _
        totA() As Double, totB() As Double, avisos As Object)

    Dim ultFila As Long
    ultFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultFila < 2 Then Exit Sub

    Dim r As Long, rfc As String, bucket As String
    Dim imp(0 To 6) As Double

    For r = 2 To ultFila
        If FacturaEsDeducibleDelMes(ws, r, c, mes, anio, usosDed) Then
            bucket = BucketDeFactura(ws, r, c, limite)
            If bucket <> "" Then
                rfc = TextoCelda(ws, r, c.rfc)
                ImportesDeFactura ws, r, c, bucket, rfc, imp, avisos
                AcumularAgg agg, rfc, imp
                If bucket = "A" Then AgregarABucket totA, imp Else AgregarABucket totB, imp
            End If
        End If
    Next r
End Sub

' {UUID en mayusculas: "Metodo de Pago"} de RecibidasXML.
Private Function ConstruirMetodoPorUuid(ws As Worksheet, c As ColumnasFactura) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set ConstruirMetodoPorUuid = d
    If c.uuid = 0 Then Exit Function

    Dim ultFila As Long, r As Long, uuid As String
    ultFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To ultFila
        uuid = UCase(TextoCelda(ws, r, c.uuid))
        If uuid <> "" Then d(uuid) = TextoCelda(ws, r, c.metodo)
    Next r
End Function


' ============================================================================
' Bucket C: PagosRecibidasXML (REP)
' ============================================================================

' Divisor para pasar a MXN los importes del documento relacionado.
' importe en MXN = importe del documento / EquivalenciaDR (vacia o 0 -> 1).
Private Function EquivalenciaDelDocumento(ws As Worksheet, ByVal r As Long, c As ColumnasRep) As Double
    Dim eq As Double
    eq = NumCelda(ws, r, c.equivalencia)
    If eq > 0 Then EquivalenciaDelDocumento = eq Else EquivalenciaDelDocumento = 1
End Function

Private Function EnPesos(ws As Worksheet, ByVal r As Long, ByVal col As Long, ByVal eq As Double) As Double
    EnPesos = NumCelda(ws, r, col) / eq
End Function

' True si el REP entra al universo. Si se excluye por algo que el contador debe
' investigar (RelUUID distinto de OK, pago en otra moneda) llena `motivo`.
' Las exclusiones normales (otro mes, cancelado, duplicado) no se reportan.
Private Function RepSeIncluye(ws As Worksheet, ByVal r As Long, c As ColumnasRep, _
        ByVal mes As Long, ByVal anio As Long, ByRef motivo As String) As Boolean
    RepSeIncluye = False
    motivo = ""
    If TextoCelda(ws, r, c.rfc) = "" Then Exit Function
    Dim fPago As Variant
    fPago = ws.Cells(r, c.fechaPago).Value
    If MesDe(fPago) <> mes Or AnioDe(fPago) <> anio Then Exit Function
    If TextoCelda(ws, r, c.estado) = "Cancelado" Then Exit Function
    If Not EsSi(ws.Cells(r, c.bancP).Value) Then Exit Function
    If c.esDuplicado > 0 Then
        If NumCelda(ws, r, c.esDuplicado) = 1 Then Exit Function
    End If
    If c.relUuid > 0 Then
        If TextoCelda(ws, r, c.relUuid) <> "OK" Then
            motivo = "RelUUID=" & TextoCelda(ws, r, c.relUuid)
            Exit Function
        End If
    End If
    ' Si el pago mismo no es MXN haria falta el TipoCambioP: no se adivina.
    Dim monedaPago As String
    monedaPago = UCase(TextoCelda(ws, r, c.monedaP))
    If monedaPago <> "" And monedaPago <> "MXN" Then
        motivo = "MonedaP=" & monedaPago & " (revisar a mano)"
        Exit Function
    End If
    RepSeIncluye = True
End Function

' Llena la bolsa de importes de un REP, ya convertidos a MXN.
Private Sub ImportesDeRep(ws As Worksheet, ByVal r As Long, c As ColumnasRep, imp() As Double)
    Dim eq As Double, k As Variant
    eq = EquivalenciaDelDocumento(ws, r, c)
    LimpiarImportes imp
    imp(IDX_BASE16) = EnPesos(ws, r, c.base16, eq)
    imp(IDX_IVA16) = EnPesos(ws, r, c.iva16, eq)
    imp(IDX_BASE8) = EnPesos(ws, r, c.base8, eq)
    imp(IDX_IVA8) = EnPesos(ws, r, c.iva8, eq)
    imp(IDX_EXENTO) = EnPesos(ws, r, c.baseExento, eq)
    imp(IDX_TASA0) = EnPesos(ws, r, c.baseCero, eq)
    For Each k In c.retenciones
        imp(IDX_RET) = imp(IDX_RET) + EnPesos(ws, r, CLng(k), eq)
    Next k
End Sub

' True si la factura que paga el REP es PUE (ya se conto al emitirse; el SAT la
' cuenta doble y la macro tambien, pero avisa). Pagos 2.0 no trae MetodoDePagoDR,
' asi que el metodo se busca en la factura relacionada por UUID.
Private Function RepPagaFacturaPue(ws As Worksheet, ByVal r As Long, c As ColumnasRep, _
        metodoPorUuid As Object) As Boolean
    Dim uuid As String, metodo As String
    uuid = UCase(TextoCelda(ws, r, c.docRelacionado))
    metodo = ""
    If uuid <> "" Then
        If metodoPorUuid.Exists(uuid) Then metodo = CStr(metodoPorUuid(uuid))
    End If
    If metodo = "" Then metodo = TextoCelda(ws, r, c.metodoDR)
    RepPagaFacturaPue = (UCase(Left(Trim(metodo), 3)) = "PUE")
End Function

Private Sub AvisosDeRep(ws As Worksheet, ByVal r As Long, c As ColumnasRep, ByVal rfc As String, _
        imp() As Double, metodoPorUuid As Object, avisos As Object)
    If RepPagaFacturaPue(ws, r, c, metodoPorUuid) Then
        Avisar avisos, rfc & ": REP que paga una factura PUE (" & _
            Format(EnPesos(ws, r, c.impPagado, EquivalenciaDelDocumento(ws, r, c)), "#,##0.00") & _
            " pagado). Se cuenta dos veces, igual que el SAT; pedir al proveedor que lo corrija."
    End If
    If imp(IDX_RET) <> 0 Then
        Avisar avisos, rfc & ": REP con IVA retenido " & Format(imp(IDX_RET), "#,##0.00") & _
            " (sumado al campo 48)."
    End If
    If imp(IDX_IVA8) <> 0 Then AvisarSiFronteraSur avisos, rfc, ValorCelda(ws, r, c.lugar)
End Sub

Private Sub ProcesarPagos(ws As Worksheet, c As ColumnasRep, ByVal mes As Long, ByVal anio As Long, _
        metodoPorUuid As Object, agg As Object, totC() As Double, _
        ByRef descartados As String, avisos As Object)

    Dim ultFila As Long
    ultFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultFila < 2 Then Exit Sub

    Dim r As Long, rfc As String, motivo As String
    Dim imp(0 To 6) As Double

    For r = 2 To ultFila
        rfc = TextoCelda(ws, r, c.rfc)
        If RepSeIncluye(ws, r, c, mes, anio, motivo) Then
            ImportesDeRep ws, r, c, imp
            AvisosDeRep ws, r, c, rfc, imp, metodoPorUuid, avisos
            AcumularAgg agg, rfc, imp
            AgregarABucket totC, imp
        ElseIf motivo <> "" Then
            descartados = descartados & "  " & rfc & "  " & motivo & "  IVA " & _
                Format(NumCelda(ws, r, c.iva16), "#,##0.00") & vbCrLf
        End If
    Next r
End Sub


' ============================================================================
' REDONDEO Art. 20 CFF + invariante SAT
' ============================================================================
Private Function CffRound(ByVal x As Double) As Long
    If x < 0 Then CffRound = -CffRound(-x): Exit Function
    Dim n As Long, cents As Long
    n = Int(x)
    cents = CLng(Round((x - n) * 100, 0))
    If cents >= 51 Then CffRound = n + 1 Else CffRound = n
End Function

' Minimo entero base >= baseInt tal que (base*tasa)\100 >= ivaInt  (tasa = 16 u 8).
' Equivale a base >= ceil(iva*100/tasa). Todo en enteros (Long): con Double,
' 1257*0.16 da 201.12000000000003 y el truncado del SAT se calcularia mal.
' Se separa iva = q*tasa + resto para que iva*100 no desborde un Long:
'   iva*100/tasa = q*100 + resto*100/tasa
Private Function AsegurarInvariante(ByVal baseInt As Long, ByVal ivaInt As Long, _
        ByVal tasaPorCiento As Long) As Long
    If ivaInt <= 0 Then AsegurarInvariante = baseInt: Exit Function
    Dim minimo As Long
    minimo = (ivaInt \ tasaPorCiento) * 100 + _
             (((ivaInt Mod tasaPorCiento) * 100) + tasaPorCiento - 1) \ tasaPorCiento
    If baseInt > minimo Then AsegurarInvariante = baseInt Else AsegurarInvariante = minimo
End Function

' Redondea los 7 importes de un proveedor y aplica la invariante a 16% y a 8%.
' Suma a ajuste16 / ajuste8 los pesos que se subieron a las bases.
Private Sub RedondearProveedor(arr As Variant, ent() As Long, ByRef ajuste16 As Long, ByRef ajuste8 As Long)
    Dim i As Long, ajustada As Long
    For i = 0 To ULT_IDX_IMPORTE
        ent(i) = CffRound(CDbl(arr(i)))
    Next i

    ajustada = AsegurarInvariante(ent(IDX_BASE16), ent(IDX_IVA16), 16)
    ajuste16 = ajuste16 + (ajustada - ent(IDX_BASE16))
    ent(IDX_BASE16) = ajustada

    ajustada = AsegurarInvariante(ent(IDX_BASE8), ent(IDX_IVA8), 8)
    ajuste8 = ajuste8 + (ajustada - ent(IDX_BASE8))
    ent(IDX_BASE8) = ajustada
End Sub

Private Function TodosLosEnterosSonCero(ent() As Long) As Boolean
    Dim i As Long
    TodosLosEnterosSonCero = True
    For i = 0 To ULT_IDX_IMPORTE
        If ent(i) <> 0 Then TodosLosEnterosSonCero = False: Exit Function
    Next i
End Function


' ============================================================================
' construir los 54 campos por RFC
' ============================================================================

' Llena los 54 campos de un proveedor nacional. Campos sin dato = "".
Private Sub ArmarCampos(campos() As String, ByVal rfc As String, ByVal op As String, ent() As Long)
    Dim j As Long
    For j = 1 To 54: campos(j) = "": Next j
    campos(1) = "04"
    campos(2) = op
    campos(3) = rfc
    If ent(IDX_BASE16) <> 0 Or ent(IDX_IVA16) <> 0 Then
        campos(12) = CStr(ent(IDX_BASE16))
        campos(22) = CStr(ent(IDX_IVA16))
    End If
    If ent(IDX_BASE8) <> 0 Or ent(IDX_IVA8) <> 0 Then
        campos(8) = CStr(ent(IDX_BASE8))
        campos(18) = CStr(ent(IDX_IVA8))
    End If
    If ent(IDX_RET) <> 0 Then campos(48) = CStr(ent(IDX_RET))
    If ent(IDX_EXENTO) <> 0 Then campos(50) = CStr(ent(IDX_EXENTO))
    If ent(IDX_TASA0) <> 0 Then campos(51) = CStr(ent(IDX_TASA0))
End Sub

Private Sub OrdenarClaves(claves As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(claves) To UBound(claves) - 1
        For j = i + 1 To UBound(claves)
            If claves(i) > claves(j) Then
                tmp = claves(i): claves(i) = claves(j): claves(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Sub ConstruirRenglones(agg As Object, opMap As Object, ByRef lineas() As String, ByRef nLineas As Long, _
        totTxt() As Long, ByRef ajuste16 As Long, ByRef ajuste8 As Long)

    Dim claves As Variant, rfc As Variant
    Dim n As Long, i As Long, k As Long
    n = agg.Count
    ReDim lineas(1 To IIf(n = 0, 1, n))
    nLineas = 0
    ajuste16 = 0: ajuste8 = 0
    For i = 0 To ULT_IDX_IMPORTE: totTxt(i) = 0: Next i

    If n = 0 Then Exit Sub
    claves = agg.Keys
    OrdenarClaves claves      ' orden alfabetico (como sorted() en Python)

    Dim ent(0 To 6) As Long
    Dim op As String
    Dim campos(1 To 54) As String

    For i = LBound(claves) To UBound(claves)
        rfc = claves(i)
        RedondearProveedor agg(rfc), ent, ajuste16, ajuste8
        If Not TodosLosEnterosSonCero(ent) Then
            If opMap.Exists(CStr(rfc)) Then op = opMap(CStr(rfc)) Else op = DEFAULT_OP
            ArmarCampos campos, CStr(rfc), op, ent

            nLineas = nLineas + 1
            lineas(nLineas) = Join(campos, "|")

            For k = 0 To ULT_IDX_IMPORTE
                totTxt(k) = totTxt(k) + ent(k)
            Next k
        End If
    Next i
End Sub

' Una linea por cada total del TXT que no sea cero (16% siempre se muestra).
Private Function TextoTotalesTxt(totTxt() As Long) As String
    Dim s As String
    s = "TXT TOTAL (CFF redon.)   base 16% " & Format(totTxt(IDX_BASE16), "#,##0") & _
        "   IVA 16% " & Format(totTxt(IDX_IVA16), "#,##0") & vbCrLf
    If totTxt(IDX_BASE8) <> 0 Or totTxt(IDX_IVA8) <> 0 Then
        s = s & "   frontera norte 8%: base " & Format(totTxt(IDX_BASE8), "#,##0") & _
            "   IVA " & Format(totTxt(IDX_IVA8), "#,##0") & vbCrLf
    End If
    If totTxt(IDX_EXENTO) <> 0 Then s = s & "   exentos (campo 50): " & Format(totTxt(IDX_EXENTO), "#,##0") & vbCrLf
    If totTxt(IDX_TASA0) <> 0 Then s = s & "   tasa 0% de REP (campo 51): " & Format(totTxt(IDX_TASA0), "#,##0") & vbCrLf
    If totTxt(IDX_RET) <> 0 Then s = s & "   IVA retenido (campo 48): " & Format(totTxt(IDX_RET), "#,##0") & vbCrLf
    TextoTotalesTxt = s
End Function

Private Function TextoDeAvisos(avisos As Object) As String
    Dim s As String, k As Variant
    For Each k In avisos.Keys
        s = s & "AVISO: " & CStr(k) & vbCrLf
    Next k
    If s <> "" Then s = vbCrLf & s
    TextoDeAvisos = s
End Function

' MsgBox corta el texto cerca de 1000 caracteres; se muestra en partes por renglones.
Private Sub MostrarMensajeLargo(ByVal msg As String, ByVal titulo As String)
    Dim partes() As String, i As Long, bloque As String
    partes = Split(msg, vbCrLf)
    For i = LBound(partes) To UBound(partes)
        If Len(bloque) + Len(partes(i)) > 800 And bloque <> "" Then
            MsgBox bloque, vbInformation, titulo
            bloque = ""
        End If
        bloque = bloque & partes(i) & vbCrLf
    Next i
    If bloque <> "" Then MsgBox bloque, vbInformation, titulo
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


' ============================================================================
' CONCILIACION CONTRA RESUMEN (por ENCABEZADO, nunca por letra de columna)
' ============================================================================
' Resumen tiene dos filas de encabezado:
'   fila 1: titulo del bloque, en la primera columna del bloque (celdas
'           combinadas), p. ej. "Compras PUE bancarizadas (G01+G03, vigentes)"
'   fila 2: subtitulo de cada columna, p. ej. "Facturas", "IVA 16%"
' "IVA 16%" se repite en varios bloques, asi que una columna obligatoria se
' identifica por el par (bloque, subtitulo). Las columnas nuevas (8%, exentos,
' tasa 0%) tienen subtitulo unico y son OPCIONALES (libros viejos no las traen).
' Resumen!Y ya NO es base de REP (ahora es "Pagado s/IVA"): no se lee.

Private Function BuscarHoja(ByVal nombre As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nombre Then Set BuscarHoja = ws: Exit Function
    Next ws
    Set BuscarHoja = Nothing
End Function

' Fila de datos cuya columna A es el numero de mes (0 si no hay).
Private Function FilaDelMesEnResumen(ws As Worksheet, ByVal mes As Long) As Long
    Dim r As Long, v As Variant
    FilaDelMesEnResumen = 0
    For r = 3 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        v = ws.Cells(r, 1).Value
        If VarType(v) = vbDouble Then
            If CLng(v) = mes Then FilaDelMesEnResumen = r: Exit Function
        End If
    Next r
End Function

Private Function UltimaColumnaResumen(ws As Worksheet) As Long
    Dim c1 As Long, c2 As Long
    c1 = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    c2 = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    If c1 > c2 Then UltimaColumnaResumen = c1 Else UltimaColumnaResumen = c2
End Function

' Extiende cada titulo de bloque (fila 1) hacia la derecha hasta el siguiente titulo.
Private Sub LlenarTitulosDeBloque(ws As Worksheet, ByVal ultCol As Long, titulos() As String)
    Dim i As Long, vigente As String, t As String
    vigente = ""
    For i = 1 To ultCol
        t = Trim(CStr(ws.Cells(1, i).Value))
        If t <> "" Then vigente = t
        titulos(i) = vigente
    Next i
End Sub

' Columna cuyo subtitulo (fila 2) es EXACTO y cuyo bloque contiene el texto. 0 si no esta.
Private Function ColResumenObligatoria(ws As Worksheet, ByVal ultCol As Long, titulos() As String, _
        ByVal textoBloque As String, ByVal subtitulo As String) As Long
    Dim i As Long
    For i = 1 To ultCol
        If Trim(CStr(ws.Cells(2, i).Value)) = subtitulo Then
            If InStr(1, titulos(i), textoBloque, vbTextCompare) > 0 Then
                ColResumenObligatoria = i
                Exit Function
            End If
        End If
    Next i
    ColResumenObligatoria = 0
End Function

' Primera columna cuyo subtitulo (fila 2) es EXACTO. 0 si no esta (columna opcional).
Private Function ColResumenOpcional(ws As Worksheet, ByVal ultCol As Long, ByVal subtitulo As String) As Long
    Dim i As Long
    For i = 1 To ultCol
        If Trim(CStr(ws.Cells(2, i).Value)) = subtitulo Then ColResumenOpcional = i: Exit Function
    Next i
    ColResumenOpcional = 0
End Function

' Si la columna obligatoria no existe, agrega su descripcion a la lista de faltantes.
Private Sub RegistrarFaltante(ByRef faltantes As String, ByVal col As Long, _
        ByVal textoBloque As String, ByVal subtitulo As String)
    If col = 0 Then
        faltantes = faltantes & "  - '" & subtitulo & "' dentro del bloque '" & textoBloque & "...'" & vbCrLf
    End If
End Sub

' Una linea "etiqueta  script  OK/DIF" comparando contra la celda de Resumen.
Private Function LineaResumen(ByVal etiqueta As String, ByVal nuestro As Double, ws As Worksheet, _
        ByVal fila As Long, ByVal col As Long, ByVal esConteo As Boolean) As String
    Dim esperado As Double, tolerancia As Double, fmt As String, estado As String
    esperado = Num(ws.Cells(fila, col).Value)
    If esConteo Then tolerancia = 0.5: fmt = "#,##0" Else tolerancia = 0.01: fmt = "#,##0.00"
    If Abs(nuestro - esperado) < tolerancia Then
        estado = "OK"
    Else
        estado = "DIF " & Format(nuestro - esperado, "0.00")
    End If
    LineaResumen = Left(etiqueta & String(28, " "), 28) & Format(nuestro, fmt) & "   " & estado & vbCrLf
End Function

' Agrega la linea de una columna opcional solo si Resumen la tiene.
Private Sub AgregarLineaOpcional(ByRef s As String, ws As Worksheet, ByVal ultCol As Long, ByVal fila As Long, _
        ByVal subtitulo As String, ByVal etiqueta As String, ByVal nuestro As Double)
    Dim col As Long
    col = ColResumenOpcional(ws, ultCol, subtitulo)
    If col > 0 Then s = s & LineaResumen(etiqueta, nuestro, ws, fila, col, False)
End Sub

' Texto cuando no se puede conciliar (sin hoja, sin fila del mes o faltan columnas).
Private Function TextoSinResumen(totA() As Double, totB() As Double, totC() As Double) As String
    TextoSinResumen = "Bucket A  IVA " & Format(totA(IDX_IVA16), "#,##0.00") & vbCrLf & _
                      "Bucket B  IVA " & Format(totB(IDX_IVA16), "#,##0.00") & vbCrLf & _
                      "Bucket C  IVA " & Format(totC(IDX_IVA16), "#,##0.00") & vbCrLf
End Function

Private Function ConciliarConResumen(ByVal mes As Long, totA() As Double, totB() As Double, _
        totC() As Double, ByVal limite As Double) As String

    Dim ws As Worksheet, fila As Long, ultCol As Long
    Set ws = BuscarHoja("Resumen")
    If ws Is Nothing Then ConciliarConResumen = TextoSinResumen(totA, totB, totC): Exit Function
    fila = FilaDelMesEnResumen(ws, mes)
    If fila = 0 Then ConciliarConResumen = TextoSinResumen(totA, totB, totC): Exit Function

    ultCol = UltimaColumnaResumen(ws)
    Dim titulos() As String
    ReDim titulos(1 To ultCol)
    LlenarTitulosDeBloque ws, ultCol, titulos

    ' --- columnas obligatorias (bloque, subtitulo exacto)
    Const BLQ_A As String = "Compras PUE bancarizadas"
    Const BLQ_B As String = "Compras en efectivo"
    Const BLQ_C As String = "REP recibidos (PPD"
    Dim cAn As Long, cAiva As Long, cBn As Long, cBiva As Long, cCn As Long, cCiva As Long
    cAn = ColResumenObligatoria(ws, ultCol, titulos, BLQ_A, "Facturas")
    cAiva = ColResumenObligatoria(ws, ultCol, titulos, BLQ_A, "IVA 16%")
    cBn = ColResumenObligatoria(ws, ultCol, titulos, BLQ_B, "Facturas")
    cBiva = ColResumenObligatoria(ws, ultCol, titulos, BLQ_B, "IVA 16%")
    cCn = ColResumenObligatoria(ws, ultCol, titulos, BLQ_C, "REPs")
    cCiva = ColResumenObligatoria(ws, ultCol, titulos, BLQ_C, "IVA 16%")

    Dim faltantes As String
    RegistrarFaltante faltantes, cAn, BLQ_A, "Facturas"
    RegistrarFaltante faltantes, cAiva, BLQ_A, "IVA 16%"
    RegistrarFaltante faltantes, cBn, BLQ_B, "Facturas"
    RegistrarFaltante faltantes, cBiva, BLQ_B, "IVA 16%"
    RegistrarFaltante faltantes, cCn, BLQ_C, "REPs"
    RegistrarFaltante faltantes, cCiva, BLQ_C, "IVA 16%"
    If faltantes <> "" Then
        MsgBox "Resumen: no encontre estas columnas (filas 1 y 2 de Resumen):" & vbCrLf & faltantes & _
               "El TXT ya se guardo, pero NO se concilio contra Resumen. Revisa los encabezados.", _
               vbCritical, "DIOT - Resumen"
        ConciliarConResumen = "Resumen NO conciliado (faltan columnas, ver aviso)." & vbCrLf & _
                              TextoSinResumen(totA, totB, totC)
        Exit Function
    End If

    Dim s As String
    s = LineaResumen("A PUE banc  documentos", totA(IDX_N), ws, fila, cAn, True)
    s = s & LineaResumen("A PUE banc  IVA 16%", totA(IDX_IVA16), ws, fila, cAiva, False)
    s = s & LineaResumen("B efectivo<=" & Format(limite, "0") & " docs", totB(IDX_N), ws, fila, cBn, True)
    s = s & LineaResumen("B efectivo  IVA 16%", totB(IDX_IVA16), ws, fila, cBiva, False)
    s = s & LineaResumen("C REP (PPD) documentos", totC(IDX_N), ws, fila, cCn, True)
    s = s & LineaResumen("C REP (PPD) IVA 16%", totC(IDX_IVA16), ws, fila, cCiva, False)

    ' --- columnas opcionales (bloques nuevos; solo si el libro las trae)
    AgregarLineaOpcional s, ws, ultCol, fila, "Base 8% facturas", "A base 8%", totA(IDX_BASE8)
    AgregarLineaOpcional s, ws, ultCol, fila, "IVA 8% facturas", "A IVA 8%", totA(IDX_IVA8)
    AgregarLineaOpcional s, ws, ultCol, fila, "Base 8% efectivo", "B base 8%", totB(IDX_BASE8)
    AgregarLineaOpcional s, ws, ultCol, fila, "IVA 8% efectivo", "B IVA 8%", totB(IDX_IVA8)
    AgregarLineaOpcional s, ws, ultCol, fila, "Base 8% REP", "C base 8%", totC(IDX_BASE8)
    AgregarLineaOpcional s, ws, ultCol, fila, "IVA 8% REP", "C IVA 8%", totC(IDX_IVA8)
    AgregarLineaOpcional s, ws, ultCol, fila, "Exento facturas", "A exento", totA(IDX_EXENTO)
    AgregarLineaOpcional s, ws, ultCol, fila, "Exento REP", "C exento", totC(IDX_EXENTO)
    AgregarLineaOpcional s, ws, ultCol, fila, "Base 0% REP", "C tasa 0%", totC(IDX_TASA0)

    ConciliarConResumen = s
End Function
