'============================================================
'     TAREA 2 – INF130 | Programación y Tratamiento de Datos
'     AUTOR: 
'     DESCRIPCIÓN:
'     Este módulo implementa las macros:
'        1) FormatoTabla()
'        2) calculardatos()
'        3) CalculaPrecioFinal()
'============================================================

Option Explicit

'------------------------------------------------------------
' 1) FORMATO DE LA TABLA – Pregunta 1
'------------------------------------------------------------
Sub FormatoTabla()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim headerRange As Range
    Dim colZona As Long, colFechaEnvio As Long, colUnidades As Long
    Dim colPorcDesc As Long
    Dim summaryRowStart As Long
    Dim unidadesDataRange As Range
    Dim formula As String
    Dim sep As String
    
    Set ws = ThisWorkbook.Sheets("Ventas")
    ws.Activate
    sep = Application.International(xlListSeparator)
    
    '--- Ubicar columnas ---
    colZona = ws.Rows(1).Find("Zona", , xlValues, xlWhole).Column
    colFechaEnvio = ws.Rows(1).Find("Fecha envío", , xlValues, xlWhole).Column
    colUnidades = ws.Rows(1).Find("Unidades", , xlValues, xlWhole).Column
    
    '--- Insertar nuevas columnas ---
    ws.Columns(colZona).Insert Shift:=xlToRight
    ws.Cells(1, colZona).Value = "Id_Cliente"
    
    ws.Columns(colFechaEnvio + 1).Insert Shift:=xlToRight
    ws.Cells(1, colFechaEnvio + 1).Value = "Porc descuento"
    colPorcDesc = colFechaEnvio + 1
    
    '--- Formato de encabezado ---
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
    With headerRange
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(0, 176, 80)
        .HorizontalAlignment = xlCenter
    End With
    
    '--- Bordes y ajuste ---
    lastRow = ws.Cells(ws.Rows.Count, colUnidades).End(xlUp).Row
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, headerRange.Columns.Count))
    With dataRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ws.Columns.AutoFit
    
    '--- Formatos numéricos ---
    ws.Columns(colPorcDesc).NumberFormat = "0.0"
    Dim colInicioMoneda As Long
    colInicioMoneda = ws.Rows(1).Find("Precio unitario", , xlValues, xlWhole).Column
    ws.Range(ws.Cells(2, colInicioMoneda), ws.Cells(lastRow, colInicioMoneda + 4)).NumberFormat = "#,##0.00"
    
    '--- Filas resumen ---
    summaryRowStart = lastRow + 2
    ws.Range("J" & summaryRowStart).Value = "Máximo"
    ws.Range("J" & summaryRowStart + 1).Value = "Mínimo"
    ws.Range("J" & summaryRowStart + 2).Value = "Promedio"
    With ws.Range("J" & summaryRowStart & ":J" & summaryRowStart + 2)
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(0, 176, 80)
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Cells(summaryRowStart, colUnidades).Formula = "=MAX(K2:K" & lastRow & ")"
    ws.Cells(summaryRowStart + 1, colUnidades).Formula = "=MIN(K2:K" & lastRow & ")"
    ws.Cells(summaryRowStart + 2, colUnidades).Formula = "=PROMEDIO(K2:K" & lastRow & ")"
    
    '--- Formato condicional ---
    Set unidadesDataRange = ws.Range("K2:K" & summaryRowStart - 2)
    unidadesDataRange.FormatConditions.Delete
    
    ' 1) Crítica + >6000
    formula = "=Y($K2>6000" & sep & "$F2=""Crítica"")"
    Dim c1 As FormatCondition
    Set c1 = unidadesDataRange.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
    With c1
        .StopIfTrue = True
        .Interior.Color = vbRed
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    
    ' 2) >6000
    Dim c2 As FormatCondition
    Set c2 = unidadesDataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="6000")
    With c2
        .Interior.Color = vbGreen
        .Font.Color = vbWhite
        .Font.Italic = True
    End With
    
    ' 3) 2500–6000
    Dim c3 As FormatCondition
    Set c3 = unidadesDataRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=Y($K2>=2500" & sep & "$K2<=6000)")
    With c3
        .Interior.Color = vbBlue
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    
    ' 4) <2500
    Dim c4 As FormatCondition
    Set c4 = unidadesDataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="2500")
    With c4
        .Interior.Color = vbYellow
        .Font.Color = vbBlack
        .Font.Bold = True
    End With
    
    MsgBox "✅ La macro 'FormatoTabla' se ejecutó correctamente.", vbInformation
End Sub

'------------------------------------------------------------
' 2) CALCULAR DATOS – Pregunta 2
'------------------------------------------------------------
Sub calculardatos()
    Dim ws As Worksheet
    Dim fila As Long
    Dim fechaPedido As Date, fechaEnvio As Date
    Dim diferencia As Long, descuento As Double
    Dim pais As String, zona As String, idCliente As String
    Dim codigoZona As String
    
    Set ws = ThisWorkbook.Sheets("Ventas")
    
    fila = InputBox("Ingrese el número de fila a analizar:", "Calcular Datos")
    If fila < 2 Then Exit Sub
    
    fechaPedido = ws.Cells(fila, ws.Rows(1).Find("Fecha pedido", , xlValues, xlWhole).Column).Value
    fechaEnvio = ws.Cells(fila, ws.Rows(1).Find("Fecha envío", , xlValues, xlWhole).Column).Value
    diferencia = DateDiff("d", fechaPedido, fechaEnvio)
    
    ' --- a) Porcentaje de descuento ---
    Select Case diferencia
        Case Is < 10: descuento = 0
        Case 10 To 24: descuento = 0.2
        Case 25 To 39: descuento = 0.3
        Case Is >= 40: descuento = 0.4
    End Select
    ws.Cells(fila, ws.Rows(1).Find("Porc descuento", , xlValues, xlWhole).Column).Value = descuento
    
    ' --- b) Id_Cliente ---
    pais = ws.Cells(fila, ws.Rows(1).Find("País", , xlValues, xlWhole).Column).Value
    zona = ws.Cells(fila, ws.Rows(1).Find("Zona", , xlValues, xlWhole).Column).Value
    
    Select Case zona
        Case "África": codigoZona = "AFR"
        Case "Asia": codigoZona = "ASI"
        Case "Australia y Oceanía": codigoZona = "AUS"
        Case "Centroamérica y Caribe": codigoZona = "CEN"
        Case "Europa": codigoZona = "EUR"
        Case "Norteamérica": codigoZona = "NOR"
        Case Else: codigoZona = "OTR"
    End Select
    
    idCliente = UCase(Left(pais, 5)) & "-" & codigoZona
    ws.Cells(fila, ws.Rows(1).Find("Id_Cliente", , xlValues, xlWhole).Column).Value = idCliente
    
    MsgBox "✅ Datos calculados correctamente para la fila " & fila, vbInformation
End Sub

'------------------------------------------------------------
' 3) CALCULAR PRECIO FINAL – Pregunta 3
'------------------------------------------------------------
Sub CalculaPrecioFinal()
    Dim ws As Worksheet
    Dim fila As Long
    Dim unidades As Double, precioU As Double, costeU As Double
    Dim porcDesc As Double, prioridad As String, canal As String
    Dim aumento As Double, factor As Double
    Dim importeVenta As Double, importeCoste As Double, precioFinal As Double
    
    Set ws = ThisWorkbook.Sheets("Ventas")
    fila = InputBox("Ingrese el número de fila a calcular:", "Calcular Precio Final")
    If fila < 2 Then Exit Sub
    
    ' Obtener valores
    unidades = ws.Cells(fila, ws.Rows(1).Find("Unidades", , xlValues, xlWhole).Column).Value
    precioU = ws.Cells(fila, ws.Rows(1).Find("Precio unitario", , xlValues, xlWhole).Column).Value
    costeU = ws.Cells(fila, ws.Rows(1).Find("Coste unitario", , xlValues, xlWhole).Column).Value
    porcDesc = ws.Cells(fila, ws.Rows(1).Find("Porc descuento", , xlValues, xlWhole).Column).Value
    prioridad = ws.Cells(fila, ws.Rows(1).Find("Prioridad", , xlValues, xlWhole).Column).Value
    canal = ws.Cells(fila, ws.Rows(1).Find("Canal de venta", , xlValues, xlWhole).Column).Value
    
    ' --- Aumento por prioridad ---
    Select Case prioridad
        Case "Baja": aumento = 0
        Case "Media": aumento = 0.1
        Case "Alta": aumento = 0.2
        Case "Crítica": aumento = 0.25
        Case Else: aumento = 0
    End Select
    
    ' --- Factor por canal ---
    If canal = "Online" Then
        factor = 0.7
    ElseIf canal = "Offline" Then
        factor = 0.95
    Else
        factor = 1
    End If
    
    ' --- Cálculos ---
    importeVenta = unidades * precioU
    importeCoste = unidades * costeU

    Dim colImporteVenta As Range
    Set colImporteVenta = ws.Rows(1).Find("Importe venta total", , xlValues, xlWhole)

    If Not colImporteVenta Is Nothing Then
        ws.Cells(fila, colImporteVenta.Column).Value = importeVenta
    Else
        MsgBox "❌ No se encontró la columna 'Importe venta total'.", vbCritical
        Exit Sub
    End If

    ws.Cells(fila, ws.Rows(1).Find("Importe coste total", , xlValues, xlWhole).Column).Value = importeCoste
    
    ' Precio final
    precioFinal = (importeVenta - (importeVenta * porcDesc)) * (1 + aumento) * factor
    ws.Cells(fila, ws.Rows(1).Find("Precio final", , xlValues, xlWhole).Column).Value = precioFinal
    
    MsgBox "✅ Precio final calculado correctamente para la fila " & fila, vbInformation
End Sub
