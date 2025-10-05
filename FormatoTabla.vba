Sub FormatoTabla()
    '-------------------------------------------------------
    '   MACRO: FormatoTabla()
    '   AUTOR: 
    '   OBJETIVO: Formatear la hoja "Ventas" según PREGUNTA 1
    '-------------------------------------------------------

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim headerRange As Range
    Dim colZona As Long, colFechaEnvio As Long
    Dim colPorcDesc As Long, colUnidades As Long
    Dim summaryRowStart As Long
    Dim unidadesDataRange As Range
    Dim formula As String
    Dim sep As String
    
    '--- Inicializar hoja ---
    Set ws = ThisWorkbook.Sheets("Ventas")
    ws.Activate
    sep = Application.International(xlListSeparator)
    
    '--- Identificar posiciones dinámicamente ---
    colZona = ws.Rows(1).Find("Zona", , xlValues, xlWhole).Column
    colFechaEnvio = ws.Rows(1).Find("Fecha envío", , xlValues, xlWhole).Column
    colUnidades = ws.Rows(1).Find("Unidades", , xlValues, xlWhole).Column
    
    '--- Insertar columnas nuevas ---
    ws.Columns(colZona).Insert Shift:=xlToRight
    ws.Cells(1, colZona).Value = "Id_Cliente"
    
    ws.Columns(colFechaEnvio + 1).Insert Shift:=xlToRight
    ws.Cells(1, colFechaEnvio + 1).Value = "Porc descuento"
    colPorcDesc = colFechaEnvio + 1
    
    '--- Formato encabezado ---
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
    With headerRange
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(0, 176, 80) ' verde
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
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
    
    '--- Agregar filas resumen ---
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
    
    '--- Formato condicional (excluye filas resumen) ---
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
    
    ' 3) Entre 2500 y 6000
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