
Sub FormatoTabla()
    ' Seleccionar la hoja "Ventas"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Ventas")
    ws.Activate

    ' Declarar variables para el rango y la última fila
    Dim lastRow As Long
    Dim dataRange As Range
    Dim headerRange As Range
    
    ' Encontrar la última fila con datos en la columna A (asumiendo que ID_Pedido es la columna original A)
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    ' --- 1. Insertar nuevas columnas ---
    ' Insertar columna para Id_Cliente en la posición A
    ws.Columns("A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Range("A1").Value = "Id_Cliente"
    
    ' Insertar columna para Porc descuento en la posición J
    ws.Columns("J").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Range("J1").Value = "Porc descuento"
    
    ' --- 2. Formato de encabezado ---
    Set headerRange = ws.Range("A1:P1") ' Ajustado al nuevo número de columnas
    With headerRange
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = vbGreen
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' --- 3. Autoajustar columnas y agregar bordes ---
    ws.Columns.AutoFit
    
    ' Actualizar la última fila y el rango de datos
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    Set dataRange = ws.Range("A1:P" & lastRow)
    
    With dataRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' --- 4. Formato de columnas específicas ---
    ' Formato para Porc descuento (Columna J)
    ws.Columns("J").NumberFormat = "0.0"
    
    ' Formato para columnas de moneda (L a P)
    ws.Range("L:P").NumberFormat = "#,##0.00 €"

    ' --- 5. Agregar filas de resumen (Máximo, Mínimo, Promedio) ---
    Dim summaryRowStart As Long
    summaryRowStart = lastRow + 2 ' Dejar una fila en blanco

    ' Rótulos
    ws.Range("J" & summaryRowStart).Value = "Máximo"
    ws.Range("J" & summaryRowStart + 1).Value = "Mínimo"
    ws.Range("J" & summaryRowStart + 2).Value = "Promedio"
    
    ' Formato para los rótulos
    With ws.Range("J" & summaryRowStart & ":J" & summaryRowStart + 2)
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = vbGreen
        .HorizontalAlignment = xlCenter
    End With
    
    ' Fórmulas para la columna Unidades (K)
    Dim unidadesRange As String
    unidadesRange = "K2:K" & lastRow
    ws.Range("K" & summaryRowStart).Formula = "=MAX(" & unidadesRange & ")"
    ws.Range("K" & summaryRowStart + 1).Formula = "=MIN(" & unidadesRange & ")"
    ws.Range("K" & summaryRowStart + 2).Formula = "=AVERAGE(" & unidadesRange & ")"

    ' --- 6. Formato Condicional para la columna Unidades (K) ---
    Dim unidadesDataRange As Range
    Set unidadesDataRange = ws.Range("K2:K" & lastRow)
    
    ' Limpiar formatos condicionales existentes
    unidadesDataRange.FormatConditions.Delete
    
    ' Regla 1: Prioridad "Crítica" y Unidades > 6000 (Debe ir primero por precedencia)
    Dim condition1 As FormatCondition
    Set condition1 = unidadesDataRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=Y($K2>6000; $F2=""Crítica"")")
    With condition1
        .StopIfTrue = True
        With .Interior
            .PatternColorIndex = xlAutomatic
            .Color = vbRed
        End With
        With .Font
            .Color = vbWhite
            .Bold = True
        End With
    End With
    
    ' Regla 2: Unidades > 6000
    Dim condition2 As FormatCondition
    Set condition2 = unidadesDataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=6000")
    With condition2.Interior
        .PatternColorIndex = xlAutomatic
        .Color = vbGreen
    End With
    With condition2.Font
        .Color = vbWhite
        .Italic = True
    End With
    
    ' Regla 3: Unidades >= 2500 (y <= 6000)
    Dim condition3 As FormatCondition
    Set condition3 = unidadesDataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=2500")
    With condition3.Interior
        .PatternColorIndex = xlAutomatic
        .Color = vbBlue
    End With
    With condition3.Font
        .Color = vbWhite
        .Bold = True
    End With

    ' Regla 4: Unidades < 2500
    Dim condition4 As FormatCondition
    Set condition4 = unidadesDataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=2500")
    With condition4.Interior
        .PatternColorIndex = xlAutomatic
        .Color = vbYellow
    End With
    With condition4.Font
        .Color = vbBlack
        .Bold = True
    End With
    
    ' Mensaje de finalización
    MsgBox "La macro 'FormatoTabla' se ha ejecutado correctamente.", vbInformation

End Sub
