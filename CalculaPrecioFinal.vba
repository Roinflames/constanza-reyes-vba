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
    ws.Cells(fila, ws.Rows(1).Find("Importe de venta total", , xlValues, xlWhole).Column).Value = importeVenta
    ws.Cells(fila, ws.Rows(1).Find("Importe de coste total", , xlValues, xlWhole).Column).Value = importeCoste
    
    ' Precio final
    precioFinal = (importeVenta - (importeVenta * porcDesc)) * (1 + aumento) * factor
    ws.Cells(fila, ws.Rows(1).Find("Precio final", , xlValues, xlWhole).Column).Value = precioFinal
    
    MsgBox "✅ Precio final calculado correctamente para la fila " & fila, vbInformation
End Sub