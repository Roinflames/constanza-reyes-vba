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