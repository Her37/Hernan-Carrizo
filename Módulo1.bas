Attribute VB_Name = "Módulo1"
Sub AplicarFiltroDesdeHoja1()
    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    Dim ultimaFila As Long, i As Long
    Dim criterios As String
    
    ' Definir hojas
    Set wsOrigen = Sheets("Hoja1")
    Set wsDestino = Sheets("Base Trabajo")
    
    ' Encontrar la última fila con datos en la columna A de Hoja1
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    
    ' Construir la cadena de criterios para el filtro
    For i = 2 To ultimaFila
        If criterios = "" Then
            criterios = "=" & wsOrigen.Cells(i, 1).Value
        Else
            criterios = criterios & ",=" & wsOrigen.Cells(i, 1).Value
        End If
    Next i
    
    ' Aplicar filtro en la columna A de Base Trabajo
    wsDestino.Range("$A$1:$AE$922").AutoFilter Field:=1, Criteria1:=Split(criterios, ","), Operator:=xlFilterValues
    
    ' Liberar memoria
    Set wsOrigen = Nothing
    Set wsDestino = Nothing
End Sub

