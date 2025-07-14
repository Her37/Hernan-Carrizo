Attribute VB_Name = "Módulo2"
Function ObtenerItemsBloqueados(codigoContrato As String) As Collection
    Dim i As Integer
    Dim tooltipTexto As String
    Dim valorPrimeraColumna As String
    Dim celdaLOEKZ As Object
    Dim resultado As New Collection
    Dim posicionScroll As Integer
    Dim scrollMax As Integer
    Dim tabla As Object
    Dim totalFilas As Integer
    
    On Error Resume Next
    Call VolverAVentanaPrincipalSAP

    session.findById("wnd[0]/tbar[0]/okcd").Text = "me33k"
    session.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        MsgBox "Error al abrir la transacción me33k."
        Exit Function
    End If
    On Error GoTo 0

    session.findById("wnd[0]/usr/ctxtRM06E-EVRTN").Text = codigoContrato
    session.findById("wnd[0]").sendVKey 0
    
    Set tabla = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220")
    scrollMax = tabla.verticalScrollbar.Maximum
    totalFilas = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220").rowCount
    posicionScroll = -1 ' Inicialmente inválido para forzar la entrada al loop
    
    Do
    Set tabla = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220")
        On Error Resume Next
        ' Solo continúa si el scroll realmente cambió
        'Debug.Print session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtRM06E-EVRTP[0," & i & "]").Text
        posicionScroll = tabla.verticalScrollbar.Position
        totalFilas = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220").rowCount
        On Error GoTo 0
        
        ' Recorre las filas visibles en la tabla
        For i = 0 To totalFilas - 1
            On Error Resume Next
            Set celdaLOEKZ = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/lblRM06E-LOEKZ[13," & i & "]")
            tooltipTexto = ""
            If Err.Number = 0 And Not celdaLOEKZ Is Nothing Then
                celdaLOEKZ.SetFocus
                celdaLOEKZ.caretPosition = 0
                DoEvents
                tooltipTexto = celdaLOEKZ.Tooltip
            End If
            On Error GoTo 0
            
            If InStr(LCase(Trim(tooltipTexto)), "bloq.") > 0 Then
                valorPrimeraColumna = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtRM06E-EVRTP[0," & i & "]").Text
                resultado.Add valorPrimeraColumna
            End If
        Next i
        
    On Error Resume Next
        ' Avanza el scroll
        session.findById("wnd[0]/usr/tblSAPMM06ETC_0220").verticalScrollbar.Position = posicionScroll + (totalFilas - scrollMax)
        If session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtRM06E-EVRTP[0," & posicionScroll & "]").Text = "_____" Then Exit Do
    On Error GoTo 0
        DoEvents
        
    Loop
    
    Set ObtenerItemsBloqueados = resultado

End Function


Sub LimpiarBloqueados(contrato As String)

ThisWorkbook.application.ScreenUpdating = False

    Dim bloqueados As Collection
    Dim item As Variant
    Dim hoja As Worksheet
    Dim valoresA As Collection
    Dim i As Long
    Dim celda As Range

    Set bloqueados = ObtenerItemsBloqueados(contrato)
    
    ' Validar existencia de la hoja
    On Error Resume Next
    Set hoja = ThisWorkbook.Sheets(contrato)
    On Error GoTo 0

    If hoja Is Nothing Then
        MsgBox "La hoja '" & contrato & "' no existe.", vbExclamation
        Exit Sub
    End If

    'hoja.Activate

    ' Guardar valores originales de la columna A (excepto encabezado)
    Set valoresA = New Collection
    For i = 2 To hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row
        If hoja.Cells(i, 1).value <> "" Then
            valoresA.Add hoja.Cells(i, 1).value
        End If
    Next i
    
    ' Limpiar la columna A antes de aplicar el filtro (vaciar la columna A)
    hoja.Range("A2:A" & hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row).ClearContents
    
    If hoja.AutoFilterMode Then hoja.AutoFilterMode = False

    ' Aplicar filtros y eliminar filas visibles por cada ítem bloqueado
    For Each item In bloqueados

        Dim rangoDatos As Range
        Set rangoDatos = hoja.Range("C1").CurrentRegion
        
        On Error Resume Next
        rangoDatos.AutoFilter Field:=3, Criteria1:=item
        If Err.Number <> 0 Then
            Debug.Print "Error en AutoFilter con item: " & item & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        Dim rangoVisible As Range

        On Error Resume Next
        Set rangoVisible = hoja.AutoFilter.Range.Offset(1, 0).Resize(hoja.AutoFilter.Range.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not rangoVisible Is Nothing Then
            On Error Resume Next
                rangoVisible.EntireRow.Delete
            On Error GoTo 0
        End If

        hoja.AutoFilterMode = False

    Next item

    ' Restaurar los valores originales en columna A
    For i = 1 To valoresA.Count
        hoja.Cells(i + 1, 1).value = valoresA(i)
    Next i

ThisWorkbook.application.ScreenUpdating = True

End Sub
