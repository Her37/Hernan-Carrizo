Attribute VB_Name = "Módulo8"
Sub RellenarCeldasVacias(hojanombre As String)
    Dim hoja As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set hoja = ThisWorkbook.Sheets(hojanombre)
    lastRow = hoja.Cells(hoja.Rows.Count, 4).End(xlUp).Row

    For i = 2 To lastRow
        If Not IsError(hoja.Cells(i, 3).value) And Not IsError(hoja.Cells(i, 4).value) Then
            If Trim(CStr(hoja.Cells(i, 3).value)) = "" And Trim(CStr(hoja.Cells(i, 4).value)) <> "" Then
                hoja.Cells(i, 3).value = hoja.Cells(i - 1, 3).value
            End If
        End If
    Next i

    For i = lastRow To 2 Step -1
        If Not IsError(hoja.Cells(i, 4).value) Then
            If Trim(CStr(hoja.Cells(i, 4).value)) = "" Then
                hoja.Rows(i).Delete
            End If
        End If
    Next i

    'MsgBox "finalizado"
End Sub

