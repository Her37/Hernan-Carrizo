VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("NewSolp")
    
    ' Verificar que solo se seleccione UNA celda
    If Target.Cells.Count > 1 Then Exit Sub
    
    ' Verificar si la celda modificada es C2
    If Not Intersect(Target, ws.Range("C2")) Is Nothing Then
    
            ' Evitar bucles infinitos al modificar valores
        application.EnableEvents = False
        Select Case Target.Value
            Case "Vigencia", "Monto", "Transferencia", "MMC"
                Me.Range("C7").Value = 28901
            Case "Licitación"
                Me.Range("C7").ClearContents
            Case Else
                Me.Range("C7").ClearContents
        End Select
        application.EnableEvents = True
        
        ' Activar la opción correspondiente en Panel
        Select Case Target.Value
            Case "Transferencia": Panel.Transferencias.Value = True
            Case "Vigencia": Panel.Vigencia.Value = True
            Case "Monto", "MMC", "Solp Complementaria": Panel.Monto.Value = True
            Case "Licitación": Panel.Licitación.Value = True
            Case Else
                MsgBox "Debe seleccionar una opción válida.", vbExclamation, "Aviso"
                Exit Sub
        End Select

    End If
End Sub


