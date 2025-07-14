VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Solpes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Consultar_Click()
 Call VerificarSAP
End Sub


Private Sub CommandButton1_Click()
Call verServicios
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Solpes")
    'Debug.Print Target.Cells.Count
    
    ' Verificar que solo se seleccione UNA celda
    If Target.Cells.Count > 1 Then Exit Sub

    ' Verificar si la selección está en la columna J
    If Not Intersect(Target, ws.Columns("J")) Is Nothing Then
        ' Cambiar el valor del TextBox2 si se selecciona una celda en la columna J
        ThisWorkbook.ActiveSheet.TextBox.Text = Target.Value
    Else
        ' Si no está en la columna J, no hacer nada
        Exit Sub
    End If
End Sub

