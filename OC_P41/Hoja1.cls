VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    'Dim HojaName As String
    Dim eleccion As VbMsgBoxResult
    
    ' Verifica si el cambio ocurri� en la celda D7
    If Not Intersect(Target, Me.Range("D7")) Is Nothing Then
        application.EnableEvents = False ' Deshabilita eventos para evitar bucles infinitos
        
        ' Obtiene el nombre para la hoja desde la celda D7
        numCont = Trim(Me.Range("D7").Value)
        
        ' Si la celda est� vac�a, muestra un mensaje y termina
        If numCont = "" Then
            MsgBox "El valor en la celda D7 no puede estar vac�o.", vbExclamation, "Error"
            application.EnableEvents = True
            Exit Sub
        End If
        
        ' Pregunta al usuario si desea realizar una nueva b�squeda
        eleccion = MsgBox("�Desea realizar una nueva gesti�n?", vbYesNo, "Contrato")
        
        If eleccion = vbYes Then
            ' Cambia el nombre de la hoja actual
            On Error Resume Next ' Maneja errores por nombres duplicados o inv�lidos
            Me.Name = numCont
            If Err.Number <> 0 Then
                MsgBox "No se pudo cambiar el nombre de la hoja. Verifique que el nombre no est� duplicado o contenga caracteres no v�lidos.", vbCritical, "Error"
                Err.Clear
            End If
            On Error GoTo 0
            
             ' Limpia los rangos especificados en la hoja actual
            With Me
                .Range("M7").Value = ""
                .Range("D13").Value = ""
                .Range("D11").Value = ""
                .Range("D15").Value = ""
                .Range("K15").Value = ""
                .Range("R15").Value = ""
                .Range("Q21").Value = ""
                .Range("P53").Value = ""
                .Range("J21").Value = ""
                .Range("N17").Value = ""
                .Range("N18").Value = ""
                .Range("J22").Value = ""
                .Range("Q22").Value = ""
            End With
            
            Call VerificarSAP
        End If
        
        application.EnableEvents = True ' Habilita eventos nuevamente
    End If

End Sub

