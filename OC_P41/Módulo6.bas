Attribute VB_Name = "Módulo6"
'*************************************
'|||REALIZADO POR HERNAN F. CARRIZO|||
'************AGO-SEP 2024*************

Sub VolverAVentanaPrincipalSAP()
    ' Declarar objetos SAP
    Dim SapGuiAuto As Object
    Dim application As Object
    
    ' Obtener la instancia de SAP GUI y la sesión actual
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    On Error GoTo 0
    
    If Not SapGuiAuto Is Nothing Then
        Set application = SapGuiAuto.GetScriptingEngine
        Set connection = application.Children(0) ' Primera conexión activa
        Set session = connection.Children(0) ' Primera sesión activa
        
        ' Volver a la ventana principal usando el botón "Back" (15) repetidamente
        Do While InStr(1, session.findById("wnd[0]").Text, "SAP Easy Access") = 0
            On Error Resume Next
            session.findById("wnd[0]/tbar[0]/btn[15]").press ' Presiona "Back"
            On Error GoTo 0
            
            If Not connection Is Nothing Then
            Exit Sub
            End If
            
            'On Error Resume Next
             ' Manejo de la ventana emergente de confirmación o de guardar cambios
            If session.Children.Count > 1 Then
                ' Verificar si aparece la ventana de confirmación o de guardar cambios
                If session.findById("wnd[1]").Text = "Confirmar" Then
                    ' Presionar el botón "Confirmar"
                    session.findById("wnd[1]/tbar[0]/btn[1]").press
                ElseIf session.findById("wnd[1]").Text = "Finaliz.doc." Then
                    ' Presionar el botón "No" en la ventana de guardar datos (tres botones)
                    session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
                End If
            End If

            
            ' Si aparece la ventana de "Salir del sistema", salir del bucle
            If session.findById("wnd[0]").Text = "Salir del sistema" Then
                Exit Do
            End If
            
            'On Error GoTo 0
            
            If Err.Number <> 0 Then Exit Do ' Salir del bucle si ocurre un error
        Loop
        
        ' Maximizar la ventana una vez que estés en SAP Easy Access
        session.findById("wnd[0]").maximize
    Else
        MsgBox "No se encontró una sesión activa de SAP."
    End If
    
End Sub

Sub PresionarBotonDinamico()
    ' Declarar objetos SAP
    Dim SapGuiAuto As Object
    Dim application As Object
    Dim connection As Object
    Dim session As Object
    Dim contenedor As Object
    Dim boton As Object
    Dim control As Object
    Dim encontrado As Boolean
    Dim i As Integer
    
    encontrado = False
    
    ' Obtener la instancia de SAP GUI y la sesión actual
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
    Set connection = application.Children(0) ' Primera conexión activa
    Set session = connection.Children(0) ' Primera sesión activa
    On Error GoTo 0
    
    ' Verificar si la sesión se obtuvo correctamente
    If session Is Nothing Then
        MsgBox "No se pudo obtener la sesión de SAP.", vbCritical
        Exit Sub
    End If

    ' Intentar localizar el contenedor con IDs dinámicos desde SAPLMEGUI:0000 a SAPLMEGUI:0020
    For i = 0 To 20
        DoEvents
        On Error Resume Next
        ' Construir el ID del subcontenedor dinámicamente
    Set contenedor = session.findById("wnd[0]") ' Localizar la ventana principal
    Set contenedor = contenedor.findById("usr") ' Localizar el contenedor de usuario
    Set contenedor = contenedor.findById("subSUB0:SAPLMEGUI:" & Format(i, "0000"))
    Set contenedor = contenedor.findById("subSUB1:SAPLMEVIEWS:1100") ' Localizar el segundo sub-contenedor
    Set contenedor = contenedor.findById("subSUB1:SAPLMEVIEWS:4000") ' Localizar el segundo sub-contenedor
    'Debug.Print contenedor.ID
        On Error GoTo 0
        
        ' Verificar si el contenedor se encontró
        If Not contenedor Is Nothing Then
            ' Recorre los controles dentro del contenedor en busca del botón
            For Each control In contenedor.Children
                ' Verificar si el control es un botón con el ToolTip correcto
                If control.Type = "GuiButton" Then
                    If control.Tooltip = "Mostar cabecera ctrl+F2" Then
                        control.press
                        encontrado = True
                        'MsgBox "Botón encontrado y presionado en SAPLMEGUI:" & Format(i, "0000"), vbInformation
                        Exit Sub ' Salir del bucle cuando se presione el botón
                    End If
                End If
            Next control
        End If
    Next i

   
End Sub


Sub ListarControles()
    ' Declarar objetos SAP
    Dim SapGuiAuto As Object
    Dim application As Object
    Dim connection As Object
    Dim session As Object
    Dim control As Object
    Dim subControl As Object
    Dim subSubControl As Object
    Dim innerControl As Object

    ' Obtener la instancia de SAP GUI y la sesión actual
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
    Set connection = application.Children(0) ' Primera conexión activa
    Set session = connection.Children(0) ' Primera sesión activa
    On Error GoTo 0

    ' Recorre todos los controles en la ventana principal
    For Each control In session.findById("wnd[0]/usr").Children
        ' Imprime información sobre el control
        Debug.Print "ID: " & control.ID
        Debug.Print "Type: " & control.Type
        Debug.Print "Text: " & control.Text
        Debug.Print "Tooltip: " & control.Tooltip
        'Debug.Print "Class: " & control.Class
        Debug.Print "-----------------------------------"

        ' Si el control es un contenedor, recorre sus hijos
        If control.Type = "GuiSimpleContainer" Then
            For Each subControl In control.Children
                Debug.Print "Sub-Control ID: " & subControl.ID
                Debug.Print "Sub-Control Type: " & subControl.Type
                Debug.Print "Sub-Control Text: " & subControl.Text
                Debug.Print "Sub-Control Tooltip: " & subControl.Tooltip
                'Debug.Print "Sub-Control Class: " & subControl.Class
                Debug.Print "-----------------------------------"

                ' Si el sub-control es un contenedor, recorre sus hijos
                If subControl.Type = "GuiSimpleContainer" Then
                    For Each subSubControl In subControl.Children
                        Debug.Print "Sub-Sub-Control ID: " & subSubControl.ID
                        Debug.Print "Sub-Sub-Control Type: " & subSubControl.Type
                        Debug.Print "Sub-Sub-Control Text: " & subSubControl.Text
                        Debug.Print "Sub-Sub-Control Tooltip: " & subSubControl.Tooltip
                        'Debug.Print "Sub-Sub-Control Class: " & subSubControl.Class
                        Debug.Print "-----------------------------------"

                        ' Si el sub-sub-control es un contenedor, recorre sus hijos
                        If subSubControl.Type = "GuiSimpleContainer" Then
                            For Each innerControl In subSubControl.Children
                                Debug.Print "Inner-Control ID: " & innerControl.ID
                                Debug.Print "Inner-Control Type: " & innerControl.Type
                                Debug.Print "Inner-Control Text: " & innerControl.Text
                                Debug.Print "Inner-Control Tooltip: " & innerControl.Tooltip
                                'Debug.Print "Inner-Control Class: " & innerControl.Class
                                Debug.Print "-----------------------------------"
                            Next innerControl
                        End If
                    Next subSubControl
                End If
            Next subControl
        End If
    Next control
End Sub



