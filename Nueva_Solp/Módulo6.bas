Attribute VB_Name = "Módulo6"
'*************************************
'|||REALIZADO POR HERNAN F. CARRIZO|||
'************AGO-SEP 2024*************
Function ZM50() As Variant
    Dim centro As String
    Dim Moneda As String
    Dim GM As String
    Dim ws As Worksheet
    Dim contrato As String
    
    Call VolverAVentanaPrincipalSAP
    
    Set ws = ThisWorkbook.Sheets("NewSolp")
    'contrato = Trim(ws.Cells(9, 6).Value)
     contrato = Panel.TextBox1.Value
    
    session.findById("wnd[0]/tbar[0]/okcd").text = "zm50"
    session.findById("wnd[0]").sendVKey 0
    
    If contrato = "" Then
        MsgBox "Ingrese un numeor de contrato."
        Exit Function
    End If
    
    session.findById("wnd[0]/usr/ctxtS_EBELN-LOW").text = Trim(contrato)
    session.findById("wnd[0]/tbar[1]/btn[8]").press


    GM = session.findById("wnd[0]/usr/lbl[74,7]").text
    Moneda = session.findById("wnd[0]/usr/lbl[88,9]").text
    centro = session.findById("wnd[0]/usr/lbl[121,9]").text
    
    ws.Range("C6").Value = GM
    ws.Range("F7").Value = Moneda
    ws.Range("F11").Value = centro
    
    Panel.TextBox35.Value = GM
    Panel.TextBox33.Value = Moneda
    Panel.TextBox34.Value = centro
    
    ZM50 = Array(GM, Moneda, centro)
    
    
    End Function
    
Function ZCO9() As Variant
    Dim Ccosto As String
    Dim PEP As String
    Dim ws As Worksheet
    Dim contrato As String
    
    Call VolverAVentanaPrincipalSAP
    
    Set ws = ThisWorkbook.Sheets("NewSolp")
    'contrato = Trim(ws.Cells(9, 6).Value)
     contrato = Panel.TextBox1.Value
    ' Maximizar la ventana de SAP
    'session.findById("wnd[0]").maximize

    session.findById("wnd[0]/tbar[0]/okcd").text = "zco9"
    session.findById("wnd[0]").sendVKey 0
    
    If contrato = "" Or Len(contrato) < 10 Then
        MsgBox "Ingrese un número de contrato valido."
        Exit Function
    End If
    
    On Error Resume Next
    session.findById("wnd[0]/usr/ctxtSE_KONNR-LOW").text = Trim(contrato)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Verificar si aparece la ventana de error (wnd[1])
    Set ventanaError = session.findById("wnd[1]", False)
    
    ' Si la ventana de error existe, presionar "OK" y salir del procedimiento
    If Not ventanaError Is Nothing Then
        MsgBox "No existen certificados para el Contrato propuesto.", vbExclamation, "Aviso"
        ventanaError.findById("tbar[0]/btn[0]").press
        Exit Function
    End If
    On Error GoTo 0

    session.findById("wnd[0]").sendVKey 82
    session.findById("wnd[0]").sendVKey 83
    session.findById("wnd[0]/usr/lbl[9,4]").SetFocus
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/usr/lbl[2,7]").SetFocus
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    
    On Error Resume Next
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT12").Select
    
    If Err.Number <> 0 Then
        Call PresionarBotonDinamico2
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT12").Select
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Obtener el valor de la celda específica
    On Error Resume Next
    Ccosto = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
    "subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT12/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/" & _
    "subSUB1:SAPLMEACCTVI:1100/ctxtMEACCT1100-SAKTO").text

    PEP = session.findById( _
        "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
        "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT12/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/" & _
        "subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-PS_POSID").text

    If Err.Number <> 0 Then
        Ccosto = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
        "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT12/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/" & _
        "subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1000/tblSAPLMEACCTVIDYN_1000TC/ctxtMEACCT1000-SAKTO[5,0]").text
        
        Dim NOrden As String
        NOrden = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
        "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT12/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/" & _
        "subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1000/tblSAPLMEACCTVIDYN_1000TC/ctxtMEACCT1000-AUFNR[7,0]").text
        
        Call VolverAVentanaPrincipalSAP
        Dim resultado As String
        PEP = BuscarPEP(NOrden) ' Guardar el valor retornado en una variable
        
            ws.Range("H2").Value = NOrden
            Panel.TextBox42.Value = PEP
            ws.Range("H3").Value = PEP
            Panel.TextBox32.Value = Ccosto
            ws.Range("F3").Value = Ccosto
    Else
            Panel.TextBox31.Value = PEP
            ws.Range("F2").Value = PEP
            Panel.TextBox32.Value = Ccosto
            ws.Range("F3").Value = Ccosto
    End If
        Err.Clear
    On Error GoTo 0
    
    ' Presionar botones
    'session.findById("wnd[0]/tbar[0]/btn[15]").press
    'session.findById("wnd[0]/tbar[0]/btn[15]").press
    'session.findById("wnd[0]/tbar[0]/btn[15]").press
    'session.findById("wnd[0]/tbar[0]/btn[15]").press
    
    ZCO9 = Array(Ccosto, PEP)

End Function

Function BuscarPEP(NOrden As String) As String
    session.findById("wnd[0]/tbar[0]/okcd").text = "iw33"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = NOrden
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKD").Select
    PEP = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKD/ssubSUB_AUFTRAG:SAPLCOIH:1130/ctxtCAUFVD-PSPEL").text
    
    BuscarPEP = PEP
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    session.findById("wnd[0]/tbar[0]/btn[15]").press
End Function

Sub VolverAVentanaPrincipalSAP()
    ' Declarar objetos SAP
    'Dim sapGuiAuto As Object
    'Dim application As Object
    'Dim connection As Object
    'Dim session As Object
    
    
    ' Obtener la instancia de SAP GUI y la sesión actual
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    On Error GoTo 0
    
    If Not SapGuiAuto Is Nothing Then
        Set application = SapGuiAuto.GetScriptingEngine
        Set connection = application.Children(0) ' Primera conexión activa
        Set session = connection.Children(0) ' Primera sesión activa
        
        ' Volver a la ventana principal usando el botón "Back" (15) repetidamente
        Do While InStr(1, session.findById("wnd[0]").text, "SAP Easy Access") = 0
            On Error Resume Next
            session.findById("wnd[0]/tbar[0]/btn[15]").press ' Presiona "Back"
            On Error GoTo 0
            
             ' Manejo de la ventana emergente de confirmación o de guardar cambios
            If session.Children.Count > 1 Then
                ' Verificar si aparece la ventana de confirmación o de guardar cambios
                If session.findById("wnd[1]").text = "Confirmar" Then
                    ' Presionar el botón "Confirmar"
                    session.findById("wnd[1]/tbar[0]/btn[1]").press
                ElseIf session.findById("wnd[1]").text = "Finaliz.doc." Then
                    ' Presionar el botón "No" en la ventana de guardar datos (tres botones)
                    session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
                End If
            End If
            
            ' Si aparece la ventana de "Salir del sistema", salir del bucle
            If session.findById("wnd[0]").text = "Salir del sistema" Then
                Exit Do
            End If
            
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
        Debug.Print "Text: " & control.text
        Debug.Print "Tooltip: " & control.Tooltip
        'Debug.Print "Class: " & control.Class
        Debug.Print "-----------------------------------"

        ' Si el control es un contenedor, recorre sus hijos
        If control.Type = "GuiSimpleContainer" Then
            For Each subControl In control.Children
                Debug.Print "Sub-Control ID: " & subControl.ID
                Debug.Print "Sub-Control Type: " & subControl.Type
                Debug.Print "Sub-Control Text: " & subControl.text
                Debug.Print "Sub-Control Tooltip: " & subControl.Tooltip
                'Debug.Print "Sub-Control Class: " & subControl.Class
                Debug.Print "-----------------------------------"

                ' Si el sub-control es un contenedor, recorre sus hijos
                If subControl.Type = "GuiSimpleContainer" Then
                    For Each subSubControl In subControl.Children
                        Debug.Print "Sub-Sub-Control ID: " & subSubControl.ID
                        Debug.Print "Sub-Sub-Control Type: " & subSubControl.Type
                        Debug.Print "Sub-Sub-Control Text: " & subSubControl.text
                        Debug.Print "Sub-Sub-Control Tooltip: " & subSubControl.Tooltip
                        'Debug.Print "Sub-Sub-Control Class: " & subSubControl.Class
                        Debug.Print "-----------------------------------"

                        ' Si el sub-sub-control es un contenedor, recorre sus hijos
                        If subSubControl.Type = "GuiSimpleContainer" Then
                            For Each innerControl In subSubControl.Children
                                Debug.Print "Inner-Control ID: " & innerControl.ID
                                Debug.Print "Inner-Control Type: " & innerControl.Type
                                Debug.Print "Inner-Control Text: " & innerControl.text
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

Sub PresionarBotonDinamico2()
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
    Set contenedor = contenedor.findById("subSUB3:SAPLMEVIEWS:1100") ' Localizar el segundo sub-contenedor
    Set contenedor = contenedor.findById("subSUB1:SAPLMEVIEWS:4002") ' Localizar el segundo sub-contenedor
    Debug.Print contenedor.ID
        On Error GoTo 0
        'session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON").press
       
       ' session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON").press
        
        ' Verificar si el contenedor se encontró
        If Not contenedor Is Nothing Then
            ' Recorre los controles dentro del contenedor en busca del botón
            For Each control In contenedor.Children
                ' Verificar si el control es un botón con el ToolTip correcto
                If control.Type = "GuiButton" Then
                    If control.Tooltip = "Mostar detalle pos.ctrl+F4" Then
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


Sub PresionarBotonDinamico3()
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
    Set contenedor = contenedor.findById("subSUB2:SAPLMEVIEWS:1100") ' Localizar el segundo sub-contenedor
    Set contenedor = contenedor.findById("subSUB1:SAPLMEVIEWS:4001") ' Localizar el segundo sub-contenedor
    'Debug.Print contenedor.ID
        On Error GoTo 0
        
        'session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4001/btnDYN_4000-BUTTON").press
        
        ' Verificar si el contenedor se encontró
        If Not contenedor Is Nothing Then
            ' Recorre los controles dentro del contenedor en busca del botón
            For Each control In contenedor.Children
                ' Verificar si el control es un botón con el ToolTip correcto
                If control.Type = "GuiButton" Then
                    If control.Tooltip = "Mostar posiciones ctrl+F3" Then
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



