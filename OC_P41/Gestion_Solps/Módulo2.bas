Attribute VB_Name = "Módulo2"
'*************************************
'|||REALIZADO POR HERNAN F. CARRIZO|||
'************AGO-SEP 2024*************

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public numCont As String
Public Application As Object
Public Connection As Object
Public Session As Object

Sub VerificarSolp()
    Dim ws As Worksheet
    Dim usuario As String, contrasena As String, valorCampo As String
    Dim sapPath As String
    Dim solpNumber As String
    Dim sapIsOpen As Boolean
    Dim startTime As Double
    Dim timeout As Double

    Set ws = ThisWorkbook.Sheets("Solpes")
    Set TextBox = ws.OLEObjects("TextBox").Object
    Set Check = ws.OLEObjects("Check").Object
    
        ' Mostrar el formulario para ingresar usuario y contraseña
    Inicio.Show vbModal
    'Call Contratos
    

    ' Obtener el valor seleccionado en el ListBox
    solpNumber = Trim(TextBox.Text)
    
    ' Verificar que el número de Solp tenga exactamente 10 dígitos
    If Len(solpNumber) <> 10 Or Not IsNumeric(solpNumber) Then
        MsgBox "Ingrese un N° de Solp correcto.", vbExclamation
        Exit Sub
    End If
    

If Not StartSAP Then
    ' Comprobar si el formulario se cerró correctamente
    If Inicio.Tag = "OK" Then
        ' Verificar si el ComboBox tiene un valor seleccionado
        If solpNumber <> "" Then
            solpNumber = Trim(solpNumber)
       
        usuario = Inicio.txtUsuario.Text
        contrasena = Inicio.txtContraseña.Text
        Else
            MsgBox "Por favor, ingrese un N° de Solp."
            Exit Sub
        End If

        ' Iniciar SAP después de cerrar el formulario
        Call IniciarSAP(usuario, contrasena)

    Else
        ' El formulario se cerró sin guardar las credenciales
        Exit Sub
    End If
End If


    Call VolverAVentanaPrincipalSAP
    
    If ws.Cells(Selection.Rows(1).Row, 3).Value = "" Then
        Call proveedores(Selection.Rows(1).Row)
        Call VolverAVentanaPrincipalSAP
    End If

    
    ' Navegar a la transacción correspondiente y seleccionar el número de Solp
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "me53n"
    Session.findById("wnd[0]").sendVKey 0
    
    ' Consultar el número de Solp ingresado
    Session.findById("wnd[0]/tbar[1]/btn[17]").Press
    Session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN").Text = solpNumber
    Session.findById("wnd[1]/tbar[0]/btn[0]").Press
    
    Call PresionarBotonDinamico

    Dim ventana As Object

    For i = 0 To 100
        On Error Resume Next
    Set ventana = Session.findById("wnd[0]") ' Localizar la ventana principal
    Set ventana = ventana.findById("usr") ' Localizar el contenedor de usuario
    Set ventana = ventana.findById("subSUB0:SAPLMEGUI:" & Format(i, "0000"))
    Set ventana = ventana.findById("subSUB1:SAPLMEVIEWS:1100") ' Localizar el segundo sub-contenedor
    Set ventana = ventana.findById("subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT2") ' Localizar el segundo sub-contenedor
     
 ' Verificar si el contenedor se encontró
    If Not ventana Is Nothing And ventana.ID <> "/app/con[0]/ses[0]/wnd[0]/usr" Then
        ' Verificar si la pestaña está seleccionada
        If ventana.Select Then
        'Else
        End If
        'Debug.Print ventana.ID
        Exit For
    End If
        On Error GoTo 0
    Next i
    
    Dim sapTable As Object
    
    For i = 0 To 100
    ' Capturar la tabla de SAP en sapTable
        On Error Resume Next
    Set sapTable = Session.findById("wnd[0]") ' Localizar la ventana principal
    Set sapTable = sapTable.findById("usr") ' Localizar el contenedor de usuario
    Set sapTable = sapTable.findById("subSUB0:SAPLMEGUI:" & Format(i, "0000"))
    Set sapTable = sapTable.findById("subSUB1:SAPLMEVIEWS:1100") ' Localizar el segundo sub-contenedor
    Set sapTable = sapTable.findById("subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT2/ssubTABSTRIPCONTROL3SUB:SAPLMERELVI:1102/cntlRELEASE_INFO_REQHD/shellcont/shell") ' Localizar el segundo sub-contenedor
    
    If Not sapTable Is Nothing And sapTable.ID <> "/app/con[0]/ses[0]/wnd[0]/usr" Then
        ' Seleccionar la columna "ICON"
        sapTable.selectColumn "ICON"
        Dim filas As Integer
        filas = sapTable.rowCount
        'Debug.Print sapTable.ID
        Exit For
    End If
        On Error GoTo 0
    Next i

    ' Iterar sobre las filas para verificar el tooltip
    For i = 0 To filas - 1
        
        TooltipText = sapTable.getCellTooltip(i, "ICON")
        
        If TooltipText = "Liberación efectuada" Then
            ' Continuar evaluando las siguientes celdas
            Do While TooltipText <> "Es posible liberar" And i < sapTable.rowCount - 1
                i = i + 1
                TooltipText = sapTable.getCellTooltip(i, "ICON")
            Loop
        End If
        
        If TooltipText = "Es posible liberar" Then
            primeraColumnaValor = sapTable.getCellValue(i, "DESCRIPTION")
            'MsgBox "Enviar mail para liberar: " & primeraColumnaValor
            
            For j = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row To 3 Step -1
                If ws.Cells(j, "J").Value = solpNumber Then
                    ws.Cells(j, "K").Value = primeraColumnaValor
                    Exit For
                End If
            Next j
            
            Exit For
        Else
            ws.Cells(j, "K").Value = "Finalizda"
        End If
    Next i

' Mensaje de finalización
'MsgBox "Proceso completado!", vbInformation
Set Connection = Nothing
Set Session = Nothing

End Sub


Function proveedores(i As Integer)
    Dim ws As Worksheet
    Dim contrato As String
    Dim contratista As String
    Dim lastRow As Long
    
    ' Establece la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Solpes")
    
    ' Obtiene la última fila con datos en la columna A
    'lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    

        contrato = ws.Range("B" & i).Value

        ' Verifica si el contrato no está vacío
        If contrato <> "" And Selection.Rows(1).Row > 5 Then
            ' Navega a la transacción ZM50 en SAP
            'Call VolverAVentanaPrincipalSAP
            Session.findById("wnd[0]/tbar[0]/okcd").Text = "zm50"
            Session.findById("wnd[0]").sendVKey 0
            
            ' Ingresa el número de contrato
            Session.findById("wnd[0]/usr/ctxtS_EBELN-LOW").Text = Trim(contrato)
            Session.findById("wnd[0]/tbar[1]/btn[8]").Press
            
            ' Obtiene el nombre del contratista
            On Error Resume Next ' Maneja errores si no se encuentra el elemento
            contratista = Session.findById("wnd[0]/usr/lbl[95,7]").Text
            On Error GoTo 0 ' Restablece el manejo de errores
            
            ' Escribe el contratista en la columna B
            ws.Range("C" & i).Value = contratista
            
            ' Regresa a la pantalla anterior en SAP
            Session.findById("wnd[0]/tbar[0]/btn[3]").Press
        End If
End Function

Sub verServicios()
    Dim ws As Worksheet
    Dim ComboBox As Object, Check As Object, ListBox As Object
    Dim contrato As String
    Dim usuario As String, contrasena As String, valorCampo As String
    
    Set ws = ThisWorkbook.Sheets("Solpes")
    'Set ListBox = ws.OLEObjects("TextBox1").Object
    Set Check = ws.OLEObjects("Check").Object
    'Set ComboBox = ws.OLEObjects("ComboBox1").Object

    ' Obtener el valor seleccionado en el ListBox
    'contrato = Trim(ComboBox.value)

    ' Obtener el valor seleccionado en el ListBox
    contrato = ws.Cells(2, "O").Value
    
If Check.Value = False Then
    ' Mostrar el formulario para ingresar usuario y contraseña
    Inicio.Show vbModal
    'Call Contratos
    ' Comprobar si el formulario se cerró correctamente
    If Inicio.Tag = "OK" Then
        ' Verificar si el ComboBox tiene un valor seleccionado
        If contrato <> "" Then
            'contrato = ws.Cells(3, "Q").value
       
            usuario = Inicio.txtUsuario.Text
            contrasena = Inicio.txtContraseña.Text
        Else
            MsgBox "Por favor, seleccione un Contrato."
            Exit Sub
        End If

        ' Iniciar SAP después de cerrar el formulario
        Call IniciarSAP(usuario, contrasena)
    Else
        ' El formulario se cerró sin guardar las credenciales
        Exit Sub
    End If

    Call VolverAVentanaPrincipalSAP
    
    Session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00007"
    Session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00007"
    Session.findById("wnd[0]/usr/ctxtS_CONTRA-LOW").Text = contrato
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/tbar[1]/btn[8]").Press

    Dim filaSAP As Integer
    Dim filaExcel As Integer
    Dim totalFilas As Integer
    Dim maxFilaVisible As Integer
    
    i = 0
    For i = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row To 6 Step -1
        ws.Cells(i, "N").Value = ""
        ws.Cells(i, "O").Value = ""
        ws.Cells(i, "R").Value = ""
        ws.Cells(i, "S").Value = ""
        ws.Cells(i, "T").Value = ""
        ws.Cells(i, "U").Value = ""
    Next i
    
     Set rng = ws.Range("N6:U" & ws.Cells(ws.Rows.Count, "N").End(xlUp).Row)
    
    ' Aplica bordes alrededor de todo el rango
    With rng.Borders
        .LineStyle = xlNone
    End With
    
    
    ' Asigna la primera fila en SAP y Excel
    filaSAP = 3 ' Primera fila visible en SAP
    filaExcel = 6 ' Primera fila donde se copiarán los datos en Excel

    ' Definir el número de filas visibles en la tabla
    maxFilaVisible = 30 ' Ajusta según cuántas filas son visibles en la tabla de SAP
    ' Bucle hasta que no haya más datos en SAP
    Do While True
        DoEvents
        On Error Resume Next
        ' Intenta copiar el valor en la fila actual de SAP
        Session.findById("wnd[0]/usr").horizontalScrollbar.Position = 0
        ws.Cells(filaExcel, "N").Value = Session.findById("wnd[0]/usr/lbl[111," & filaSAP & "]").Text
        ws.Cells(filaExcel, "U").Value = Session.findById("wnd[0]/usr/lbl[59," & filaSAP & "]").Text
        ws.Cells(filaExcel, "O").Value = Session.findById("wnd[0]/usr/lbl[149," & filaSAP & "]").Text

        ' Si no hay más filas, se detecta un error y se sale del bucle
        If Err.Number <> 0 Then
            Exit Do
            End If
        On Error GoTo 0

        ' Incrementar las filas
        filaSAP = filaSAP + 1
        filaExcel = filaExcel + 1

        ' Si alcanzamos la última fila visible, hacemos scroll hacia abajo
        If filaSAP > maxFilaVisible Then
            Session.findById("wnd[0]/usr").verticalScrollbar.Position = Session.findById("wnd[0]/usr").verticalScrollbar.Position + maxFilaVisible
            'session.findById("wnd[0]/usr").verticalScrollbar.position
            filaSAP = 3 ' Volver a la primera fila visible después del scroll
        End If
    Loop
    
    Session.findById("wnd[0]/usr").verticalScrollbar.Position = 0
    ' Asigna la primera fila en SAP y Excel
    filaSAP = 3 ' Primera fila visible en SAP
    filaExcel = 6 ' Primera fila donde se copiarán los datos en Excel

    ' Definir el número de filas visibles en la tabla
    maxFilaVisible = 30 ' Ajusta según cuántas filas son visibles en la tabla de SAP
    ' Bucle hasta que no haya más datos en SAP
    
        Do While True
        DoEvents
        On Error Resume Next
        ' Intenta copiar el valor en la fila actual de SAP
        Session.findById("wnd[0]/usr").horizontalScrollbar.Position = 106
        ws.Cells(filaExcel, "R").Value = Session.findById("wnd[0]/usr/lbl[99," & filaSAP & "]").Text
        ws.Cells(filaExcel, "S").Value = Session.findById("wnd[0]/usr/lbl[84," & filaSAP & "]").Text
        ws.Cells(filaExcel, "T").Value = Session.findById("wnd[0]/usr/lbl[116," & filaSAP & "]").Text

        ' Si no hay más filas, se detecta un error y se sale del bucle
        If Err.Number <> 0 Then
            Exit Do
            End If
        On Error GoTo 0

        ' Incrementar las filas
        filaSAP = filaSAP + 1
        filaExcel = filaExcel + 1

        ' Si alcanzamos la última fila visible, hacemos scroll hacia abajo
        If filaSAP > maxFilaVisible Then
            Session.findById("wnd[0]/usr").verticalScrollbar.Position = Session.findById("wnd[0]/usr").verticalScrollbar.Position + maxFilaVisible
            'session.findById("wnd[0]/usr").verticalScrollbar.position
            filaSAP = 3 ' Volver a la primera fila visible después del scroll
        End If
    Loop
    
 

    Set rng = ws.Range("N5:U" & filaExcel - 1)
    
    ' Aplica bordes alrededor de todo el rango
    With rng.Borders
        .LineStyle = xlContinuous       ' Estilo de línea continua
        .Color = RGB(0, 0, 0)           ' Color negro (puedes cambiarlo)
        .Weight = xlThin                ' Grosor de la línea (puedes usar xlMedium, xlThick, etc.)
    End With
    
    'session.findById("wnd[0]/tbar[0]/btn[15]").press
    'session.findById("wnd[0]/tbar[0]/btn[15]").press
Else

    Set rng = ws.Range("N6:U" & ws.Cells(ws.Rows.Count, "N").End(xlUp).Row)
    
    ' Aplica bordes alrededor de todo el rango
    With rng.Borders
        .LineStyle = xlNone
    End With
    
        i = 0
        ws.Cells(3, "Q").Value = ""
    For i = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row To 6 Step -1
        ws.Cells(i, "N").Value = ""
        ws.Cells(i, "O").Value = ""
        ws.Cells(i, "R").Value = ""
        ws.Cells(i, "S").Value = ""
        ws.Cells(i, "T").Value = ""
        ws.Cells(i, "U").Value = ""
    Next i
    
End If

End Sub

Sub VerificarSolpes()
    Dim usuario As String, contrasena As String, valorCampo As String
    Dim ws As Worksheet
    Dim sapPath As String
    Dim solpNumber As String
    Dim actCel As Long
    Dim Gestor As String
    Dim Monto As String
    Dim Moneda As String
    Dim txtBreve As String
    Dim Fecha As String
    Dim TxtComp As String
    Dim rng As Range
    
    ManejarVentanaWindows

    Set ws = ThisWorkbook.Sheets("Solpes")
    'Set ComboBox1 = ws.OLEObjects("ComboBox1").Object

    ' Mostrar el formulario para ingresar usuario y contraseña
    Inicio.Show vbModal

    ' Comprobar si el formulario se cerró correctamente
    If Inicio.Tag = "OK" Then
        usuario = Inicio.txtUsuario.Text
        contrasena = Inicio.txtContraseña.Text
        ' Iniciar SAP después de cerrar el formulario
        Call IniciarSAP(usuario, contrasena)
    Else
        ' El formulario se cerró sin guardar las credenciales
        Exit Sub
    End If

   Call VolverAVentanaPrincipalSAP
    
    Dim celda As Range
    For Each celda In Selection
        If ws.Cells(celda.Row, 3).Value = "" Then
            Call proveedores(celda.Row)
            Call VolverAVentanaPrincipalSAP
            'Exit For ' Sale del bucle al encontrar la primera celda vacía
        End If
    Next celda

    
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "me53n"
    Session.findById("wnd[0]").sendVKey 0

    actCel = ActiveCell.Row

    For Each celda In Selection
        solpNumber = Trim(ws.Cells(celda.Row, "J").Value)
        
        ' Verificar que el número de Solp tenga exactamente 10 dígitos
        If Len(solpNumber) <> 10 Or Not IsNumeric(solpNumber) Then
            MsgBox "El número de Solp en la Fila " & celda.Row & " no es correcto. No se pudo verificar.", vbExclamation
            GoTo NextIteration
        End If
        
         ' Verificar que el número de Solp tenga exactamente 10 dígitos
        'If ws.Cells(j, "L").Value = "ADJ" Then
            'GoTo NextIteration
        'End If
        
        'Dim bPopupDetected As Boolean
        'bPopupDetected = False
        

        Session.findById("wnd[0]/tbar[1]/btn[17]").Press
        Session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN").Text = solpNumber
        Session.findById("wnd[1]/tbar[0]/btn[0]").Press
             
        Call PresionarBotonDinamico

        Dim ventana As Object
                       
        For i = 0 To 100
    
            On Error Resume Next
            Set ventana = Session.findById("wnd[0]") ' Localizar la ventana principal
            Set ventana = ventana.findById("usr") ' Localizar el contenedor de usuario
            Set ventana = ventana.findById("subSUB0:SAPLMEGUI:" & Format(i, "0000"))
            Set ventana = ventana.findById("subSUB1:SAPLMEVIEWS:1100") ' Localizar el segundo sub-contenedor
            Set ventana = ventana.findById("subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT2") ' Localizar el segundo sub-contenedor
            On Error GoTo 0
            If Not ventana Is Nothing And ventana.ID <> "/app/con[0]/ses[0]/wnd[0]/usr" Then
                If ventana.Select Then Exit For
            End If
        Next i

        Dim sapTable As Object

        For i = 0 To 100

            On Error Resume Next
            Set sapTable = Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB1:SAPLMEVIEWS:1100/" & _
                                            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT2/" & _
                                            "ssubTABSTRIPCONTROL3SUB:SAPLMERELVI:1102/cntlRELEASE_INFO_REQHD/shellcont/shell")
            On Error GoTo 0
            If Not sapTable Is Nothing Then
                sapTable.selectColumn "ICON"
                Exit For
            End If
        Next i

    ' Iterar sobre las filas para verificar el tooltip
    For i = 0 To sapTable.rowCount - 1
            TooltipText = sapTable.getCellTooltip(i, "ICON")
            If TooltipText = "Liberación efectuada" Then
                Do While TooltipText <> "Es posible liberar" And i < sapTable.rowCount - 1
                    i = i + 1
                    TooltipText = sapTable.getCellTooltip(i, "ICON")
                Loop
            End If
            If TooltipText = "Es posible liberar" Then
                Dim primeraColumnaValor As String
                primeraColumnaValor = sapTable.getCellValue(i, "DESCRIPTION")
                ws.Cells(celda.Row, "K").Value = primeraColumnaValor
                ws.Cells(celda.Row, "L").Value = ConvertirValor(ws.Cells(celda.Row, "K").Value)
                Exit For
            Else
                ws.Cells(celda.Row, "K").Value = "Adjudicada"
                ws.Cells(celda.Row, "L").Value = ConvertirValor(ws.Cells(celda.Row, "K").Value)
            End If
    Next i

        
        Call PresionarBotonDinamico2
        
        'Persona de contacto
        For i = 0 To 20
            On Error Resume Next
            Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB3:SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT11").Select
            On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i
                                 
        'Persona de contacto
        For i = 0 To 20
            On Error Resume Next
        Gestor = Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB3:SAPLMEVIEWS:1100/" & _
                                  "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                                  "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT11/" & _
                                  "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:3328/txtMEREQ3328-EKNAM").Text
            On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i
        
        ws.Cells(celda.Row, "M").Value = Gestor
        
 If ws.Cells(celda.Row, "G").Value = "" Or ws.Cells(celda.Row, "H").Value = "" Or ws.Cells(celda.Row, "M").Value = "" Or ws.Cells(celda.Row, "D").Value = "" Or ws.Cells(celda.Row, "F").Value = "" Then
       
        'Descripcion texto
        Call PresionarBotonDinamico
        Call PresionarBotonDinamico4
        
        TxtComp = ""
        For i = 0 To 20
            On Error Resume Next
        TxtComp = Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB1:SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT1/ssubTABSTRIPCONTROL3SUB:SAPLMEGUI:1230/" & _
                "subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").Text
            On Error GoTo 0
        If TxtComp <> "" Then
            Err.Clear
            Exit For
        End If
        Next i

        'Fecha
        For i = 0 To 20
            On Error Resume Next
            Fecha = Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/" & _
                         "tabpTABREQDT11/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:3328/ctxtMEREQ3328-ERDAT").Text
            On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i
        
        dia = Left(Fecha, 2)
        mes = Mid(Fecha, 4, 2)
        anio = Right(Fecha, 4)
        fechaConvertida = DateSerial(CLng(anio), CLng(mes), CLng(dia))
        'fecha ultima modificacion aumento monto
        ws.Cells(celda.Row, "E").Value = fechaConvertida
        
            
        'Datos Clientes
        For i = 0 To 20
            On Error Resume Next
        Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                         "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15").Select
            On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i
        
        'Monto/moneda para Transferecnias
        For i = 0 To 20
            On Error Resume Next
        Monto = Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/ssubSUB2:SAPLXM02:9000/txtEBAN-ZMONTO1").Text
            On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i
        
        For i = 0 To 20
            On Error Resume Next
        Moneda = Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/ssubSUB2:SAPLXM02:9000/ctxtEKKO-WAERS").Text
            On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i

            'Monto/moneda
        For i = 0 To 20
            On Error Resume Next
            Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                         "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/" & _
                         "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/" & _
                         "tabsGRILLA/tabpPUSH4").Select
            On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i
        
        For i = 0 To 20
            On Error Resume Next
        Monto = Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB3:SAPLMEVIEWS:1100/" & _
                                 "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                                 "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/" & _
                                 "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/" & _
                                 "tabsGRILLA/tabpPUSH4/ssubSUB4:SAPLXM02:9004/txtEBAN-ZZMOSOL").Text
            On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i
        
        For i = 0 To 20
            On Error Resume Next
        Moneda = Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB3:SAPLMEVIEWS:1100/" & _
                                  "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
                                  "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/" & _
                                  "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/" & _
                                  "tabsGRILLA/tabpPUSH4/ssubSUB4:SAPLXM02:9004/ctxtEBAN-ZZMON").Text
            On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i
        
    If Moneda = "" Then
        For i = 0 To 20
            On Error Resume Next
        Moneda = Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB3:SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/" & _
                "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/" & _
                "ssubSUB2:SAPLXM02:9000/ctxtEKKO-WAERS").Text
            On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i
    End If
             
        For i = 0 To 20
            On Error Resume Next
        Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
            "subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "TXZ01"
            On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i
        
        Call PresionarBotonDinamico3

        For i = 0 To 20
            On Error Resume Next
        txtBreve = Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                    "subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").getCellValue(0, "TXZ01")
            On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i
        
        ws.Cells(celda.Row, "G").Value = Monto
        ws.Cells(celda.Row, "H").Value = Moneda
        ws.Cells(celda.Row, "D").Value = txtBreve
        ws.Cells(celda.Row, "F").Value = TxtComp
        
        'Vuelve a la pestañña texto
        For i = 0 To 20
            On Error Resume Next
        Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
            "subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT1").Select
             On Error GoTo 0
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
        Next i
End If

    Set rng = ws.Range("B6:M" & ws.Cells(ws.Rows.Count, "J").End(xlUp).Row)
    
    ' Aplica bordes alrededor de todo el rango
    With rng.Borders
        .LineStyle = xlContinuous       ' Estilo de línea continua
        .Color = RGB(0, 0, 0)           ' Color negro (puedes cambiarlo)
        .Weight = xlThin                ' Grosor de la línea (puedes usar xlMedium, xlThick, etc.)
    End With
    

NextIteration:
    Next celda

    Session.findById("wnd[0]/tbar[0]/btn[15]").Press

    MsgBox "Solp's actualizadas", vbExclamation

End Sub


Function ConvertirValor(valor As String) As String
    Select Case valor
        Case "Oficina de Contratos": ConvertirValor = "OC"
        Case "Gerente (Genérico)": ConvertirValor = "GT"
        Case "Jefe Departamento": ConvertirValor = "JD"
        Case "Director": ConvertirValor = "DI"
        Case "Presupuesto para SRV": ConvertirValor = "PR"
        Case "Adjudicada": ConvertirValor = "ADJ"
        Case Else: ConvertirValor = ""
    End Select
End Function


Sub ManejarVentanaWindows()

    Dim hwnd As Long
    Const WindowTitle As String = "SAP GUI for Windows 800" ' <--- ¡IMPORTANTE: REEMPLAZA CON EL TÍTULO EXACTO DE TU VENTANA!

    ' Pausa para asegurar que la ventana haya aparecido completamente
    Sleep 500 ' 0.5 segundos

    ' Buscar la ventana por su título
    ' El título debe ser EXACTO (mayúsculas, minúsculas, espacios)
    hwnd = FindWindow(vbNullString, WindowTitle)

    If hwnd <> 0 Then ' Si la ventana fue encontrada
        ' Poner la ventana al frente para asegurar que reciba los SendKeys
        SetForegroundWindow hwnd
        Sleep 100 ' Pequeña pausa para asegurar el enfoque

        ' Enviar una tecla, por ejemplo, ENTER para presionar el botón por defecto (OK/Sí)
        'ThisWorkbook.Application.SendKeys "{ENTER}"
        ' O para ESC para cancelar:
        ThisWorkbook.Application.SendKeys "{ESC}"

        'MsgBox "Ventana de Windows con título '" & WindowTitle & "' detectada y acción enviada.", vbInformation
    Else
        'MsgBox "La ventana de Windows con título '" & WindowTitle & "' NO fue encontrada.", vbExclamation
    End If

End Sub

