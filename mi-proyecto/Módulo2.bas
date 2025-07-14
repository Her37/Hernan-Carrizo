Attribute VB_Name = "Módulo2"
'*************************************
'|||REALIZADO POR HERNAN F. CARRIZO|||
'************MAY-JUL 2024*************

Option Explicit

Public nombreHoja As String
Public sapGuiAuto As Object
Public application As Object
Public session As Object
Public connection As Object

Sub Botón2_Haga_clic_en()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim wsS As Worksheet
    Dim wsB As Worksheet
    Dim rangoOrigen As Range
    Dim rangoOrigen2 As Range
    Dim Election As Integer
    Dim Election2 As Integer
    Dim rutaArchivo As Variant
    Dim libro As Workbook

    nombreHoja = "zm50" ' & Format(Date, "DD-MM-YY")

    ' Define las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets(4)
    Set wsDestino = ThisWorkbook.Sheets(1)
    Set wsB = ThisWorkbook.Sheets(3)

    ' Desactiva todos los filtros
    wsOrigen.AutoFilterMode = False
    wsDestino.AutoFilterMode = False

    ' Define el rango de origen
    Set rangoOrigen = wsOrigen.Range("A:AJ")

    Election2 = MsgBox("Bienvendio a Gestión de Contratos: ¿Desea Iniciar una Nueva Gestión?", vbQuestion + vbYesNoCancel, "Work&Service")

    Select Case Election2
        Case vbYes
        
        Election = MsgBox("Ya realizó la descraga SAP del reporte zm50?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestión de Contratos")
   
        If Election = vbNo Then
    
        Call DescargaSAPzm50
        Exit Sub
            
        ElseIf Election = vbYes Then
        
            MsgBox "Por favor, seleccione el reporte zm50.", vbInformation, "Nueva Gestión"

            ' Solicitar al usuario que seleccione el archivo
            rutaArchivo = ThisWorkbook.application.GetOpenFilename("Archivos de Excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", , "Seleccione el archivo de Excel")

            If rutaArchivo = False Then
                MsgBox "No se seleccionó ningún archivo.", vbExclamation
                Exit Sub
            End If

            ' Abrir el archivo seleccionado
            Set libro = Workbooks.Open(rutaArchivo)

            ' Configurar hoja de trabajo
            Set wsS = libro.Sheets(1)
            Set rangoOrigen2 = wsS.Range("A:AJ")

            rangoOrigen2.Copy Destination:=wsOrigen.Range("A1")
            libro.Close SaveChanges:=False

            ' Copiar los datos del rango de origen al rango de destino
            rangoOrigen.Copy Destination:=wsOrigen.Range("A1")
            wsOrigen.Name = nombreHoja
        End If

        Case vbNo
            Call Cont_Act
            Call Cont_Venc
            UserForm1.ComboBox1.Value = wsB.Range("AG1").Value
        Case Else
            Exit Sub
    End Select
    
    UserForm1.Show
End Sub


Sub Cont_Venc()
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim ListItem As ListItem
    Dim i As Integer
    
    ' Configurar la hoja de destino
    Set wsDestino = ThisWorkbook.Sheets("Grids")
    ultimaFila = wsDestino.Cells(wsDestino.Rows.Count, "B").End(xlUp).Row
    
    UserForm1.CheckBox1.Value = False

' Configurar el ListView2
With UserForm1.ListView2
    .View = lvwReport
    .Gridlines = True
    .FullRowSelect = True
    .HideSelection = False
    .ColumnHeaders.Clear
    .ListItems.Clear

    ' Agregar los encabezados manualmente
    .ColumnHeaders.Add , , "Clase", 30
    .ColumnHeaders.Add , , "Contrato", 60
    .ColumnHeaders.Add , , "Descripción", 180
    .ColumnHeaders.Add , , "G.Merc.", 40
    .ColumnHeaders.Add , , "Proveedor", 70
    .ColumnHeaders.Add , , "F. Hasta", 50
End With

' Recorrer las filas y agregar los datos al ListView2
For i = ultimaFila To 2 Step -1
    If wsDestino.Cells(i, 3).Interior.Color = RGB(255, 0, 0) Then
        With UserForm1.ListView2
            Set ListItem = .ListItems.Add(, , wsDestino.Cells(i, 1).Value)
            ListItem.ListSubItems.Add , , wsDestino.Cells(i, 2).Value
            ListItem.ListSubItems.Add , , wsDestino.Cells(i, 4).Value
            ListItem.ListSubItems.Add , , wsDestino.Cells(i, 6).Value
            ListItem.ListSubItems.Add , , wsDestino.Cells(i, 8).Value
            ListItem.ListSubItems.Add , , wsDestino.Cells(i, 17).Value
        End With
    End If
Next i

' Actualizar el título del Frame2
UserForm1.Frame2.Caption = "Contratos Vencidos Actuales: " & UserForm1.ListView2.ListItems.Count

End Sub

Sub Cont_Act()
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim i As Integer
    Dim ListItem As ListItem

    ' Configurar la hoja de destino
    Set wsDestino = ThisWorkbook.Sheets("Grids")
    ultimaFila = wsDestino.Cells(wsDestino.Rows.Count, "B").End(xlUp).Row
    
    UserForm1.CheckBox2.Value = False
    
 'Configurar el ListView2
    With UserForm1.ListView1
        .View = lvwReport
        .Gridlines = True
        .Font = 9
        .FullRowSelect = True
        .HideSelection = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
    ' Agregar los encabezados manualmente
    .ColumnHeaders.Add , , "Clase", 30
    .ColumnHeaders.Add , , "Contrato", 60
    .ColumnHeaders.Add , , "Descripción", 180
    .ColumnHeaders.Add , , "G.Merc.", 40
    .ColumnHeaders.Add , , "Proveedor", 70
    .ColumnHeaders.Add , , "F. Desde", 50

    End With
 ' Recorrer las filas y agregar los datos al ListView1
    For i = ultimaFila To 2 Step -1
        If wsDestino.Cells(i, 3).Interior.Color = RGB(255, 255, 0) Then
            ' Agregar un nuevo ítem al ListView1
            Set ListItem = UserForm1.ListView1.ListItems.Add(, , wsDestino.Cells(i, 1).Value)
            ListItem.ListSubItems.Add , , wsDestino.Cells(i, 2).Value
            ListItem.ListSubItems.Add , , wsDestino.Cells(i, 4).Value
            ListItem.ListSubItems.Add , , wsDestino.Cells(i, 6).Value
            ListItem.ListSubItems.Add , , wsDestino.Cells(i, 8).Value
            ListItem.ListSubItems.Add , , wsDestino.Cells(i, 16).Value
        End If

    Next i

    ' Actualizar el título del Frame1
    UserForm1.Frame1.Caption = "Contratos Nuevos: " & UserForm1.ListView1.ListItems.Count
    
End Sub

Sub IniciarSAP()
    Dim loginSuccess As Boolean
    Dim sapPath As String
    Dim sapIsOpen As Boolean
    Dim startTime As Double
    Dim timeout As Double
    
    ' Ruta del archivo ejecutable de SAP Logon
    sapPath = "C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe"
    
    ' Verificar si el archivo SAP Logon existe en la ruta especificada
    If Dir(sapPath) = "" Then
        MsgBox "No se encuentra el archivo SAP Logon en la ruta especificada."
        Exit Sub
    End If
    
    ' Ejecutar el archivo de SAP Logon para abrir la aplicación
    Shell sapPath, vbNormalFocus
    
    ' Establecer un tiempo límite para esperar que SAP Logon se abra (por ejemplo, 60 segundos)
    timeout = 60 ' Tiempo de espera
    startTime = Timer
    sapIsOpen = False
    
    ' Verificar continuamente si SAP GUI está disponible
    Do
        Set sapGuiAuto = Nothing
        On Error Resume Next
        Set sapGuiAuto = GetObject("SAPGUI")
        On Error GoTo 0
        If Not sapGuiAuto Is Nothing Then
            sapIsOpen = True
            Exit Do
        End If
        
        ' Salir si el tiempo de espera excede el límite
        If Timer - startTime > timeout Then
            MsgBox "SAP Logon tardó demasiado en abrirse."
            Exit Sub
        End If
        
        DoEvents ' Permitir que el sistema procese otros eventos
    Loop
    
    ' Obtener el motor de scripting de SAP GUI
    Set application = sapGuiAuto.GetScriptingEngine
    
    ' Verificar si se pudo obtener la instancia de SAP GUI
    If application Is Nothing Then
        MsgBox "No se pudo obtener la instancia de SAP GUI. Asegúrate de que SAP GUI esté abierto o que el scripting esté habilitado."
        Exit Sub
    End If
    
    ' Intentar abrir una conexión si no hay ninguna conexión activa
    If application.Children.Count = 0 Then
        ' Abre una conexión al servidor SAP especificado (ajusta el nombre de la conexión)
          
    ' Intentar establecer la conexión
    Set connection = application.OpenConnection("H172 C11 [SAP] - Producción Link", True)
        
    Else
        ' Si ya hay conexiones activas, usar la primera
        Set connection = application.Children(0)
    End If
    
        ' Verificar si se pudo establecer la conexión
    If connection Is Nothing Then
        ' Si hay un error (por ejemplo, la conexión no se puede establecer), muestra un mensaje y sale del programa
        MsgBox "No se pudo establecer la conexión con SAP."
        Set connection = Nothing
        Set session = Nothing
        Exit Sub
    End If

    
    ' Verificar si ya hay una sesión activa
    If connection.Children.Count > 0 Then
        Set session = connection.Children(0)
        
        ' Comprobar si estamos en la pantalla de login (campo de usuario)
        If Not session.findById("wnd[0]/usr/txtRSYST-BNAME", False) Is Nothing Then
            ' Estamos en la pantalla de login, proceder con el logueo
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "AR31591057"
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "SAPservicio%2025"
            session.findById("wnd[0]").sendVKey 0
 
             ' Verificar si el login fue exitoso
            On Error Resume Next
            loginSuccess = session.findById("wnd[0]/usr/txtRSYST-BNAME", False) Is Nothing
            On Error GoTo 0
            If Not loginSuccess Then
                MsgBox "Usuario o contraseña inválidos. Verifica tus credenciales."
                Set connection = Nothing
                Set session = Nothing
                Exit Sub
            End If
            
        End If
    Else
        MsgBox "No se encontró ninguna sesión activa en la conexión."
        Set application = Nothing
        Set connection = Nothing
        Set session = Nothing
        Exit Sub
    End If
    

End Sub

Sub DescargaSAPzm50()
    Dim Eleccion As Integer
        
    Eleccion = MsgBox("La Descarga de zm50 se ejecutara como proceso de fondo desde SAP. Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Descarga zm50")
    
If Eleccion = vbYes Then
    
    MsgBox "Espere a que SAP finalice la descarga del reporte zm50, guardelo y  vuleva a iniciar la gestión.", vbInformation, "Nueva Gestión"
    
    Call VolverAVentanaPrincipalSAP
    Call IniciarSAP
    
    If Not session Is Nothing Then
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").Text = "zm50"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[1]/btn[17]").press
        session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
        session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
        session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
        session.findById("wnd[1]/tbar[0]/btn[8]").press
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell 7, "TEXT"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "7"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
        session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select
        session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "VPN1"
        session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").caretPosition = 4
        session.findById("wnd[1]").sendVKey 0
        session.findById("wnd[2]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[13]").press
        session.findById("wnd[1]/usr/btnSOFORT_PUSH").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[11]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
    End If
    
ElseIf Eleccion = vbNo Then
    Exit Sub
End If

End Sub

Sub VolverAVentanaPrincipalSAP()
    ' Obtener la instancia de SAP GUI y la sesión actual
    On Error Resume Next
    Set sapGuiAuto = GetObject("SAPGUI")
    On Error GoTo 0
    
    ' Verificar si se pudo establecer la conexión
    If connection Is Nothing Then
        Set application = Nothing
        Set connection = Nothing
        Set session = Nothing
        Exit Sub
    End If
    
    If Not sapGuiAuto Is Nothing Then
        Set application = sapGuiAuto.GetScriptingEngine
        Set connection = application.Children(0) ' Primera conexión activa
        Set session = connection.Children(0) ' Primera sesión activa
        
        ' Volver a la ventana principal usando el botón "Back" (15) repetidamente
        Do While InStr(1, session.findById("wnd[0]").Text, "SAP Easy Access") = 0
            On Error Resume Next
            session.findById("wnd[0]/tbar[0]/btn[15]").press ' Presiona "Back"
            On Error GoTo 0
            
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
            
            If Err.Number <> 0 Then Exit Do ' Salir del bucle si ocurre un error
        Loop
        
        ' Maximizar la ventana una vez que estés en SAP Easy Access
        session.findById("wnd[0]").maximize
    Else
        MsgBox "No se encontró una sesión activa de SAP."
    End If
End Sub

Sub DescargaSAPzco9()
    Dim nombreArchivo As String
    Dim libro As Workbook
    Dim rutaArchivo As String
    Dim Eleccion As Integer
    Dim wsB As Worksheet
    Dim wsco9 As Worksheet
    Dim lastRow As Long
    Dim rangoOrigen As Range
    
    ' Acceder a la hoja de Excel y la columna deseada
    Set wsB = ThisWorkbook.Sheets("Base Trabajo")
    lastRow = wsB.Cells(wsB.Rows.Count, "A").End(xlUp).Row
    
    MsgBox "Espere a que SAP finalice la descarga del reporte zco9, guárdelo y vuelva a iniciar la gestión.", vbInformation, "Nueva Gestión"
    
    Call VolverAVentanaPrincipalSAP
    Call IniciarSAP
    
    ' Validar que la sesión de SAP está activa
    If session Is Nothing Then
        MsgBox "No se encontró una sesión activa de SAP.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Ejecutar transacción ZCO9 en SAP
    With session
        .findById("wnd[0]").maximize
        .findById("wnd[0]/tbar[0]/okcd").Text = "zco9"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/usr/ctxtSE_KONNR-LOW").Text = ""
        .findById("wnd[0]/usr/btn%_SE_KONNR_%_APP_%-VALU_PUSH").press
    End With
    
    ' Copiar datos desde la hoja de Excel a SAP
    wsB.Range("A2:A" & lastRow).Copy
    session.findById("wnd[1]/tbar[0]/btn[24]").press ' Pegar en SAP
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Guardar archivo en SAP
    With session
        .findById("wnd[0]/mbar/menu[4]/menu[5]/menu[2]/menu[2]").Select
        .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
        .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
        .findById("wnd[1]/tbar[0]/btn[0]").press
        .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "zco9.XLS"
        .findById("wnd[1]/tbar[0]/btn[11]").press
    End With

    ' Cerrar SAP
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    session.findById("wnd[0]/tbar[0]/btn[15]").press

    ThisWorkbook.Sheets(6).Activate
    
    ' Solicitar al usuario que seleccione el archivo
    MsgBox "Por favor, seleccione el reporte zco9 recién descargado.", vbInformation, "Work&Service"

    On Error Resume Next
    rutaArchivo = ThisWorkbook.application.GetOpenFilename("Archivos de Excel (*.xls; *.xlsx ; *.xlsm), *.xls; *.xlsx; *.xlsm", , "Seleccione el archivo de Excel")
    On Error GoTo 0

    If rutaArchivo = "False" Or rutaArchivo = "" Then
        MsgBox "No se seleccionó ningún archivo.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Abrir el archivo seleccionado
    Set libro = Workbooks.Open(rutaArchivo)
    If libro Is Nothing Then
        MsgBox "No se pudo abrir el archivo seleccionado.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Copiar datos a la hoja "zco9"
    Set wsco9 = ThisWorkbook.Sheets("zco9")
    Set rangoOrigen = libro.Sheets(1).Range("A:AJ")

    rangoOrigen.Copy Destination:=wsco9.Range("A1")
    libro.Close SaveChanges:=False

    MsgBox "Reporte de Certificaciones cargado.", vbInformation, "Work&Service"
End Sub


