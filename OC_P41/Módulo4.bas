Attribute VB_Name = "M�dulo4"
'*************************************
'|||REALIZADO POR HERNAN F. CARRIZO|||
'************AGO-SEP 2024*************

Public application As Object
Public connection As Object
Public session As Object

'*************************************
'|||REALIZADO POR HERNAN F. CARRIZO|||
'************AGO-SEP 2024*************

Function IniciarSAP(usuario As String, contrasena As String) As Boolean
    Dim SapGuiAuto As Object
    Dim sapPath As String
    Dim sapIsOpen As Boolean
    Dim startTime As Double
    Dim timeout As Double
    Dim credenciales As Variant
    Dim loginSuccess As Boolean
    
    ' Inicializar retorno
    IniciarSAP = False
    
    ' Paso 1: Cargar las credenciales guardadas
    credenciales = CargarCredenciales()
    
    ' Verificar si las credenciales est�n vac�as
    If credenciales(0) = "" Or credenciales(1) = "" Then
        User = usuario
        Pass = contrasena
        Unload Inicio
    Else
        ' Usar las credenciales guardadas
        User = credenciales(0)
        Pass = credenciales(1)
        Unload Inicio
    End If
    
    ' Paso 2: Ruta del archivo ejecutable de SAP Logon
    sapPath = "C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe"

    ' Verificar si el archivo SAP Logon existe en la ruta especificada
    If Dir(sapPath) = "" Then
        MsgBox "No se encuentra el archivo SAP Logon en la ruta especificada."
        Exit Function
    End If

    ' Ejecutar el archivo de SAP Logon para abrir la aplicaci�n
    shell sapPath, vbNormalFocus

    ' Establecer un tiempo l�mite para esperar que SAP Logon se abra (por ejemplo, 30 segundos)
    timeout = 30
    startTime = Timer
    sapIsOpen = False

    ' Verificar continuamente si SAP GUI est� disponible
    Do
        On Error Resume Next
        Set SapGuiAuto = GetObject("SAPGUI")
        On Error GoTo 0
        If Not SapGuiAuto Is Nothing Then
            sapIsOpen = True
            Exit Do
        End If
        ' Salir si el tiempo de espera excede el l�mite
        If Timer - startTime > timeout Then
            MsgBox "SAP Logon tard� demasiado en abrirse."
            Exit Function
        End If
        DoEvents ' Permitir que el sistema procese otros eventos
    Loop
    
    ' Obtener el motor de scripting de SAP
    Set application = SapGuiAuto.GetScriptingEngine

    ' Verificar si la instancia de SAP GUI fue obtenida correctamente
    If application Is Nothing Then
        MsgBox "No se pudo obtener la instancia de SAP GUI. Aseg�rate de que SAP GUI est� abierto o que el scripting est� habilitado."
        Exit Function
    End If

    ' Paso 3: Intentar abrir una conexi�n si no hay ninguna conexi�n activa
    If application.Children.Count = 0 Then
    
        On Error Resume Next
            ' Abre una conexi�n al servidor SAP especificado (ajusta el nombre de la conexi�n)
        Set connection = application.OpenConnection("H172 C11 [SAP] - Producci�n Link", True)
        On Error GoTo 0
               
        ' Verificar si se pudo establecer la conexi�n
        If connection Is Nothing Then
            MsgBox "No se pudo abrir la conexi�n a SAP. Verifica el nombre del servidor o conexi�n."
            Exit Function
        End If
    Else
        ' Si ya hay conexiones activas, usar la primera
        Set connection = application.Children(0)
    End If
    
    ' Paso 4: Verificar si ya hay una sesi�n activa
    If connection.Children.Count > 0 Then
        Set session = connection.Children(0)

        ' Comprobar si estamos en la pantalla de login (campo de usuario)
        If Not session.findById("wnd[0]/usr/txtRSYST-BNAME", False) Is Nothing Then
            ' Estamos en la pantalla de login, proceder con el logueo
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = User
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Pass
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").SetFocus
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = Len(contrasena)
            session.findById("wnd[0]").SendVKey 0
            
             ' Verificar si el login fue exitoso
            On Error Resume Next
            loginSuccess = session.findById("wnd[0]/usr/txtRSYST-BNAME", False) Is Nothing
            On Error GoTo 0
            If Not loginSuccess Then
                MsgBox "Usuario o contrase�a inv�lidos. Verifica tus credenciales."
                Set connection = Nothing
                Set session = Nothing
                Exit Function
            End If
            
        End If

    Else
        MsgBox "No se encontr� ninguna sesi�n activa en la conexi�n."
        Exit Function
    End If
    
    ' Si todo sali� bien, devolver True
    IniciarSAP = True
End Function



