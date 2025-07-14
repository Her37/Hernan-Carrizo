Attribute VB_Name = "Módulo3"
'*************************************
'|||REALIZADO POR HERNAN F. CARRIZO|||
'************AGO-SEP 2024*************
Public numCont As String
Public application As Object
Public connection As Object
Public session As Object
Public nombrePagina As String
Public usuario As String, contrasena As String


Sub NuevaSolp()
    Dim ws As Worksheet
    Dim valorCampo As String
    Dim fecha_Hoy As String, textB As String, TextC As String, contrato As String
    Dim N_Servicio As Long, Tipo As String, NPEP As Variant, CCoste As String, zm_50 As Variant
    Dim Monto As String, Fecha1 As String, Fecha2 As String, Fecha As String
    Dim lastRow As Long
    Dim i As Integer
    Dim j As Integer
    Dim mensajeError As String
    Dim grupoArticulo As String
    Dim opcionSeleccionada As String
    
    Set ws = ThisWorkbook.Sheets("NewSolp")


    ' No hay sesiones activas, mostrar el formulario para iniciar sesión
    Inicio.Show vbModal
    
    ' Comprobar si el formulario se cerró correctamente
    If Inicio.Tag = "OK" Then
        usuario = Inicio.txtUsuario.text
        contrasena = Inicio.txtContraseña.text
        ' Iniciar SAP después de cerrar el formulario
        If IniciarSAP(usuario, contrasena) Then
        Else
            Exit Sub
        End If
    Else
        ' El formulario se cerró sin guardar las credenciales
        Exit Sub
    End If
    
    If Panel.TextBox1.Value = "" Then
        MsgBox "Ingrese un contrato"
        Exit Sub
    End If
    
            
    If Panel.TextBox35.Value = "" Or Panel.TextBox33.Value = "" Or Panel.TextBox34.Value = "" Then
        Call ZM50
        If Panel.TextBox31.Value = "" And Panel.TextBox42.Value = "" Then
            Call ZCO9
        End If
    End If

    TextV = Panel.TextBox40.Value
    'TextV = ws.Range("B2").Value
    TextC = Panel.TextBox39.Value
    'TextC = ws.Range("B3").Value
    Monto = Trim(Panel.TextBox38.Value)
    'Monto = ws.Range("F4").Value
    'contrato = ws.Range("F9").Value
    contrato = Trim(Panel.TextBox1.Value)
    contrato2 = Trim(Panel.TextBox41.Value)
    
   
    Fecha = ws.Range("C5").Value
    fecha_Hoy = Replace(Fecha, "/", ".")
    
    Fecha = Panel.TextBox36.Value
    Fecha1 = Replace(Fecha, "/", ".")
    'Fecha1 = ws.Range("F5").Value
    Fecha = Panel.TextBox37.Value
    Fecha2 = Replace(Fecha, "/", ".")
    'Fecha2 = ws.Range("F6").Value
    
    grupoArticulo = ws.Range("D6").Value
    GM = Trim(Panel.TextBox35.Value)
    Moneda = Panel.TextBox33.Value
    centro = Panel.TextBox34.Value
    NPEP = Trim(Panel.TextBox31.Value)
    CCoste = Trim(Panel.TextBox32.Value)
    Proyecto = Trim(Panel.TextBox42.Value)
    'N_Servicio = Trim(ws.Range("C7").Value)
    
    
Select Case True
    Case Panel.Transferencias.Value
        opcionSeleccionada = ws.Range("C2").Value
    Case Panel.Vigencia.Value
        opcionSeleccionada = ws.Range("C2").Value
        If Fecha2 = "" Then
            MsgBox "Ingrese Fecha"
            Exit Sub
        End If
    Case Panel.Monto.Value
        opcionSeleccionada = ws.Range("C2").Value
        If Monto = "" Then
            MsgBox "Ingrese codigo de Imputación"
            Exit Sub
        End If
    Case Panel.Licitación.Value
        opcionSeleccionada = ws.Range("C2").Value
        If NPEP = "" Or Proyecto = "" Then
            MsgBox "Ingrese codigo de Imputación"
            Exit Sub
        End If
    Case Else
        MsgBox "Debe seleccionar una opción.", vbExclamation, "Aviso"
        Exit Sub
End Select
    
  
 
    On Error Resume Next
    Call VolverAVentanaPrincipalSAP
    On Error GoTo 0

    session.findById("wnd[0]/tbar[0]/okcd").text = "me51n"
    session.findById("wnd[0]").sendVKey 0

    On Error Resume Next
    ' Inserta texto en el editor de texto
    session.findById( _
        "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/" & _
        "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/" & _
        "tabpTABREQHDT1/ssubTABSTRIPCONTROL3SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/" & _
        "subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").text = TextC
 
    ' Comprobar si hubo un error
    If Err.Number <> 0 Then
        Call PresionarBotonDinamico
        ' Inserta texto en el editor de texto
        session.findById( _
        "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/" & _
        "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/" & _
        "tabpTABREQHDT1/ssubTABSTRIPCONTROL3SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/" & _
        "subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").text = TextC
        
        Err.Clear
    End If
    
    On Error GoTo 0 ' Restaurar manejo normal de errores

    ' Presiona botón
                
    Dim boton As Object
    
    Set boton = session.findById( _
        "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4001" & _
        "/btnDYN_4000-BUTTON")
    
    If boton.IconName = "DAAREX" Then
        boton.press
        Debug.Print "el boton se presionó"
    End If

    ' Modifica múltiples celdas en la cuadrícula
    Dim grid As Object
    Set grid = session.findById( _
        "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/" & _
        "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell")

    grid.modifyCell 0, "BNFPO", "10"
    grid.modifyCell 0, "KNTTP", "U"
    grid.modifyCell 0, "EPSTP", "D"
    grid.modifyCell 0, "TXZ01", TextV
    grid.modifyCell 0, "ELPEI", "D"
    grid.modifyCell 0, "EEIND", fecha_Hoy
    'grid.modifyCell 0, "WGBEZ", Trim(GM) & Trim(grupoArticulo)
    'MsgBox "El Centro propuesto corresponde al usuario?"
    grid.modifyCell 0, "NAME1", centro
    grid.modifyCell 0, "EKGRP", "700"
    grid.modifyCell 0, "BEDNR", "NEED"
    grid.modifyCell 0, "EKORG", "EDES"
    grid.setCurrentCell 1, "BNFPO"
    grid.pressEnter

Do
    grid.modifyCell 0, "WGBEZ", Trim(GM) & Trim(grupoArticulo)
    grid.pressEnter
    ' Verificamos si aparece un mensaje de error en la barra de estado
    mensajeError = session.findById("wnd[0]/sbar").text

    If InStr(mensajeError, "Indique grupo de artículos") > 0 Then
        ' Si aparece el mensaje de error, pedimos al usuario que ingrese un nuevo valor
        grupoArticulo = InputBox("por favor ingrese dos nuevos núemeros para el GM (01,02,03,etc)", "Grupo de Artículos")
        
        ' Si el usuario cancela el InputBox, salimos del bucle
        If grupoArticulo = "" Then
            MsgBox "Proceso cancelado."
            Exit Sub
        End If
    Else
        ' Si no hay error, continuamos con el proceso
        Exit Do
    End If
Loop
 
 On Error GoTo ErrorHandler ' Manejo de errores

Select Case opcionSeleccionada
        Case "Licitación"
        maxFilaVisible = 10 ' Ajusta según cuántas filas son visibles en SAP
        filaSAP = 0 ' Iniciar en la primera fila visible
        j = 1 ' Iniciar desde el primer elemento en el ListBox

        ' Recorrer los elementos deL ListBox
        For i = 0 To Panel.ListBox2.ListCount - 2
            Do
                On Error Resume Next ' Ignorar errores en SAP
                errFlag = False
                
                ' Intentar ingresar el valor en SAP desde el ListBox
                session.findById( _
                    "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                    "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/" & _
                    "tabpTABREQDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/" & _
                    "tblSAPLMLSPTC_VIEW/ctxtESLL-SRVPOS[2," & filaSAP & "]").text = Panel.ListBox2.List(j)
                
                ' Comprobar si hubo un error
                If Err.Number <> 0 Then
                    errFlag = True
                    Err.Clear
                End If
                
                On Error GoTo 0 ' Restaurar manejo normal de errores

                ' Si hay un error, mover el scroll y volver a intentar
                If errFlag Then
                    ' Mover el scroll hacia abajo
                        session.findById( _
                        "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                        "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/" & _
                        "tabpTABREQDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/" & _
                        "tblSAPLMLSPTC_VIEW").verticalScrollbar.Position = maxFilaVisible
                    
                    filaSAP = 1 ' Reiniciar el contador de fila visible
                Else
                    filaSAP = filaSAP + 1 ' Avanzar a la siguiente fila visible en SAP
                    j = j + 1 ' Avanzar al siguiente elemento en el ListBox
                End If
                
            Loop While errFlag ' Repetir hasta que no haya error
            
        Next i

            
        Case Else
            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                "tabsREQ_ITEM_DETAIL/tabpTABREQDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/" & _
                "subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/ctxtESLL-SRVPOS[2,0]").text = Panel.ListBox2.List(0)
                On Error GoTo 0
    End Select

    ' Envío de tecla y pestañas adicionales
    session.findById("wnd[0]").sendVKey 0
    session.findById( _
        "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/" & _
        "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
        "tabsREQ_ITEM_DETAIL/tabpTABREQDT15").Select

' Procesa el caso de Tipo
Select Case opcionSeleccionada
    Case "Transferencias"
            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/" & _
                "ssubSUB2:SAPLXM02:9000/radEBAN-ZZTRAN").Select
            session.findById("wnd[0]").sendVKey 0
                       
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/" & _
                         "ssubSUB2:SAPLXM02:9000/ctxtEBAN-ZCONTRAO").text = contrato
        
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/" & _
                         "ssubSUB2:SAPLXM02:9000/ctxtEBAN-ZCONTRAD").text = contrato2
        
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/" & _
                         "ssubSUB2:SAPLXM02:9000/txtEBAN-ZMONTO1").text = Monto
                   
        session.findById("wnd[0]").sendVKey 0
            
Case "Licitación"
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                    "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                    "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                    "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/" & _
                    "ssubSUB2:SAPLXM02:9000/radEBAN-ZZNUEVO").Select
                session.findById("wnd[0]").sendVKey 0
                
                      ' Completa los campos Fechas
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1").Select
                         
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEINI").text = Fecha1
                         
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEFIN").text = Fecha2
        
        ' Completa los campos de Presupuesto
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5").Select
                         
        If Not NPEP = "" Then
            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/" & _
                "ctxtT_ZEBAN_GRILLA-POSID[2,0]").text = NPEP
        Else
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/ctxtT_ZEBAN_GRILLA-PROY[1,0]").text = Proyecto
        End If
        
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/" & _
            "ctxtT_ZEBAN_GRILLA-CLVCO[5,0]").text = CCoste
            
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
             "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
             "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
             "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/" & _
             "tblSAPLXM02C_ZEBAN/txtT_ZEBAN_GRILLA-MOSOLI[7,0]").text = Monto
             
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
            "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH4/ssubSUB4:SAPLXM02:9004/ctxtEBAN-ZZMON").text = Moneda
                
        session.findById("wnd[0]").sendVKey 0
        
 Case "Vigencia"
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/" & _
                "ssubSUB2:SAPLXM02:9000/radEBAN-ZZAMPLA").Select
            session.findById("wnd[0]").sendVKey 0
            
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
            "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/ssubSUB2:SAPLXM02:9000/ctxtEBAN-ZZNUMB").SetFocus
            
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/" & _
            "tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/" & _
            "ssubSUB2:SAPLXM02:9000/ctxtEBAN-ZZNUMB").text = Trim(contrato)
            session.findById("wnd[0]").sendVKey 0
            
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
            "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/ssubSUB2:SAPLXM02:9000/radEBAN-ZZMOFEC").Select
            
            ' Completa los campos Fechas
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1").Select
        session.findById("wnd[0]").sendVKey 0
        
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEINI").text = Fecha1
                         
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEFIN").text = Fecha2
            
Case "Monto", "MMC"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
        "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/" & _
        "tabsGRILLA/tabpPUSH2/ssubSUB2:SAPLXM02:9000/radEBAN-ZZAMPLA").Select
        session.findById("wnd[0]").sendVKey 0
        
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
        "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/" & _
        "tabsGRILLA/tabpPUSH2/ssubSUB2:SAPLXM02:9000/ctxtEBAN-ZZNUMB").text = contrato
        session.findById("wnd[0]").sendVKey 0
        
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
        "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/" & _
        "tabsGRILLA/tabpPUSH2/ssubSUB2:SAPLXM02:9000/radEBAN-ZZMOPRE").Select
        session.findById("wnd[0]").sendVKey 0
        
        If opcionSeleccionada = "MMC" Then
        
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
        "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/" & _
        "tabsGRILLA/tabpPUSH2/ssubSUB2:SAPLXM02:9000/radEBAN-ZZAJMMC").Select
        
        End If
        
        session.findById("wnd[0]").sendVKey 0
        

    session.findById("wnd[1]").Close
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]").Close
        
                              ' Completa los campos Fechas
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1").Select
                         
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEINI").text = Fecha1
                         
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEFIN").text = Fecha2
                         
        
        ' Completa los campos de Presupuesto
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5").Select
                         
        If Not NPEP = "" Then
            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/" & _
                "ctxtT_ZEBAN_GRILLA-POSID[2,0]").text = NPEP
        Else
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/ctxtT_ZEBAN_GRILLA-PROY[1,0]").text = Proyecto
        End If
        
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/" & _
            "ctxtT_ZEBAN_GRILLA-CLVCO[5,0]").text = CCoste
            
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
             "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
             "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
             "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/" & _
             "tblSAPLXM02C_ZEBAN/txtT_ZEBAN_GRILLA-MOSOLI[7,0]").text = Monto
             
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
            "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH4/ssubSUB4:SAPLXM02:9004/ctxtEBAN-ZZMON").text = Moneda
                
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[1]").Close
        
Case "Solp Complementaria"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/" & _
            "ssubSUB2:SAPLXM02:9000/radEBAN-ZZNUEVO").Select
        session.findById("wnd[0]").sendVKey 0
                
                      ' Completa los campos Fechas
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1").Select
                         
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEINI").text = Fecha1
                         
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEFIN").text = Fecha2
        
        ' Completa los campos de Presupuesto
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5").Select
                         
        If Not NPEP = "" Then
            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/" & _
                "ctxtT_ZEBAN_GRILLA-POSID[2,0]").text = NPEP
        Else
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/ctxtT_ZEBAN_GRILLA-PROY[1,0]").text = Proyecto
        End If
        
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/" & _
            "ctxtT_ZEBAN_GRILLA-CLVCO[5,0]").text = CCoste
            
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
             "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
             "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
             "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/" & _
             "tblSAPLXM02C_ZEBAN/txtT_ZEBAN_GRILLA-MOSOLI[7,0]").text = Monto
             
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
            "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH4/ssubSUB4:SAPLXM02:9004/ctxtEBAN-ZZMON").text = Moneda
                
        session.findById("wnd[0]").sendVKey 0
        
Case Else
    MsgBox "Opción no válida"
End Select


session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15").Select
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1").Select
session.findById("wnd[0]").sendVKey 0



Dim text As Variant

text = session.findById("wnd[0]/sbar").text
If Not text = "" Then
    MsgBox "ERROR: " & text, vbCritical
    Exit Sub
Else
    Dim eleccion As VbMsgBoxResult
    Dim solp As Variant
        eleccion = MsgBox("Confirma grabar esta Solp?", vbYesNo, "SOLP")
        
    If eleccion = vbNo Then
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        Exit Sub ' Si elige Cancelar, salir del procedimiento
            
    ElseIf eleccion = vbYes Then
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        solp = session.findById("wnd[0]/sbar").text 'COPIAR NUMERO DE SOLP
        'Panel.Label47.Name = solp
        ws.Range("F10").Value = solp
        Panel.Label47.Caption = ws.Range("F10").Value
        MsgBox "N° solp: " & ws.Range("F10").Value

    End If
End If

    ' Libera los objetos
    Set session = Nothing
    Set connection = Nothing
    Set application = Nothing
    Set SapGuiAuto = Nothing
    
' Manejo de errores
ErrorHandler:
    'MsgBox "Error: " & Err.Description, vbCritical, "Error"
    Err.Clear
End Sub


