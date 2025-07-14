Attribute VB_Name = "Módulo3"
'*************************************
'|||REALIZADO POR HERNAN F. CARRIZO|||
'************AGO-SEP 2024*************

Sub NuevaSolp()
    Dim ws As Worksheet
    Dim usuario As String, contrasena As String, valorCampo As String
    Dim GM As String, fecha_Hoy As String, textB As String, TextC As String
    Dim N_Servicio As Long, Tipo As String, NPEP As Variant, CCoste As String, centro As String
    Dim Monto As Long, Moneda As String, Fecha1 As String, Fecha2 As String, Fecha As String, contrato As String
    Dim lastRow As Long
    Dim i As Integer
    Dim j As Integer
    Dim mensajeError As String
    Dim grupoArticulo As String

    
    Set ws = ThisWorkbook.Sheets("NewSolp")
    
        
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
            
    
    Tipo = ws.Range("B2").value
    GM = ws.Range("B6").value
    TextV = ws.Range("B4").value
    TextC = ws.Range("B3").value
    CCoste = ws.Range("F3").value
    Monto = ws.Range("F4").value
    Moneda = ws.Range("F8").value
    contrato = ws.Range("F9").value
    grupoArticulo = ws.Range("C6").value
    centro = ws.Range("F12").value
    
    Fecha = ws.Range("B5").value
    fecha_Hoy = Replace(Fecha, "/", ".")
    
    Fecha = ws.Range("F5").value
    Fecha1 = Replace(Fecha, "/", ".")
    
    Fecha = ws.Range("F6").value
    Fecha2 = Replace(Fecha, "/", ".")

    NPEP = ZCO9()
    ws.Range("F3").value = NPEP(0)
    ws.Range("F2").value = NPEP(1)


    Call VolverAVentanaPrincipalSAP
    
    ' Maximiza la ventana principal de SAP
    'session.findById("wnd[0]").maximize

    ' Selecciona y hace doble clic en el nodo "F00021"
    Dim shell As Object
    Set shell = session.findById( _
        "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell")
    shell.selectedNode = "F00021"
    shell.doubleClickNode "F00021"

    ' Inserta texto en el editor de texto
    session.findById( _
        "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/" & _
        "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/" & _
        "tabpTABREQHDT1/ssubTABSTRIPCONTROL3SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/" & _
        "subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").Text = TextC

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
    grid.modifyCell 0, "WGBEZ", Trim(GM) & Trim(grupoArticulo)
    grid.modifyCell 0, "NAME1", "00" & centro
    grid.modifyCell 0, "EKGRP", "700"
    grid.modifyCell 0, "BEDNR", "NEED"
    grid.modifyCell 0, "EKORG", "EDES"
    grid.setCurrentCell 1, "BNFPO"
    grid.pressEnter

Do
    grid.modifyCell 0, "WGBEZ", Trim(GM) & Trim(grupoArticulo)
    grid.pressEnter
    ' Verificamos si aparece un mensaje de error en la barra de estado
    mensajeError = session.findById("wnd[0]/sbar").Text

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

    ' Carga de servicios para licitaciones
Select Case Tipo
        Case "Licitación"
            maxFilaVisible = 10 ' Ajusta según cuántas filas son visibles en SAP
            filaSAP = 0 ' Iniciar en la primera fila visible
            j = 0 ' Iniciar desde la primera fila en la hoja de Excel

            For i = 0 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
                Do
                    On Error Resume Next ' Ignorar errores en SAP
                    errFlag = False
                    
                    ' Intentar ingresar el valor en SAP
                    session.findById( _
                        "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                        "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/" & _
                        "tabpTABREQDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/" & _
                        "tblSAPLMLSPTC_VIEW/ctxtESLL-SRVPOS[2," & filaSAP & "]").Text = ws.Range("B" & j + 7).value
                    
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
                            "tblSAPLMLSPTC_VIEW").verticalScrollbar.Position = _
                            session.findById( _
                            "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
                            "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/" & _
                            "tabpTABREQDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/" & _
                            "tblSAPLMLSPTC_VIEW").verticalScrollbar.Position + maxFilaVisible
                        
                        filaSAP = 1 ' Reiniciar el contador de fila visible
                    Else
                        filaSAP = filaSAP + 1 ' Avanzar a la siguiente fila visible en SAP
                        j = j + 1 ' Avanzar en la hoja de Excel
                    End If
                    
                Loop While errFlag ' Repetir hasta que no haya error
                
            Next i
            
        Case Else
            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                "tabsREQ_ITEM_DETAIL/tabpTABREQDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/" & _
                "subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/ctxtESLL-SRVPOS[2,0]").Text = 28901
                On Error GoTo 0
    End Select

    ' Envío de tecla y pestañas adicionales
    session.findById("wnd[0]").sendVKey 0
    session.findById( _
        "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/" & _
        "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
        "tabsREQ_ITEM_DETAIL/tabpTABREQDT15").Select

   ' Procesa el caso de Tipo
    Select Case Tipo
        Case "Transferencia Montos"
            session.findById( _
                "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH2/" & _
                "ssubSUB2:SAPLXM02:9000/radEBAN-ZZTRAN").Select
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
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEINI").Text = Fecha1
                         
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEFIN").Text = Fecha2
        
        ' Completa los campos de Presupuesto
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5").Select
                         
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/" & _
            "ctxtT_ZEBAN_GRILLA-POSID[2,0]").Text = NPEP(1)
        
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/" & _
            "ctxtT_ZEBAN_GRILLA-CLVCO[5,0]").Text = NPEP(0)
            
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
             "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
             "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
             "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/" & _
             "tblSAPLXM02C_ZEBAN/txtT_ZEBAN_GRILLA-MOSOLI[7,0]").Text = Monto
             
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
            "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH4/ssubSUB4:SAPLXM02:9004/ctxtEBAN-ZZMON").Text = "ARP"
                
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
            "ssubSUB2:SAPLXM02:9000/ctxtEBAN-ZZNUMB").Text = Trim(contrato)
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
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEINI").Text = Fecha1
                         
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEFIN").Text = Fecha2
            
        Case "MMC"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
        "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/" & _
        "tabsGRILLA/tabpPUSH2/ssubSUB2:SAPLXM02:9000/radEBAN-ZZAMPLA").Select
        session.findById("wnd[0]").sendVKey 0
        
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
        "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/" & _
        "tabsGRILLA/tabpPUSH2/ssubSUB2:SAPLXM02:9000/ctxtEBAN-ZZNUMB").Text = contrato
        session.findById("wnd[0]").sendVKey 0
        
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
        "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/" & _
        "tabsGRILLA/tabpPUSH2/ssubSUB2:SAPLXM02:9000/radEBAN-ZZMOPRE").Select
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
        "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/" & _
        "tabsGRILLA/tabpPUSH2/ssubSUB2:SAPLXM02:9000/radEBAN-ZZAJMMC").Select
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
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEINI").Text = Fecha1
                         
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH1/ssubSUB1:SAPLXM02:9001/ctxtEBAN-ZZFEFIN").Text = Fecha2
                         
        
                ' Completa los campos de Presupuesto
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
                         "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
                         "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5").Select
        session.findById("wnd[0]").sendVKey 0
        

session.findById("wnd[1]").Close
'session.findById("wnd[1]").Close

        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/" & _
            "ctxtT_ZEBAN_GRILLA-POSID[2,0]").Text = NPEP(1)
            
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
            "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
            "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/tblSAPLXM02C_ZEBAN/" & _
            "ctxtT_ZEBAN_GRILLA-CLVCO[5,0]").Text = NPEP(0)
            
        session.findById( _
            "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" & _
             "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/" & _
             "tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/" & _
             "ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsGRILLA/tabpPUSH5/ssubSUB5:SAPLXM02:9005/" & _
             "tblSAPLXM02C_ZEBAN/txtT_ZEBAN_GRILLA-MOSOLI[7,0]").Text = Monto
        session.findById("wnd[0]").sendVKey 0
        
        session.findById("wnd[1]").Close
        
        Case Else
            MsgBox "Opción no válida"
End Select
        
Dim eleccion As VbMsgBoxResult
Dim solp As Long
        eleccion = MsgBox("Confirma grabar esta Solp?", vbYesNo, "SOLP")
        
    If eleccion = vbNo Then
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        Exit Sub ' Si elige Cancelar, salir del procedimiento
            
    ElseIf eleccion = vbYes Then
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        ws.Range("F11").value = session.findById("wnd[0]/sbar").Text 'COPIAR NUMERO DE SOLP
        'solp = ws.Range("F11").value
        MsgBox "N° solp: " & ws.Range("F11").value
    End If
        
    ' Libera los objetos
    Set session = Nothing
    Set connection = Nothing
    Set application = Nothing
    Set sapGuiAuto = Nothing

End Sub

