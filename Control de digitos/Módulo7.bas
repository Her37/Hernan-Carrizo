Attribute VB_Name = "Módulo7"
'*************************************
'|||REALIZADO POR HERNAN F. CARRIZO|||
Option Explicit
Public cancelarProceso As Boolean
Public ws As Worksheet
Public sessionNuevoModo As Object
Public duracionActual As Double

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function FormatearTiempo(duracionHoras As Double) As String
    Dim horas As Double
    Dim minutos As Double
    Dim segundos As Double

    ' Convertir las horas decimales en horas y minutos
    horas = Int(duracionHoras / 3600)
    minutos = Int((duracionHoras Mod 3600) / 60) ' Resto en minutos
    segundos = duracionHoras Mod 60

    ' Formato de tiempo hh:mm
    FormatearTiempo = Format(horas, "00") & ":" & Format(minutos, "00") & ":" & Format(segundos, "00")
End Function

Sub ZCO9_certificados()
    Dim usuario As String, contrasena As String
    Dim i As Integer
    Dim j As Integer
    Dim labelText As String
    Dim moreRows As Boolean
    Dim maxLabels As Integer
    Dim lastRow As Long
    Dim scrollPos As Integer
    Dim MaxScroll As Double
    Dim prevScrollPos As Integer
    Dim hojanombre As String
    Dim eleccion As VbMsgBoxResult
    Dim hojaExistente As Boolean
    Dim nuevaConsulta As Boolean
    Dim nuevaHoja As Worksheet
    Set ws = Nothing
    Dim reconfirmacion As VbMsgBoxResult
    Dim logSheet As Worksheet
    Dim nextLogRow As Long
    Dim Nfactura As String
    Dim contratoNumero As String
    Dim labelTextAnterior As String
    Dim duracionMax As Double
    Dim InicioOK As Boolean
    
    ManejarVentanaWindows
    
    Inicio.Show vbModal
       
    If Inicio.Tag = "Exit" Then
        Exit Sub
        Unload Panel
    End If
    
    InicioOK = IniciarSAP(usuario, contrasena)

    If Not InicioOK Then
        MsgBox "No se pudo iniciar SAP.", vbCritical
        Exit Sub
    End If
          
   ' Obtener el nombre del campo CombotBox1
    hojanombre = Panel.ComboBox1.value '
    
        ' Asegúrate de que el ComboBox tiene un valor seleccionado
    If hojanombre <> "" And Len(hojanombre) <= 10 Then
        ' Asigna el valor seleccionado a la variable numCont
        contratoNumero = Trim(Panel.ComboBox1.value)
    Else
        MsgBox "Ingrese un número de Contrato valido.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar si la hoja con el nombre ya existe
    hojaExistente = False
    On Error Resume Next ' Evitar errores si la hoja no existe
    Set ws = ThisWorkbook.Sheets(hojanombre)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        hojaExistente = True
    End If
      
    ' Si existe la hoja y elige que sí (sobrescribir), se pregunta antes de sobrescribir
    If hojaExistente Then
        eleccion = MsgBox("El contrato '" & hojanombre & "' se actualizará desde el ultimo Certificado cargado. ¿Desea continuar?", vbYesNo, "Contrato existente")
        
        If eleccion = vbNo Then
            Exit Sub ' Si elige Cancelar, salir del procedimiento
        End If
    Else
            MsgBox "Se cargara un nuevo contrato a la base de datos."
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            ws.Name = hojanombre ' Crear la nueva hoja con el nombre del campo TextBox1
            ws.Cells(1, 1).value = "CERTIFICADO"
            ws.Cells(1, 2).value = "HOJA DE ENTRADA"
            ws.Cells(1, 3).value = "POSICION"
            ws.Cells(1, 4).value = "SERVICIO"
            ws.Cells(1, 5).value = "TEXTO"
            ws.Cells(1, 6).value = "VALOR"
            ws.Columns("A:F").AutoFit
            ws.Range("C3:C100000").Formula = "=IF(E3="""","""",C2)"
    End If
    
    ' Desactivar la actualización de pantalla y eventos para mejorar el rendimiento
    ThisWorkbook.application.ScreenUpdating = False
    ThisWorkbook.application.Calculation = xlCalculationManual
    ThisWorkbook.application.EnableEvents = False
    
    Panel.Hide
    
    Call IniciarSAP(usuario, contrasena)
    
    Call VolverAVentanaPrincipalSAP
    
    If connection Is Nothing Then
        Exit Sub
    End If

    Temporizador.Show vbModeless
    Temporizador.Duracion.Caption = "00:00:00"
    Temporizador.Caption = "Cargando Contrato: " & Panel.ComboBox1.value
    'duracionActual = 0
    
    On Error Resume Next
        session.findById("wnd[0]/tbar[0]/okcd").Text = "zco9"
        session.findById("wnd[0]").sendVKey 0
    On Error GoTo 0
    
    If Err.Number <> 0 Then
        Set connection = Nothing
        Set session = Nothing
        Unload Temporizador
        Err.Clear
        Exit Sub ' Salir del bucle si ocurre un error
    End If
  

    session.findById("wnd[0]/usr/ctxtSE_KONNR-LOW").Text = contratoNumero
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    Dim ventanaEmergente As Object
    ' Verifica si existe la ventana emergente
    On Error Resume Next
    Set ventanaEmergente = session.findById("wnd[1]") ' Ventana emergente usualmente es "wnd[1]"
    On Error GoTo 0

    If Not ventanaEmergente Is Nothing Then
        ' Verifica el contenido del mensaje en la ventana emergente
        Dim mensaje As String
        mensaje = ventanaEmergente.findById("usr/txtMESSTXT1").Text

        If mensaje = "No existen certificados para la seleccion" Then
            ' Toma una acción
            ventanaEmergente.findById("tbar[0]/btn[0]").press ' Botón "OK" o "Aceptar"
            MsgBox "No existen certificados para el Contrato seleccionado."
            ThisWorkbook.application.DisplayAlerts = False ' Desactiva las alertas de Excel
            ThisWorkbook.Sheets(hojanombre).Delete ' Elimina la hoja sin mostrar la confirmación
            ThisWorkbook.application.DisplayAlerts = True ' Reactiva las alertas de Excel
            Unload Temporizador
            Exit Sub
        End If
    End If

  ' Inicializar valores
    maxLabels = 29 ' Cantidad máxima de etiquetas visibles por pantalla
    moreRows = True ' Bandera para saber si hay más filas
    scrollPos = 0 ' Posición inicial de la barra de desplazamiento
    MaxScroll = session.findById("wnd[0]/usr").verticalScrollbar.Maximum ' Máximo valor de la barra de scroll
    cancelarProceso = False
    
    If MaxScroll = 0 Then
        duracionMax = (maxLabels * 30 * 4) 'duracion maxima aproximada en segundos
    Else
        duracionMax = (MaxScroll * 30 * 4) 'duracion maxima aproximada en segundos
    End If
    
    duracionActual = duracionMax
    Temporizador.Duracion.Caption = FormatearTiempo(duracionActual)
    
    ' Bucle hasta que no haya más filas o la barra de desplazamiento esté al máximo
    Do While moreRows And scrollPos < MaxScroll
        ' Bucle para capturar todas las etiquetas visibles en la pantalla actual
        For i = 0 To maxLabels
            On Error Resume Next ' Ignorar errores si no encuentra más etiquetas
                lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
                labelText = session.findById("wnd[0]/usr/lbl[9," & 4 + i & "]").Text
                Nfactura = Trim(session.findById("wnd[0]/usr/lbl[52," & 4 + i & "]").Text)
            
                        ' Comparar el labelText actual con el anterior
            If labelText = labelTextAnterior Then
                Exit For
            End If
  
            If WorksheetFunction.CountIf(ws.Columns("A"), labelText) = 0 Or ws.Cells(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, "A").value = Val(labelText) Then
                If Nfactura <> "0" Then
                    session.findById("wnd[0]/usr/lbl[9," & 4 + i & "]").SetFocus
                    session.findById("wnd[0]").sendVKey 2
                    
                    If ws.Cells(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, "A").value = Val(labelText) Then
                        ws.Cells(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, 1).value = labelText ' Escribir en la hoja
                    Else
                        ws.Cells(lastRow + 1, 1).value = labelText ' Escribir en la hoja
                    End If
            
                    Call hojaEntrada
            
                End If
            End If
            
            ' Asignar el valor de labelText a labelTextAnterior para la siguiente iteración
            labelTextAnterior = labelText
            On Error GoTo 0 ' Restaurar manejo normal de errores

        Next i
    On Error Resume Next ' Ignorar errores si no encuentra más etiquetas
        ' Guardar posición actual de scroll
        prevScrollPos = scrollPos
        ' Mover la barra de desplazamiento hacia abajo
        session.findById("wnd[0]/usr").verticalScrollbar.Position = scrollPos + maxLabels
        scrollPos = session.findById("wnd[0]/usr").verticalScrollbar.Position
    On Error GoTo 0 ' Restaurar manejo normal de errores
    
       ' Calcular tiempo restante en función de la posición actual
       If MaxScroll = 0 Then
            duracionActual = duracionActual - (i * 30 * 4)
            
       Else
            duracionActual = duracionActual - (maxLabels * 30 * 4)
       End If
       
       If duracionActual < 0 Then duracionActual = duracionActual + (maxLabels * 30 * 4)  ' Asegurar que no sea negativo
       Temporizador.Duracion.Caption = FormatearTiempo(duracionActual)
    
        ' Si no cambia la posición de scroll, no hay más filas
        If prevScrollPos = scrollPos Then
            moreRows = False
        End If
    Loop
   
For i = 0 To maxLabels
    On Error Resume Next ' Ignorar errores si no encuentra más etiquetas
        labelTextAnterior = ""
        lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
        labelText = session.findById("wnd[0]/usr/lbl[9," & 4 + i & "]").Text
        Nfactura = Trim(session.findById("wnd[0]/usr/lbl[52," & 4 + i & "]").Text)
        'Debug.Print Err.Number
        
        If Err.Number <> 0 Then
            Exit For
        End If
    On Error GoTo 0 ' Restaurar manejo normal de errores
    
    ' Comparar el labelText actual con el anterior
    If labelText = labelTextAnterior Then
        Exit For
    End If
    
    If WorksheetFunction.CountIf(ws.Columns("A"), labelText) = 0 Or ws.Cells(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, "A").value = Val(labelText) Then
        If Nfactura <> "0" Then
            session.findById("wnd[0]/usr/lbl[9," & 4 + i & "]").SetFocus
            session.findById("wnd[0]").sendVKey 2
            
            If ws.Cells(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, "A").value = Val(labelText) Then
                ws.Cells(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, 1).value = labelText ' Escribir en la hoja
            Else
                ws.Cells(lastRow + 1, 1).value = labelText ' Escribir en la hoja
            End If
            
            Call hojaEntrada
            If cancelarProceso = True Then
                Exit For
            End If
           
        End If
    End If
    
    ' Asignar el valor de labelText a labelTextAnterior para la siguiente iteración
    labelTextAnterior = labelText

    
           ' Calcular tiempo restante en función de la posición actual
       'duracionActual = duracionActual - (30 * 4)
       If duracionActual < 0 Then duracionActual = 0 'Asegurar que no sea negativo
       Temporizador.Duracion.Caption = FormatearTiempo(duracionActual)
    
Next i

    If Panel.CheckBox2.value = True Then
        Call LimpiarBloqueados(hojanombre)
        Call VolverAVentanaPrincipalSAP
    End If
  
    If cancelarProceso = True Then
        Temporizador.Hide
        Panel.Show vbModeless
        'session.findById("wnd[0]/tbar[0]/btn[3]").press
        'session.findById("wnd[0]/tbar[0]/btn[15]").press
        MsgBox "Proceso pausado.", vbInformation
        On Error Resume Next
       ' Intentar cerrar todas las sesiones de SAP utilizadas
        sessionNuevoModo.findById("wnd[0]/tbar[0]/btn[3]").press
        sessionNuevoModo.findById("wnd[0]/tbar[0]/btn[15]").press
        sessionNuevoModo.findById("wnd[0]/tbar[0]/btn[15]").press
        sessionNuevoModo.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        On Error GoTo 0
        ' Guardar el libro y mostrar mensaje al finalizar
        'ThisWorkbook.Save
        'connection.CloseSession "ses[0]"
        Set session = Nothing
        Set connection = Nothing
        Set application = Nothing
        Set sessionNuevoModo = Nothing
        Exit Sub
        
        
    ElseIf session Is Nothing Or session.Info.SystemName = "" Then
        MsgBox "Ocurrio un error. Carga interrumpida. La conexión con SAP se perdió.", vbExclamation
        Debug.Print Err.Number
        Set session = Nothing
        Set connection = Nothing
        Set application = Nothing
        Set sessionNuevoModo = Nothing
        ' Activar y traer Excel al frente
        ThisWorkbook.Activate
        ThisWorkbook.Windows(1).Visible = True
        Temporizador.Duracion.Caption = "00:00"
        Unload Temporizador
        Panel.Show vbModeless
            With ThisWorkbook.application
                .ScreenUpdating = True
                .Calculation = xlCalculationAutomatic
                .EnableEvents = True
            End With
        Exit Sub
    Else
        On Error Resume Next
       ' Intentar cerrar todas las sesiones de SAP utilizadas
        sessionNuevoModo.findById("wnd[0]/tbar[0]/btn[3]").press
        sessionNuevoModo.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        connection.CloseSession "ses[0]"
        On Error GoTo 0
        
          ' Guardar el número de contrato y la fecha de finalización exitosa
        On Error Resume Next
        Set logSheet = ThisWorkbook.Sheets("Registro")
        On Error GoTo 0
        
        ' Si la hoja de registro no existe, crearla
        If logSheet Is Nothing Then
            Set logSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            logSheet.Name = "RegistroProcesos"
            logSheet.Cells(1, 1).value = "Número de Contrato"
            logSheet.Cells(1, 2).value = "Fecha de Finalización"
        End If
    
        ' Encontrar la siguiente fila disponible en la hoja de registro
        nextLogRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1
        logSheet.Cells(nextLogRow, 1).value = Trim(Panel.ComboBox1.value)
        logSheet.Cells(nextLogRow, 2).value = Now() ' Fecha y hora actuales
        MsgBox "Carga finalizada con éxito.", vbInformation
    End If
        
' Activar y traer Excel al frente
ThisWorkbook.Activate
ThisWorkbook.Windows(1).Visible = True
Temporizador.Duracion.Caption = "00:00"
Unload Temporizador
Panel.Show vbModeless

Set session = Nothing
Set connection = Nothing
Set application = Nothing
Set sessionNuevoModo = Nothing

With ThisWorkbook.application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
End With

Call RellenarCeldasVacias(hojanombre)

End Sub

Sub hojaEntrada()
    Dim j As Integer
    Dim labelText As String
    Dim rowNumber As Integer
    Dim maxLabels As Integer
    Dim scrollPos As Long
    Dim prevScrollPos As Long
    Dim moreRows As Boolean
    Dim MaxScroll As Long
    Dim lastRow As Long
    Dim ValorHE As String
    Dim page As Boolean

    ' Inicializar valores
    maxLabels = 26 ' Cantidad máxima de etiquetas visibles por pantalla
    moreRows = True ' Bandera para saber si hay más filas
    scrollPos = 0 ' Posición inicial de la barra de desplazamiento
    MaxScroll = session.findById("wnd[0]/usr").verticalScrollbar.Maximum ' Máximo valor de la barra de scroll
    
    
   Do While moreRows And scrollPos < MaxScroll
        ' Bucle para capturar todas las etiquetas visibles en la pantalla actual
        If cancelarProceso Then Exit Sub
        DoEvents ' Permite que el botón "Cancelar" funcione

        For j = 0 To maxLabels
            On Error Resume Next
            lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
            labelText = session.findById("wnd[0]/usr/lbl[2," & 7 + j & "]").Text
            
            If Err.Number <> 0 Then
                Err.Clear ' Limpiar el error para continuar
                Exit For ' Si ocurre un error, salir del bucle.
            End If
            'Call pedido(rowcount)
            'Call SolPe(rowcount)
            On Error GoTo 0 ' Restaurar manejo normal de errores
            
              ' Verificar si labelText ya existe en la columna B antes de escribir
        If WorksheetFunction.CountIf(ws.Columns("B"), labelText) = 0 Or ws.Cells(ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, "B").value = Val(labelText) Then
            ' Si no existe, escribir en la hoja y ejecutar MSRV6
            'Debug.Print ws.Cells(ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, "B").value & WorksheetFunction.CountIf(ws.Columns("B"), labelText)
            
                If ws.Cells(ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, "B").value = Val(labelText) Then
                    ws.Cells(ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, 2).value = labelText ' Escribir en la hoja
                    
                    ws.Cells(ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, 3).value = pos(j)
                Else
                    ws.Cells(lastRow + 1, 2).value = labelText ' Escribir en la hoja
                    
                    ws.Cells(lastRow + 1, 3).value = pos(j)
                End If
            Call MSRV6
            
            ' Calcular tiempo restante en función de la posición actual
            duracionActual = duracionActual - 4
            If duracionActual < 0 Then duracionActual = 0 ' Asegurar que no sea negativo
            Temporizador.Duracion.Caption = FormatearTiempo(duracionActual)

        End If
        
        If cancelarProceso Then Exit Sub
        DoEvents ' Permite que el botón "Cancelar" funcione
            


        Next j

        ' Guardar posición actual de scroll
        prevScrollPos = scrollPos
        On Error Resume Next
        ' Mover la barra de desplazamiento hacia abajo
        session.findById("wnd[0]/usr").verticalScrollbar.Position = scrollPos + maxLabels
        scrollPos = session.findById("wnd[0]/usr").verticalScrollbar.Position
        On Error GoTo 0 ' Restaurar manejo normal de errores
             
        ' Si no cambia la posición de scroll, no hay más filas
        If prevScrollPos = scrollPos Then
            moreRows = False
        End If
Loop

    maxLabels = 25 ' Cantidad máxima de etiquetas visibles por pantalla
    For j = 0 To maxLabels
        On Error Resume Next
            lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
            labelText = session.findById("wnd[0]/usr/lbl[2," & 7 + j & "]").Text
                
                'Debug.Print Err.Number
            
            If Err.Number <> 0 Then
                Err.Clear ' Limpiar el error para continuar
                Exit For ' Si ocurre un error, salir del bucle y hacer scroll
            End If
            'Call pedido(rowcount)
            'Call SolPe(rowcount)
        On Error GoTo 0 ' Restaurar manejo normal de errores
        
              ' Verificar si labelText ya existe en la columna B antes de escribir
        If WorksheetFunction.CountIf(ws.Columns("B"), labelText) = 0 Or ws.Cells(ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, "B").value = Val(labelText) Then
            ' Si no existe, escribir en la hoja y ejecutar MSRV6
            
                If ws.Cells(ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, "B").value = Val(labelText) Then
                    ws.Cells(ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, 2).value = labelText ' Escribir en la hoja
                    ws.Cells(ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, 3).value = pos(j)
                Else
                    ws.Cells(lastRow + 1, 2).value = labelText ' Escribir en la hoja
                    ws.Cells(lastRow + 1, 3).value = pos(j)
                End If
            Call MSRV6
            
            ' Calcular tiempo restante en función de la posición actual
            duracionActual = duracionActual - 4
            If duracionActual < 0 Then duracionActual = 0 ' Asegurar que no sea negativo
            Temporizador.Duracion.Caption = FormatearTiempo(duracionActual)
            
        End If
            
            If cancelarProceso Then Exit Sub
            DoEvents ' Permite que el botón "Cancelar" funcione
            ' Calcular tiempo restante en función de la posición actual

    Next j

    session.findById("wnd[0]/tbar[0]/btn[3]").press
    'session.findById("wnd[0]/tbar[0]/btn[3]").press
    
End Sub
Function pedido(rowNumber As Integer)
    On Error GoTo ErrorHandler ' Manejo de errores
    'Dim ws As Worksheet
    'Set ws = ThisWorkbook.Sheets("Certificados")
    Dim value As String
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    ' Establece el foco en el campo SAP
    session.findById("wnd[0]/usr/lbl[2," & 7 + rowNumber & "]").SetFocus
    session.findById("wnd[0]").sendVKey 2 ' Simula una tecla (posiblemente F2)

    ' Extrae el valor del campo SAP
    value = session.findById("wnd[0]/usr/lbl[15,0]").Text
    
    ' Si quieres guardar el valor en la hoja de Excel, descomenta esta línea
     ws.Cells(lastRow + 1, 10).value = value
    
    ' Vuelve a la pantalla anterior en SAP
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    

    Exit Function

ErrorHandler:
    MsgBox "Ocurrió un error al intentar obtener el valor de SAP.", vbCritical
    pedido = "" ' En caso de error, se devuelve un valor vacío
End Function

Function pos(rowNumber As Integer) As String
    Dim posicion As String
    Dim ValorPos As String
    Dim MaxScroll As String
    Dim i As Integer
    Dim j As Integer
    Dim ValorHE As String
    Dim scrollPos As Integer

    ' Primera posición
    session.findById("wnd[0]/usr/lbl[2," & 7 + rowNumber & "]").SetFocus
    session.findById("wnd[0]").sendVKey 2
    ValorHE = Trim(session.findById("wnd[0]/usr/lbl[48,6]").Text)
    
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    scrollPos = 1
    
    j = 0

Do
    Dim encontrado As Boolean
    encontrado = False
    
    If cancelarProceso Then Exit Function
    DoEvents ' Permite que el botón "Cancelar" funcione
    
    For i = 0 To 20
        On Error Resume Next
        ValorPos = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/" & _
                "subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/" & _
                "tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[10," & j & "]").Text
        
        If Err.Number = 0 And ValorPos <> "" Then
            encontrado = True ' Encontró un valor válido
            Err.Clear ' Limpiar el error
            Exit For ' Salir del bucle interno
        End If
    Next i

    ' Si no encontró ningún valor después de todas las iteraciones, mover el scroll
    If Not encontrado Then
        Err.Clear
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/" & _
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").verticalScrollbar.Position = scrollPos
        scrollPos = scrollPos + 1
        j = j - 2
    End If
    
    If Not ValorHE = ValorPos Then
        j = j + 1 ' Seguir buscando en la siguiente fila
    End If

    On Error GoTo 0

Loop Until ValorHE = ValorPos


For i = 0 To 20
        On Error Resume Next
    posicion = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(i, "0000") & "/" & _
                                "subSUB2:SAPLMEVIEWS:1100/" & _
                                "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/" & _
                                "tblSAPLMEGUITC_1211/txtMEPO1211-KTPNR[29," & j & "]").Text
                                              
            If posicion <> "" Or Err.Number = 0 Then
                Err.Clear ' Limpiar el error para continuar
                Exit For ' Si ocurre un error, salir del bucle y hacer scroll
            End If
            
        On Error GoTo 0
Next i
                
' Si después del bucle no se encontró ningún valor, mostrar mensaje
If posicion = "" Or Err.Number <> 0 Then
    MsgBox "No se encontró ninguna posición para el índice " & j, vbExclamation, "Aviso"
End If


' Volver atrás dos veces
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


    pos = posicion
End Function

Sub MSRV6()
    Dim lastRow As Long
    Dim lastRow2 As Long
    Dim NumServ As Variant
    Dim textB As Variant
    Dim cant As Variant
    Dim valNet As Variant
    Dim sapGuiAuto As Object
    'Dim application As Object
    'Dim connection As Object
    'Dim session As Object
    'Dim sessionNuevoModo As Object
    Dim rowCount As Integer
    Dim i As Integer
    Dim nuevoModoAbierto As Boolean

    'Set ws = ThisWorkbook.Sheets("Certificados")
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    lastRow2 = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Conectar a SAP
    Set sapGuiAuto = GetObject("SAPGUI")
    Set application = sapGuiAuto.GetScriptingEngine
    Set connection = application.Children(0)
    Set session = connection.Children(0)

  ' Verificar si el modo con MSRV6 ya está abierto
    nuevoModoAbierto = False
    
    If cancelarProceso Then Exit Sub
    DoEvents ' Permite que el botón "Cancelar" funcione

        On Error Resume Next
        Set sessionNuevoModo = connection.Children(1)
        'Debug.Print sessionNuevoModo.Info.Transaction
        On Error GoTo 0

        ' Verificar si la sesión es válida y está activa
        If Not sessionNuevoModo Is Nothing Then
        
            If sessionNuevoModo.Info.Transaction = "MSRV6" Then
                nuevoModoAbierto = True
            End If
        End If

    ' Si no está abierto el modo, abrir un nuevo modo y acceder a MSRV6
    If Not nuevoModoAbierto Then
        ' Abrir un nuevo modo en SAP
 
        session.SendCommand "/o" ' Abrir nuevo modo con MSRV6
        session.findById("wnd[1]/tbar[0]/btn[5]").press
        
        ' Esperar a que la nueva ventana se abra
        Do While connection.Children.Count = 1 ' Mientras solo haya una ventana abierta
            DoEvents ' Permitir que la aplicación continúe funcionando
        Loop

        ' Referenciar al nuevo modo
        Set sessionNuevoModo = connection.Children(connection.Children.Count - 1)
'---------------------------------
'MSRV6
'---------------------------------
    'session.findById("wnd[1]").maximize
    sessionNuevoModo.findById("wnd[0]/tbar[0]/okcd").Text = "msrv6"
    sessionNuevoModo.findById("wnd[0]").sendVKey 0
        'sessionNuevoModo.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00024"
    End If

    ' pega los dato traido de Excel en SAP
    sessionNuevoModo.findById("wnd[0]/usr/ctxtEBELN-LOW").Text = ws.Range("B" & lastRow2).value
    'sessionNuevoModo.findById("wnd[0]/usr/btn%_EBELN_%_APP_%-VALU_PUSH").press
    'ws.Range("B" & lastRow2).Copy
    'sessionNuevoModo.findById("wnd[1]/tbar[0]/btn[16]").press
    'sessionNuevoModo.findById("wnd[1]/tbar[0]/btn[24]").press ' Botón para pegar los datos copiados

    ' Ejecutar acciones siguientes
    'sessionNuevoModo.findById("wnd[1]/tbar[0]/btn[8]").press ' Botón de continuar
    sessionNuevoModo.findById("wnd[0]/tbar[1]/btn[8]").press ' Botón de ejecutar

    ' Obtener la cantidad de filas
    rowCount = sessionNuevoModo.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").rowCount

    'On Error Resume Next ' Para evitar que el código se detenga en caso de error
    For i = 0 To rowCount - 2
        NumServ = sessionNuevoModo.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(i, "SRVPOS")
        textB = sessionNuevoModo.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(i, "KTEXT1")
        'cant = sessionNuevoModo.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(i, "MENGE")
        valNet = sessionNuevoModo.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(i, "NETWR")

        If Err.Number <> 0 Then
            Debug.Print "Error en la fila " & i & ": " & Err.Description
            Err.Clear
        Else
            ws.Cells(lastRow2 + i, 4).value = NumServ
            ws.Cells(lastRow2 + i, 5).value = textB
            'ws.Cells(lastRow + i , 6).value = cant
            
            Dim value As Double ' Cambia a Double para permitir decimales
            value = CDbl(Replace(valNet, ".", "")) ' Reemplaza el punto y convierte a Double
            
            If Err.Number <> 0 Then
                    valNet = Replace(valNet, ",", "") ' Quita separador de miles
                    value = CDbl(valNet) ' Convierte a número
                    Err.Clear
            End If
                
            ws.Cells(lastRow2 + i, 6).value = value
            ws.Cells(lastRow2 + i, 6).NumberFormat = "#.##0,00"
        End If
    Next i

    If cancelarProceso Then Exit Sub
    DoEvents ' Permite que el botón "Cancelar" funcione
    ' Cerrar el nuevo modo
    'sessionNuevoModo.findById("wnd[0]/tbar[0]/btn[15]").press ' Presiona el botón de cerrar en el nuevo modo
    sessionNuevoModo.findById("wnd[0]/tbar[0]/btn[3]").press

End Sub


Function SolPe(rowNumber As Integer)
    'Dim ws As Worksheet
    Dim lastRow As Long
    Dim valNet As Variant
    Dim sapGuiAuto As Object
    Dim application As Object
    Dim connection As Object
    Dim session As Object
    Dim sessionNuevoModo As Object
    Dim rowCount As Integer
    Dim i As Integer

    'Set ws = ThisWorkbook.Sheets("Certificados")
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    ' Conectar a SAP
    Set sapGuiAuto = GetObject("SAPGUI")
    Set application = sapGuiAuto.GetScriptingEngine
    Set connection = application.Children(0)
    Set session = connection.Children(0)
    
    rowCount = 0

  ' Verificar si el modo con ME53N ya está abierto
    nuevoModoAbierto = False

        On Error Resume Next
        Set sessionNuevoModo = connection.Children(2)
        'Debug.Print sessionNuevoModo.Info.Transaction
        On Error GoTo 0

        ' Verificar si la sesión es válida y está activa
        If Not sessionNuevoModo Is Nothing Then
        'Debug.Print sessionNuevoModo.Info.Transaction
            If sessionNuevoModo.Info.Transaction = "ME53N" Then
                nuevoModoAbierto = True
            End If
        End If

        
    ' Si no está abierto el modo, abrir un nuevo modo y acceder a MSRV6
    If Not nuevoModoAbierto Then
        ' Abrir un nuevo modo en SAP
 
        session.SendCommand "/o" ' Abrir nuevo modo con MSRV6
        session.findById("wnd[1]/tbar[0]/btn[5]").press
        
        ' Esperar a que la nueva ventana se abra
        Do While connection.Children.Count = 1 ' Mientras solo haya una ventana abierta
            DoEvents ' Permitir que la aplicación continúe funcionando
        Loop

        ' Referenciar al nuevo modo
        Set sessionNuevoModo = connection.Children(connection.Children.Count - 1)

        ' Aquí se realiza la navegación o acción dentro del nuevo modo
        sessionNuevoModo.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00004"
    End If

    'On Error Resume Next ' Para evitar que el código se detenga en caso de error
  
    sessionNuevoModo.findById("wnd[0]/tbar[1]/btn[17]").press
    'sessionNuevoModo.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/radMEPO_SELECT-BSTYP_F").SetFocus
    sessionNuevoModo.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/radMEPO_SELECT-BSTYP_F").Select
    sessionNuevoModo.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").Text = ws.Cells(rowNumber + 2, 10).value
    sessionNuevoModo.findById("wnd[1]").sendVKey 0
    
    posicion = sessionNuevoModo.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-KTPNR[29,0]").Text

        If Err.Number <> 0 Then
            Debug.Print "Error en la fila " & rowNumber & ": " & Err.Description
            Err.Clear
        Else

            ws.Cells(rowNumber + 2, 3).value = posicion

        End If

    'On Error GoTo 0 ' Vuelve a la gestión normal de errores

    ' Cerrar el nuevo modo
    'sessionNuevoModo.findById("wnd[0]/tbar[0]/btn[15]").press ' Presiona el botón de cerrar en el nuevo modo

End Function

Sub infoServ()
'Dim ws As Worksheet
Dim rowCount As Integer
Dim i As Integer
Dim sapTable As Object
Dim NumServ As Variant, textB As Variant, cant As Variant, valNet As Variant
Dim lastRow As Long
Dim value As Double
Dim filasConInfo As Integer

'Set ws = ThisWorkbook.Sheets("Certificados")
lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row


' Obtener la tabla
Set sapTable = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
         "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/" & _
         "subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW")
     
     ' Verificar el número total de filas en la tabla
rowCount = sapTable.rowCount

' Inicializar el contador de filas con información
    filasConInfo = 0
    
    ' Iterar a través de todas las filas de la tabla
    For i = 0 To rowCount - 1
 On Error Resume Next
      NumServ = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
        "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/" & _
        "subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/ctxtESLL-SRVPOS[2," & i & "]").Text
 On Error GoTo 0
        ' Si la columna tiene datos (no está vacía), contarla como una fila válida
        If Trim(NumServ) <> "" Then
            filasConInfo = filasConInfo + 1
        End If
    Next i


' Iterar a través de todas las filas de la tabla en SAP
For i = 0 To filasConInfo - 1
    On Error Resume Next ' Para evitar que el código se detenga en caso de error
    
    ' Obtener los valores de las celdas para la fila i
      NumServ = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/" & _
        "subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/" & _
        "subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/ctxtESLL-SRVPOS[2," & i & "]").Text
        'Debug.Print "Número de servicio en la fila " & numServ
        
    textB = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
        "subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/" & _
        "txtESLL-KTEXT1[3," & i & "]").Text
                'Debug.Print "Número de text en la fila " & textB
        
    cant = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
        "subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/" & _
        "txtESLL-MENGE[4," & i & "]").Text
                'Debug.Print "Número de cant en la fila " & cant
                
    valNet = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" & _
    "subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/" & _
    "txtESLL-TBTWR[6," & i & "]").Text
            'Debug.Print "Número de precio en la fila " & valNet
            
    If Err.Number <> 0 Then
        Debug.Print "Error en la fila " & i & ": " & Err.Description
        Err.Clear
    Else
        ' Colocar los valores en las celdas de Excel
        ws.Cells(lastRow + i + 1, 4).value = NumServ
        ws.Cells(lastRow + i + 1, 5).value = textB
        'ws.Cells(lastRow + i + 1, 6).value = cant
        
        ' Convertir valNet a número y colocar en Excel
        value = CDbl(Replace(valNet, ".", "")) ' Reemplaza el punto decimal para la conversión
        ws.Cells(lastRow + i + 1, 6).value = value
    End If
    
    On Error GoTo 0 ' Vuelve a la gestión normal de errores
Next i

End Sub


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
        ThisWorkbook.application.SendKeys "{ESC}"

        'MsgBox "Ventana de Windows con título '" & WindowTitle & "' detectada y acción enviada.", vbInformation
    Else
        'MsgBox "La ventana de Windows con título '" & WindowTitle & "' NO fue encontrada.", vbExclamation
    End If

End Sub




