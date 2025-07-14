Attribute VB_Name = "Módulo1"
'*************************************
'|||REALIZADO POR HERNAN F. CARRIZO|||
'************AGO-SEP 2024*************
Option Explicit
Public numCont As String

Sub Contratos()
    Dim ws1 As Worksheet
    'Dim ws2 As Worksheet
    'Dim ws3 As Worksheet
    'Dim ComboBox1 As Object
    'Dim lastRow As Long
    'Dim i As Long

    ' Establecer la hoja de trabajo
    Set ws1 = ThisWorkbook.Sheets(numCont)
    'Set ws2 = ThisWorkbook.Sheets("Contratos")
 
    'Set ComboBox1 = ws1.OLEObjects("ComboBox1").Object
    
    ' Obtener la última fila con datos en la columna A
    'lastRow = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    
    ' Limpiar el ListBox antes de agregar elementos
    'ComboBox1.Clear
    numCont = Trim(ws1.Cells(7, 4).Value)
    
    ' Cargar los valores de la columna A en el ListBox
    'For i = 2 To lastRow
        'ComboBox1.AddItem ws2.Cells(i, 1).value
    'Next i
    
    ' Asegúrate de que el ComboBox tiene un valor seleccionado
    'If ComboBox1.ListIndex <> -1 Or ComboBox1.value <> "" Then
        ' Asigna el valor seleccionado a la variable NumCont
        'numCont = Trim(ComboBox1.value)
   ' End If


End Sub

Sub MostrarCotizacion()
    Dim pag As Object
    Dim html As Object
    Dim euroCompra As String
    Dim euroVenta As String
    Dim dolarCompra As String
    Dim dolarVenta As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets(numCont)
    
    ' Crear el objeto InternetExplorer
    Set pag = CreateObject("InternetExplorer.Application")
    pag.Visible = False ' Para hacerlo visible cambiar a True
    
    ' Cargar la página web
    pag.Navigate "https://www.lanacion.com.ar/tema/euro-hoy-tid66142/"
    
    ' Esperar a que la página se cargue
    Do While pag.Busy Or pag.ReadyState <> 4
        DoEvents
        If ws.Range("W4").Font.Color = RGB(0, 0, 255) And ws.Range("Y4").Font.Color = RGB(0, 0, 255) Then
            ws.Range("W4").Font.Color = RGB(0, 0, 0)
            ws.Range("Y4").Font.Color = RGB(0, 0, 0)
        Else
            ws.Range("W4").Font.Color = RGB(0, 0, 255)
            ws.Range("Y4").Font.Color = RGB(0, 0, 255)
    End If
    
    ws.Range("W4").Font.Color = RGB(0, 0, 0)
    ws.Range("Y4").Font.Color = RGB(0, 0, 0)
            
    Loop
    
    ' Obtener el HTML de la página
    Set html = pag.Document
    
    ' Buscar y extraer los valores
    ExtractCurrencyValues html, "Euro", euroCompra, euroVenta
    ExtractCurrencyValues html, "Dólar oficial", dolarCompra, dolarVenta
    
    ' Cerrar el objeto InternetExplorer
    pag.Quit
    Set pag = Nothing
    
    Dim promedioEuro As Double
    Dim promedioDolar As Double
    promedioDolar = (Val(dolarVenta) + Val(dolarCompra)) / 2
    promedioEuro = (Val(euroVenta) + Val(euroCompra)) / 2
    
    ws.Range("W4").Value = "Valor Promedio Euro: "
    ws.Range("X4").Value = promedioEuro
    ws.Range("Y4").Value = "Valor Promedio Dolar: "
    ws.Range("Z4").Value = promedioDolar
   
End Sub

Sub ExtractCurrencyValues(html As Object, currencyType As String, ByRef compra As String, ByRef venta As String)
    Dim allH2 As Object
    Dim h2 As Object
    Dim p As Object
    Dim spans As Object
    Dim strongs As Object
    Dim i As Integer
    
    ' Obtener todos los elementos h2
    Set allH2 = html.getElementsByTagName("h2")
    
    ' Iterar sobre cada elemento h2
    For Each h2 In allH2
        ' Verificar la clase y el texto para encontrar el tipo de moneda
        If h2.className = "dolar-title --fourxs" And h2.innerText = currencyType Then
            ' Buscar el siguiente elemento que es un nodo de tipo elemento
            Set p = h2.ParentNode.NextSibling
            Do While Not p Is Nothing
                ' Asegurarse de que p sea un elemento
                If p.NodeType = 1 And p.tagName = "P" And p.className = "com-text --sixxs" Then
                    ' Obtener todos los span y strong dentro de p
                    Set spans = p.getElementsByTagName("span")
                    Set strongs = p.getElementsByTagName("strong")
                    
                    ' Iterar sobre los span y strong
                    For i = 0 To spans.Length - 1
                        If spans(i).innerText = "Compra" Then
                            ' Eliminar caracteres no deseados y convertir el valor a numérico
                            compra = Replace(Replace(strongs(i).innerText, "$", ""), ",", ".")
                        ElseIf spans(i).innerText = "Venta" Then
                            ' Eliminar caracteres no deseados y convertir el valor a numérico
                            venta = Replace(Replace(strongs(i).innerText, "$", ""), ",", ".")
                        End If
                    Next i
                    Exit For
                End If
                ' Avanzar al siguiente nodo
                Set p = p.NextSibling
            Loop
        End If
    Next h2
End Sub


Sub VerificarSAP()
    Dim usuario As String, contrasena As String, valorCampo As String
    Dim dia As Integer, mes As Integer, anio As Integer
    Dim fechaConvertida As Date, fecha_fin_Ctto As Date, valorCampoConvertido As Variant
    Dim i As Integer, j As Integer, maxRows As Integer
    Dim ws As Worksheet
    
           
    numCont = Trim(ThisWorkbook.Sheets(1).Range("D7").Value)
    ThisWorkbook.Sheets(1).Name = numCont
    
    ' Si la celda está vacía, muestra un mensaje y termina
       If numCont = "" Then
           MsgBox "El valor en la celda D7 no puede estar vacío.", vbExclamation, "Error"
           Exit Sub
       End If

    
    Set ws = ThisWorkbook.Sheets(numCont)
    
    ' Mostrar el formulario para ingresar usuario y contraseña
    Inicio.Show vbModal
    'Call Contratos
    
' Comprobar si el formulario se cerró correctamente
    If Inicio.Tag = "OK" Then
        ' Verificar si el ComboBox tiene un valor seleccionado
        If numCont <> "" Then
            numCont = Trim(numCont)
       
        usuario = Inicio.txtUsuario.Text
        contrasena = Inicio.txtContraseña.Text
        Else
            MsgBox "Por favor, seleccione un Contrato."
            Exit Sub
        End If

        ' Iniciar SAP después de cerrar el formulario
            
        Call MostrarCotizacion
        If IniciarSAP(usuario, contrasena) Then
        Else
            Exit Sub
        End If
        
    Else
        ' El formulario se cerró sin guardar las credenciales
        Exit Sub
    End If

    Call VolverAVentanaPrincipalSAP
    
    ' -----------------------------------------------
    ' ZM57
    ' -----------------------------------------------
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").Text = "zm57"
    session.findById("wnd[0]").SendVKey 0

            
    If Err.Number <> 0 Then
        Set connection = Nothing
        Set session = Nothing
        Exit Sub ' Salir del bucle si ocurre un error
    End If
    On Error GoTo 0

    session.findById("wnd[0]/usr/ctxtS_KDATE-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtS_KDATE-LOW").SetFocus
    session.findById("wnd[0]").SendVKey 0

    ' Colocar el número de contrato ingresado
    session.findById("wnd[0]/usr/ctxtS_EBELN-LOW").Text = numCont
    session.findById("wnd[0]/usr/ctxtS_EBELN-LOW").SetFocus
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    ' Validar si se encontraron resultados para el número de contrato
    Dim filaContratos As Integer
    filaContratos = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").rowCount
    
    If filaContratos = 0 Then
        MsgBox "No se encontraron resultados para el número de contrato ingresado: " & numCont & ".", vbExclamation
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        Exit Sub
    End If

    ' Obtener y copiar el valor de la columna Moneda Contrato
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "WAERS"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SetFocus

    Dim Moneda As String
    Moneda = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "WAERS")
    ws.Range("R17").Value = Moneda
    
    ' Obtener y copiar el valor de la columna Monto Contrato
    
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "KTWRT"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SetFocus
    valorCampo = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "KTWRT")
    ws.Range("N17").Value = valorCampo
    
    ' Obtener y copiar el valor de la columna Monto Residual
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "SALDO_SIN_COMP"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SetFocus
    valorCampo = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "SALDO_SIN_COMP")
    ws.Range("N18").Value = valorCampo
    
    ' Cerrar la transacción
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    
    '-----------------------------------------
    'ME33K
    '-----------------------------------------
    
    'session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00005"
    session.findById("wnd[0]/tbar[0]/okcd").Text = "me33k"
    session.findById("wnd[0]").SendVKey 0
  
    'Se inserta el numero de Contrato elegido

    session.findById("wnd[0]/usr/ctxtRM06E-EVRTN").SetFocus
    session.findById("wnd[0]/usr/ctxtRM06E-EVRTN").Text = numCont
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    
    ' Copiar el nombre del Porveedor
    session.findById("wnd[0]/usr/txtLFA1-NAME1").SetFocus
    valorCampo = session.findById("wnd[0]/usr/txtLFA1-NAME1").Text
    ws.Range("M7").Value = valorCampo
    
    ' Copiar el nombre del Servicio
    session.findById("wnd[0]/usr/ssubCUSTSCR1:SAPLXM06:0201/ctxtEKKO_CI-ZMERCO").SetFocus
    valorCampo = session.findById("wnd[0]/usr/ssubCUSTSCR1:SAPLXM06:0201/ctxtEKKO_CI-ZMERCO").Text
    ws.Range("D13").Value = valorCampo
    
    ' Copiar el GM
    session.findById("wnd[0]/usr/ssubCUSTSCR1:SAPLXM06:0201/txtEKKO_CI-ZZTITULO").SetFocus
    valorCampo = session.findById("wnd[0]/usr/ssubCUSTSCR1:SAPLXM06:0201/txtEKKO_CI-ZZTITULO").Text
    ws.Range("D11").Value = valorCampo

    ' Copiar Fecha Inicio Contrato
    session.findById("wnd[0]/usr/ctxtEKKO-KDATB").SetFocus
    valorCampo = session.findById("wnd[0]/usr/ctxtEKKO-KDATB").Text
    dia = Left(valorCampo, 2)
    mes = Mid(valorCampo, 4, 2)
    anio = Right(valorCampo, 4)
    fechaConvertida = DateSerial(CLng(anio), CLng(mes), CLng(dia))
    'fecha inicio Contrato
    ws.Range("D15").Value = fechaConvertida

    ' Copiar Fecha Fin Operativa Contrato
    session.findById("wnd[0]/usr/ctxtEKKO-KDATE").SetFocus
    valorCampo = session.findById("wnd[0]/usr/ctxtEKKO-KDATE").Text
    dia = Left(valorCampo, 2)
    mes = Mid(valorCampo, 4, 2)
    anio = Right(valorCampo, 4)
    fecha_fin_Ctto = DateSerial(CLng(anio), CLng(mes), CLng(dia))
    'fecha finalizacion actual
    ws.Range("R15").Value = fecha_fin_Ctto
        
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/mbar/menu[2]/menu[4]/menu[1]").Select
    session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").selectColumn "UDATE"
    session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").pressToolbarButton "&SORT_DSC"
    session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").selectColumn "FTEXT"
    session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").pressToolbarButton "&SORT_DSC"
    
    On Error Resume Next
    valorCampo = session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").GetCellValue(0, "FTEXT")
    On Error GoTo 0
    
    If valorCampo = "Valor total en liberación" Then
    
    valorCampo = session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").GetCellValue(0, "UDATE")
    dia = Left(valorCampo, 2)
    mes = Mid(valorCampo, 4, 2)
    anio = Right(valorCampo, 4)
    fechaConvertida = DateSerial(CLng(anio), CLng(mes), CLng(dia))
    'fecha ultima modificacion aumento monto
    ws.Range("Q21").Value = fechaConvertida
        
    valorCampo = session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").GetCellValue(0, "F_OLD")
    
    i = 0
    Do While session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").GetCellValue(i, "F_OLD") <> "0,00 " & Moneda
        i = i + 1
    Loop
        
    valorCampo = session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").GetCellValue(i, "F_NEW")
    
    valorCampo = Replace(valorCampo, " " & Moneda, "")
    valorCampo = CDbl(valorCampo)
    
    'Valor omonto del contrato
    ws.Range("P53").Value = valorCampo
    ws.Range("P53").NumberFormat = "#,##0.00"
    
    'Numero de cambios de montos
    ws.Range("J21").Value = i
    
    If i = 0 Then
        ws.Range("Q21").Value = ""
    End If
    
    Else
        MsgBox "no encontro datos"
    End If
    
    session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").selectColumn "UDATE"
    session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").pressToolbarButton "&SORT_ASC"
    session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").selectColumn "FTEXT"
    session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").pressToolbarButton "&SORT_DSC"
    
    ' Establece un límite máximo para el número de filas a recorrer
    maxRows = session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").rowCount
    j = 0
    ' Buscar "Fin período validez"
    For i = 0 To maxRows - 1
        If (session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").GetCellValue(i, "FTEXT") = "Fin período validez") Then
            ' Bucle para encontrar "Fin período validez"
            Do While session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").GetCellValue(i, "FTEXT") <> "Fin período validez"
                i = i + 1
            Loop
            
        valorCampoConvertido = Format(CDate(Replace(fecha_fin_Ctto, ".", "/")), "dd.mm.yyyy")
            Do While (session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").GetCellValue(i + j, "F_NEW")) <> valorCampoConvertido
                j = j + 1
            Loop
            
        'cantidad de cambios de vigencia
        ws.Range("J22").Value = j + 1
        
        valorCampo = session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").GetCellValue(j + i, "UDATE")
        dia = Left(valorCampo, 2)
        mes = Mid(valorCampo, 4, 2)
        anio = Right(valorCampo, 4)
        fechaConvertida = DateSerial(CLng(anio), CLng(mes), CLng(dia))
        'fecha ultima modificacion vigencia
        ws.Range("Q22").Value = fechaConvertida
        
        valorCampo = session.findById("wnd[0]/usr/subME_CHANGES:SAPLMECD2:0100/cntlMEALV_GRID_CONTROL_MECD2/shellcont/shell").GetCellValue(i, "F_OLD")
        dia = Left(valorCampo, 2)
        mes = Mid(valorCampo, 4, 2)
        anio = Right(valorCampo, 4)
        fechaConvertida = DateSerial(CLng(anio), CLng(mes), CLng(dia))
        'fecha finalizacion original del contrato
        ws.Range("K15").Value = fechaConvertida
        
        Exit For
        
        Else
            ws.Range("K15").Value = ws.Range("R15").Value
            ws.Range("J22").Value = j
        End If
    Next i
    
    ' Cerrar la transacción
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    
    
    connection.CloseSession ("ses[0]")
    Set session = Nothing
    Set connection = Nothing
    Set application = Nothing
End Sub




    
