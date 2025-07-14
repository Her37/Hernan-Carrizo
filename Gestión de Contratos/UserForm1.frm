VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "WORKS & SERVICE "
   ClientHeight    =   11028
   ClientLeft      =   -300
   ClientTop       =   -1812
   ClientWidth     =   17460
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'|||REALIZADO POR HERNAN F. CARRIZO|||
'************MAY-AGO 2024*************

Option Explicit

Dim ws As Worksheet
Dim wsZ As Worksheet
Dim wsM As Worksheet
Dim wsD As Worksheet
Dim nombreArchivo As String
Dim libro As Workbook
Dim rutaArchivo As String
Dim a As Long
Dim ErrorMsg As Integer
Dim macro As String
Dim WithEvents wb As WebBrowser
Attribute wb.VB_VarHelpID = -1
Public j As Integer
Public Bottom As Boolean
Public numCont As Variant

Sub Act_List2()
    Dim wsOrigen As Worksheet
    Dim ultimaFila As Long
    Dim ListItem As ListItem
    Dim datos As Variant
    Dim comentarios As Object
    Dim i As Long
    Dim numCont As String

    CheckBox1.Value = True

    ' Define la hoja de origen
    Set wsOrigen = ThisWorkbook.Sheets("Superados")

    ' Encuentra la última fila con datos en columna B
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, "B").End(xlUp).Row

    ' Leer toda la hoja de datos en una matriz
    datos = wsOrigen.Range("A1:R" & ultimaFila).Value ' Ajusta el rango según tus columnas

    ' Crear un diccionario para los comentarios
    Set comentarios = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(datos, 1)
        If Not IsEmpty(datos(i, 2)) Then ' Columna B
            If Not wsOrigen.Cells(i, 2).Comment Is Nothing Then
                comentarios(datos(i, 2)) = True
            End If
        End If
    Next i

    ' Configurar el ListView2
    With Me.ListView2
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False
        .ColumnHeaders.Clear
        .ListItems.Clear
            
        ' Agregar los encabezados manualmente
        .ColumnHeaders.Add , , "Fila", 30
        .ColumnHeaders.Add , , "Contrato", 60
        .ColumnHeaders.Add , , "Descripción", 180
        .ColumnHeaders.Add , , "G.Merc.", 40
        .ColumnHeaders.Add , , "Proveedor", 70
        .ColumnHeaders.Add , , "F. Hasta", 50
    End With

    ' Agregar elementos al ListView2
    For i = 2 To UBound(datos, 1) ' Empieza en la fila 2 (omitir encabezados)
        Set ListItem = Me.ListView2.ListItems.Add(, , i)
        ListItem.ListSubItems.Add , , datos(i, 2) ' Contrato
        ListItem.ListSubItems.Add , , datos(i, 4) ' Descripción
        ListItem.ListSubItems.Add , , datos(i, 6) ' G.Merc.
        ListItem.ListSubItems.Add , , datos(i, 8) ' Proveedor
        ListItem.ListSubItems.Add , , datos(i, 17) ' F. Hasta

        ' Cambiar color si hay un comentario
        numCont = datos(i, 2)
        If comentarios.exists(numCont) Then
            ListItem.ForeColor = RGB(255, 0, 0) ' Cambia el color a rojo
        End If
    Next i

    ' Actualiza el título del marco con el número de contratos vencidos
    UserForm1.Frame2.Caption = "Contratos Finalizados Históricos : " & UserForm1.ListView2.ListItems.Count
End Sub


Private Sub CheckBox1_Click()
Dim item As MSComctlLib.ListItem

    ' Verifica el estado del CheckBox
    If CheckBox1.Value = True Then
        
    Call Act_List2
    
    Else
        ' Si CheckBox1 está desmarcado, llama a la subrutina Cont_Venc
        Call Cont_Venc
        For Each item In Me.ListView2.ListItems
            item.Selected = False
        Next item
    End If
End Sub

Private Sub Act_List1()
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim datos As Variant
    Dim contractDict As Object
    Dim ListItem As ListItem
    Dim i As Long
    Dim numCont As String
    
    CheckBox2.Value = True
    
    ' Define la hoja de origen
    Set ws = ThisWorkbook.Sheets("Grids")
    
    ' Encuentra la última fila con datos en columna B
    ultimaFila = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Leer los datos en memoria
    datos = ws.Range("A1:H" & ultimaFila).Value ' Ajustar columnas según necesidad
    
    ' Crear el diccionario para almacenar contratos y comentarios
    Set contractDict = CreateObject("Scripting.Dictionary")
    
    ' Configura el ListView1
    With Me.ListView1
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        ' Agregar los encabezados de columna
        .ColumnHeaders.Add , , "Fila", 30
        .ColumnHeaders.Add , , datos(1, 2), 60
        .ColumnHeaders.Add , , datos(1, 4), 190
        .ColumnHeaders.Add , , datos(1, 6), 50
        .ColumnHeaders.Add , , datos(1, 8), 100
    End With
    
    ' Llenar el ListView1 con contratos activos
    For i = LBound(datos, 1) + 1 To UBound(datos, 1) ' Desde la segunda fila (ignorar encabezados)
        If datos(i, 3) = "ACTIVO" Or ws.Cells(i, 3).Interior.Color = RGB(0, 255, 0) Then
            Set ListItem = Me.ListView1.ListItems.Add(, , i)
            ListItem.ListSubItems.Add , , datos(i, 2) ' Contrato
            ListItem.ListSubItems.Add , , datos(i, 4) ' Descripción
            ListItem.ListSubItems.Add , , datos(i, 6) ' Otro dato
            ListItem.ListSubItems.Add , , datos(i, 8) ' Otro dato
        End If
    Next i
    
    ' Llenar el diccionario con los números de contrato y comentarios
    For i = LBound(datos, 1) + 1 To UBound(datos, 1)
        Dim contractNumber As String
        If UCase(datos(i, 3)) = "NO" Then
            Exit For
        Else
            contractNumber = LCase(datos(i, 2))
            If Not contractDict.exists(contractNumber) Then
                contractDict.Add contractNumber, Not ws.Cells(i, 2).Comment Is Nothing
            End If
        End If
    Next i
    
    ' Iterar sobre cada ítem en el ListView y actualizar el color si hay comentario
    For Each ListItem In Me.ListView1.ListItems
        numCont = LCase(ListItem.SubItems(1))
        If contractDict.exists(numCont) And contractDict(numCont) Then
            ListItem.ForeColor = RGB(255, 0, 0) ' Cambia el color a rojo
        End If
    Next ListItem
    
    ' Liberar el diccionario
    Set contractDict = Nothing
    
    ' Actualiza el título del marco con el número de contratos activos
    Me.Frame1.Caption = "Contratos Activos : " & Me.ListView1.ListItems.Count
End Sub


Private Sub CheckBox2_Click()
Dim item As MSComctlLib.ListItem

    ' Verifica el estado del CheckBox
    If CheckBox2.Value = True Then

     Call Act_List1
        
    Else
        ' Si CheckBox2 está desmarcado, llama a la subrutina Cont_Act
        Call Cont_Act
        For Each item In Me.ListView1.ListItems
            item.Selected = False
        Next item
    End If
End Sub

Private Sub CommandButton10_Click()
    Dim ws As Worksheet
    Dim selectedIndex As Variant
    Dim numCont As Variant
    Dim comentario As String
    Dim coleccion As Collection
    Dim item As MSComctlLib.ListItem
    Dim contractDict As Object
    Dim i As Long
    Dim lastRow As Long

    Set coleccion = New Collection
    Set ws = ThisWorkbook.Sheets("Grids")
    Set contractDict = CreateObject("Scripting.Dictionary")

    ' Llenar el diccionario con números de contrato e índices de filas de Grids
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 1 To lastRow
        Dim contractNumber As String
        contractNumber = LCase(ws.Cells(i, 2).Value)
        If Not contractDict.exists(contractNumber) Then
            contractDict.Add contractNumber, i
        End If
    Next i

    ' Llenar el diccionario con números de contrato e índices de filas de Superados
    lastRow = wsD.Cells(wsD.Rows.Count, "B").End(xlUp).Row
    For i = 1 To lastRow
        Dim contractNumberD As String
        contractNumberD = LCase(wsD.Cells(i, 2).Value)
        If Not contractDict.exists(contractNumberD) Then
            contractDict.Add contractNumberD, i
        End If
    Next i

    ' Contar los ítems seleccionados
    For Each item In Me.ListView2.ListItems
        If item.Selected Then
            coleccion.Add item.index
        End If
    Next item

    ' Asegúrate de que hay al menos un ítem seleccionado
    If coleccion.Count = 0 Then
        MsgBox "No se ha seleccionado ningún ítem.", vbExclamation, "Advertencia"
        Exit Sub
    End If

    ' Iterar sobre los índices de los ítems seleccionados
    For Each selectedIndex In coleccion
        ' Obtener el ítem actual usando el índice de la colección
        Set item = Me.ListView2.ListItems(selectedIndex)
        numCont = LCase(item.SubItems(1))

        ' Pedir al usuario que ingrese el comentario
        comentario = InputBox("Ingrese el comentario que desea agregar: " & vbCrLf & "Contrato seleccionado: " & numCont, "Comentario")

        ' Verificar si el usuario ingresó un comentario
        If comentario = "" Then
            MsgBox "No se ingresó ningún comentario.", vbExclamation
            Exit Sub
        End If

        ' Determinar la hoja de trabajo a usar
        Dim currentWs As Worksheet
        If CheckBox1.Value = False Then
            Set currentWs = ws
        Else
            Set currentWs = wsD
            item.ForeColor = RGB(255, 100, 30)
        End If

        ' Agregar o actualizar el comentario en la hoja correspondiente
        If contractDict.exists(numCont) Then
            Dim rowIndex As Long
            rowIndex = contractDict(numCont)
            With currentWs.Cells(rowIndex, 2)
                If Not .Comment Is Nothing Then
                    .Comment.Visible = False
                    .Comment.Text Text:=comentario
                Else
                    .AddComment comentario
                    .Comment.Visible = True
                End If
            End With
        End If
    Next selectedIndex

    ' Liberar el diccionario
    Set contractDict = Nothing
End Sub


Private Sub CommandButton11_Click()
 Dim ws As Worksheet
    Dim selectedIndex As Integer
    Dim numCont As Variant
    Dim i As Integer
    Dim comentario As String
    Dim item As MSComctlLib.ListItem

    Set ws = ThisWorkbook.Sheets("Grids")
    
    For Each item In Me.ListView1.ListItems
        If item.Selected Then

If CheckBox2.Value = False Then

    numCont = Me.ListView1.selectedItem.SubItems(1)
     ' Pedir al usuario que ingrese el comentario
    comentario = InputBox("Ingrese el comentario que desea agregar al contrato: " & numCont, "Comentario")

    ' Verificar si el usuario ingresó un comentario
    If comentario = "" Then
        MsgBox "No se ingresó ningún comentario.", vbExclamation
        Exit Sub
    End If

    For i = 1 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        If LCase(ws.Cells(i, 2).Value) Like LCase(numCont) Then
            If Not ws.Cells(i, 2).Comment Is Nothing Then
                ws.Cells(i, 2).Comment.Visible = False
                ws.Cells(i, 2).Comment.Text Text:=comentario
            Else
                ws.Cells(i, 2).AddComment comentario
                ws.Cells(i, 2).Comment.Visible = True
            End If
            Exit For
        End If
    Next i

End If

If CheckBox2.Value = True Then

    numCont = Me.ListView1.selectedItem.SubItems(1)
    ' Pedir al usuario que ingrese el comentario
    comentario = InputBox("Ingrese el comentario que desea agregar al contrato: " & numCont, "Comentario")

    ' Verificar si el usuario ingresó un comentario
    If comentario = "" Then
        MsgBox "No se ingresó ningún comentario.", vbExclamation
        Exit Sub
    End If

    For i = 1 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        If LCase(ws.Cells(i, 2).Value) Like LCase(numCont) Then
            If Not ws.Cells(i, 2).Comment Is Nothing Then
                ws.Cells(i, 2).Comment.Visible = False
                ws.Cells(i, 2).Comment.Text Text:=comentario
            Else
                ws.Cells(i, 2).AddComment comentario
                ws.Cells(i, 2).Comment.Visible = True
            End If
            Exit For
        End If
    Next i
End If
End If
Next item
End Sub

Private Sub CommandButton12_Click()
    Call EnviarCorreo
End Sub


Private Sub CommandButton14_Click()
Call DescargaSAPzco9
End Sub

Private Sub CommandButton3_Click()
    Call SeleccionarArchivo
End Sub

Sub SeleccionarArchivo()
    MsgBox "Por favor, seleccione la última Gestión Grids realizada.", vbInformation, "Work&Service"
    ' Solicitar al usuario que seleccione el archivo
    rutaArchivo = ThisWorkbook.application.GetOpenFilename("Archivos de Excel (*.xls; *.xlsx ; *.xlsm), *.xls; *.xlsx; *.xlsm", , "Seleccione el archivo de Excel")
    
    If rutaArchivo = "Falso" Then
        MsgBox "No se seleccionó ningún archivo.", vbExclamation
        Exit Sub
    End If
      
    ' Abrir el archivo seleccionado
    Set libro = Workbooks.Open(rutaArchivo)
    nombreArchivo = libro.Name
    Me.TextBox1.Value = nombreArchivo
    
    ' Configurar hojas de trabajo
    Set ws = ThisWorkbook.Sheets("Grids")
    Set wsM = ThisWorkbook.Sheets(4)
    
    On Error Resume Next
    Set wsZ = libro.Sheets("Grids")
        ' Verificar si ocurrió un error
    If Err.Number <> 0 Then
        'MsgBox "Ocurrió un error: " & Err.Description
        MsgBox "Ingrese un archivo valido", vbExclamation
        libro.Close
        Me.TextBox1.Value = ""
        rutaArchivo = ""
        ' Limpiar el objeto Err
        Err.Clear
        ' Desactivar el manejo de errores
        On Error GoTo 0
        Exit Sub
    End If
    
    Set wsD = Workbooks(macro).Worksheets("Superados")
   
 On Error Resume Next
 
        ws.AutoFilter.Sort.SortFields.Clear
        ws.AutoFilter.Sort.SortFields.Add2 Key:=Range _
        ("C1:C" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortTextAsNumbers
        
    With ws.AutoFilter.Sort
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
On Error GoTo 0
ThisWorkbook.Activate

End Sub

Private Sub CommandButton4_Click()
    Call RealizarCorte
End Sub

Sub RealizarCorte()
    Dim Hasta As Long
    Dim Desde As Long
    Dim celda As Range
    Dim searchValue As String
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim wsB As Worksheet
    Dim rangoOrigen As Range
    Dim rangoDestino As Range
    Dim rangoBusqueda As Range
    Dim filaCorte As Long

    ' Verificar que se haya ingresado una fecha de corte
    If Me.ComboBox1.Value = "" Then
        searchValue = InputBox("Por favor, ingresa Fecha de Corte:", "Fecha Hasta:")
        If searchValue <> "" Then
            Me.ComboBox1.Value = searchValue
            Me.ComboBox1.SetFocus
        Else
            MsgBox "No se ingresó Fecha de Corte.", vbExclamation, "Error Búsqueda"
            Me.ComboBox1.SetFocus
            Exit Sub
        End If
    End If

    ' Desactivar el modo de filtro automático en ambas hojas
    Worksheets("Grids").AutoFilterMode = False
    Worksheets(5).AutoFilterMode = False

    Me.Frame1.Caption = "Contratos Nuevos: "
    Me.Frame2.Caption = "Contratos Vencidos: "

    ' Configurar el ListView1
    With Me.ListView1
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False
        .ColumnHeaders.Clear
        .ListItems.Clear
    End With

    ' Configurar el ListView2
    With Me.ListView2
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False
        .ColumnHeaders.Clear
        .ListItems.Clear
    End With

    ' Definir las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets(4)
    Set wsDestino = ThisWorkbook.Sheets("Grids")
    Set wsB = ThisWorkbook.Sheets("Base Trabajo")

    ' Definir el rango de origen y destino
    Set rangoOrigen = wsOrigen.Range("A1:AZ" & wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row)
    Set rangoDestino = wsDestino.Range("A1")

    ' Desactivar la actualización de pantalla para mejorar el rendimiento
    ThisWorkbook.application.ScreenUpdating = False

    ' Copiar los datos del rango de origen al rango de destino
    rangoOrigen.Copy Destination:=rangoDestino
    wsDestino.Rows(1).Font.Bold = True

    ' Obtener el valor del ComboBox
    searchValue = Me.ComboBox1.Value

    ' Actualizar el rango de búsqueda en columna P
    Set rangoBusqueda = wsDestino.Range("P2:P" & wsDestino.Cells(wsDestino.Rows.Count, "P").End(xlUp).Row)

    ' Buscar el valor en el rango
    On Error Resume Next
    Set celda = rangoBusqueda.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0

    ' Verificar si se encontró el valor
    If Not celda Is Nothing Then
        Desde = celda.Row + 1
        Hasta = wsDestino.Cells(wsDestino.Rows.Count, "P").End(xlUp).Row
        
    Else
        MsgBox "El valor no se encontró en el rango especificado."
        ThisWorkbook.application.ScreenUpdating = True
        Exit Sub
    End If

    ' Copiar el valor de corte en la hoja "Base Trabajo"
    wsDestino.Cells(celda.Row, "P").Copy
    wsB.Range("AG1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ' Eliminar las filas hasta la fecha de corte
    wsDestino.Rows(Desde & ":" & Hasta).Delete Shift:=xlUp

    ' Buscar y eliminar las filas que contengan la palabra "ANULADO" después de eliminar las filas hasta la fecha de corte
    Set rangoBusqueda = wsDestino.Range("A2:AZ" & wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row)
    
    For Each celda In rangoBusqueda.Cells
        If UCase(celda.Value) = "ANULADO" Then
            celda.EntireRow.Delete
        End If
    Next celda

    ' Mostrar el mensaje final de éxito con la fila correcta
    MsgBox "Corte realizado exitosamente a partir de la fila: " & wsDestino.Cells(wsDestino.Rows.Count, "P").End(xlUp).Row, vbInformation, "Corte Exitoso"

    ' Restaurar la actualización de pantalla
    ThisWorkbook.application.ScreenUpdating = True
End Sub

Private Sub CommandButton6_Click()
    Dim cmt As Comment

    Unload UserForm1
    'Call EnviarCorreo

    Set ws = ThisWorkbook.Sheets("Grids")
    Set wsD = ThisWorkbook.Sheets("Superados")
    
    ' Recorrer todos los comentarios en la hoja y ocultarlos
    For Each cmt In ws.Comments
        cmt.Visible = False
    Next cmt
    
    For Each cmt In wsD.Comments
        cmt.Visible = False
    Next cmt
    
 End Sub
 
Sub EnviarCorreo()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim correoDestino As String
    Dim asuntoCorreo As String
    Dim cuerpoCorreo As String
    Dim mesActual As String
    Dim item As MSComctlLib.ListItem
    Dim i As Integer
    Dim header As ColumnHeader
    Dim wsB As Worksheet
    
    Set wsB = ThisWorkbook.Sheets("Base Trabajo")
    
    ' Obtener el nombre del mes actual
    mesActual = Format(DateAdd("m", -1, Date), "mmmm")


    ' Dirección de correo electrónico del destinatario
    correoDestino = "joaquin.o.sanchez@enel.com"
    
    ' Asunto del correo
    asuntoCorreo = "Altas y Bajas del mes " & mesActual
    
    ' Inicializar el cuerpo del correo con HTML y CSS
    cuerpoCorreo = "<html><head><style>" & _
                   "body { font-family: Arial, sans-serif; line-height: 1.6; }" & _
                   "h2 { color: #333333; animation: fadeIn 2s; }" & _
                   "table { width: auto; border-collapse: collapse; }" & _
                   "th { background-color: #4CAF50; color: white; padding: 8px; text-align: center; }" & _
                   "td { padding: 8px; text-align: center; border-bottom: 1px solid #ddd; border-left: 1px solid #ddd; border-right:1px solid #ddd; }" & _
                   "td:last-child { color: red; }" & _
                   "tr:hover { background-color: #f5f5f5; }" & _
                   "th, td { white-space: nowrap; }" & _
                   "</style></head><body>" & _
                   "<p>Hola Joaquin,</p>" & _
                   "<p>Te comparto los contratos que se dieron de baja y los que aparecieron nuevos este mes para que me los confirmes en caso de tener que finalizarlos o dejarlos activos. La fecha de corte es al " & wsB.Range("AG1").Value & " inclusive.</p>"
    
    ' Iniciar la tabla de contratos nuevos
    cuerpoCorreo = cuerpoCorreo & "<h2>Contratos Nuevos:</h2><table><tr>"
    
    ' Agregar encabezados para ListView1
    For Each header In Me.ListView1.ColumnHeaders
        cuerpoCorreo = cuerpoCorreo & "<th style='width:" & header.Width & "px;'>" & header.Text & "</th>"
    Next header
    cuerpoCorreo = cuerpoCorreo & "<th style='width:200px;'>Observaciones</th></tr>" ' Agregar columna "Observaciones" con ancho fijo
    
    ' Construir el texto de las filas del ListView1 (contratos Nuevos)
    For Each item In Me.ListView1.ListItems
        cuerpoCorreo = cuerpoCorreo & "<tr><td>" & item.Text & "</td>" ' Primera columna
        ' Recorrer las subcolumnas del ListView1
        For i = 1 To item.ListSubItems.Count
            cuerpoCorreo = cuerpoCorreo & "<td>" & item.ListSubItems(i).Text & "</td>"
        Next i
        cuerpoCorreo = cuerpoCorreo & "<td></td></tr>" ' Celda vacía para "Observaciones"
    Next item
    
    ' Cerrar la tabla de contratos nuevos
    cuerpoCorreo = cuerpoCorreo & "</table><br>"

    ' Iniciar la tabla de contratos vencidos
    cuerpoCorreo = cuerpoCorreo & "<h2>Contratos Vencidos:</h2><table><tr>"
    
    ' Agregar encabezados para ListView2
    For Each header In Me.ListView2.ColumnHeaders
        cuerpoCorreo = cuerpoCorreo & "<th style='width:" & header.Width & "px;'>" & header.Text & "</th>"
    Next header
    cuerpoCorreo = cuerpoCorreo & "<th style='width:200px;'>Observaciones</th></tr>" ' Agregar columna "Observaciones" con ancho fijo
    
    ' Construir el texto de las filas del ListView2 (contratos Vencidos)
    For Each item In Me.ListView2.ListItems
        cuerpoCorreo = cuerpoCorreo & "<tr><td>" & item.Text & "</td>" ' Primera columna
        ' Recorrer las subcolumnas del ListView2
        For i = 1 To item.ListSubItems.Count
            cuerpoCorreo = cuerpoCorreo & "<td>" & item.ListSubItems(i).Text & "</td>"
        Next i
        cuerpoCorreo = cuerpoCorreo & "<td></td></tr>" ' Celda vacía para "Observaciones"
    Next item
    
    ' Cerrar la tabla de contratos vencidos
    cuerpoCorreo = cuerpoCorreo & "</table><br>"
    
    ' Añadir un cierre al cuerpo del correo
    cuerpoCorreo = cuerpoCorreo & "<p>Saludos,<br>Hernan Carrizo</p></body></html>"

    ' Crear una instancia de Outlook y un nuevo correo electrónico
    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    If OutApp Is Nothing Then Set OutApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    
    If OutApp Is Nothing Then
        MsgBox "No se pudo iniciar Outlook.", vbExclamation
        Exit Sub
    End If
    
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .To = correoDestino
        .Subject = asuntoCorreo
        .HTMLBody = cuerpoCorreo ' Usar HTMLBody para correos con formato HTML
        .Display ' Cambia a .Send para enviar automáticamente
    End With
    
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub


Sub ActualizarGrids(ws As Worksheet, wsZ As Worksheet, nombreArchivo As String, libro As Workbook)
    Dim rangoBusqueda As Range
    Dim rangoBusquedaZ As Range
    Dim celda As Range
    Dim valorBuscado As Variant
    Dim i As Integer
    Dim j As Integer
    Dim ultimaFila As Long
    Dim ultimaFilaZ As Long
    Dim ultimaFilaD As Long
    Dim b As Integer
    Dim a As Integer
    Dim ListItem As ListItem
    
    Set ws = ThisWorkbook.Sheets("Grids")
    Set wsD = ThisWorkbook.Sheets("Superados")
    
    ' Inicializar contador de filas
    a = 0
    b = 1
    
    With Me.ListView1
    .View = lvwReport
    .Gridlines = True
    .FullRowSelect = True
    .HideSelection = False
    .ColumnHeaders.Clear
    .ListItems.Clear

    End With
    
   ' Configurar el ListView2
    With Me.ListView2
    .View = lvwReport
    .Gridlines = True
    .FullRowSelect = True
    .HideSelection = False
    .ColumnHeaders.Clear
    .ListItems.Clear
    
    End With

    On Error Resume Next

    Windows(nombreArchivo).Activate
    ActiveWorkbook.Worksheets("Grids").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Grids").AutoFilter.Sort.SortFields.Add2 Key:=Range _
        ("C1:C" & wsZ.Cells(wsZ.Rows.Count, "C").End(xlUp).Row), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Grids").AutoFilter.Sort
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    Workbooks(macro).Worksheets("Grids").Activate
    
    ultimaFila = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    ultimaFilaD = wsD.Cells(wsD.Rows.Count, "B").End(xlUp).Row
    
    ws.Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Range("C1").Value = "Estado"

    wsD.Rows(1).Value = ws.Rows(1).Value
    wsD.Rows(1).Font.Bold = True
    
        ' Configurar barra de progreso
    Me.ProgressBar1.Visible = True
    Me.ProgressBar1.Min = 1
    Me.ProgressBar1.Max = 100
    Me.ProgressBar1.Value = 1
    
'DEPURACION DE CONTRATOS VENCIDOS O Q NO CORRESPONDEN A GRIDS

    ws.Range("AL2").FormulaR1C1 = "=+VLOOKUP(RC[-36],Superados!C[-36],1,0)"
    ws.Range("AL2").AutoFill Destination:=Range("AL2:AL" & ultimaFila)
    ws.Range("AL1").Value = "Cal_Aux1"
      
  ' Recorrer cada celda en el rango de búsqueda desde la última fila hacia la primera
    For i = ws.Cells(ws.Rows.Count, "AL").End(xlUp).Row To 2 Step -1
        valorBuscado = ws.Cells(i, "AL").Value
        
        Me.ProgressBar1.Value = 100 - (i / ultimaFila) * 100
        DoEvents  ' Actualiza la barra de progreso en tiempo real
        
        ' Si el valor no es un error, eliminar la fila
        If Not IsError(valorBuscado) Then
            For j = wsD.Cells(wsD.Rows.Count, "B").End(xlUp).Row To 2 Step -1
            If LCase(wsD.Cells(j, "B").Value) Like LCase(valorBuscado) Then
                If wsD.Cells(j, "C").Value = "NO" Then
                    ws.Rows(i).Delete
                    Exit For
                End If
                If CDate(ws.Cells(i, "Q").Value) Like CDate(wsD.Cells(j, "Q").Value) Then
                    ws.Rows(i).Delete
                End If
            End If
            Next j
        End If
    Next i

    Me.ProgressBar1.Value = 1
    ultimaFila = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

'BUSQUEDA DE CONTRATOS NUEVOS

    ' Insertar columna y aplicar fórmula VLOOKUP
    ws.Columns("AM:AM").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Range("AM2").FormulaR1C1 = "=VLOOKUP(RC[-37],'[" & nombreArchivo & "]Grids'!C[-37]:C[-36],2,0)"
    ws.Range("AM2").AutoFill Destination:=ws.Range("AM2:AM" & ultimaFila), Type:=xlFillDefault
    ws.Range("AM1").Value = "Cal_Aux2"
    
    ' Definir el rango de resultados de la búsqueda
    Set rangoBusqueda = ws.Range("AM2:AM" & ultimaFila)
       
    ' Recorrer cada celda en el rango de búsqueda
    For Each celda In rangoBusqueda.Cells
        valorBuscado = celda.Value
            ' Actualizar barra de progreso
        Me.ProgressBar1.Value = (celda.Row / rangoBusqueda.Rows.Count) * 100
        DoEvents  ' Actualiza la barra de progreso en tiempo real
        
        If IsError(valorBuscado) Then
            ws.Cells(celda.Row, 3).Value = "NUEVO"
            ws.Cells(celda.Row, 3).Interior.Color = RGB(255, 255, 0)
            ' Actualizar la lista con los contratos nuevos en UserForm1.ListView1

        Else
            ws.Cells(celda.Row, 3).Value = "ACTIVO"
            ws.Cells(celda.Row, 3).Interior.Color = RGB(0, 255, 0)
        
        End If
    Next celda

    Me.ProgressBar1.Value = 1
    
'BUSQUEDA DE CONTRATOS VENCIDOS

    'a = 1
    ' Definir el rango de resultados de la búsqueda

    ultimaFila = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    ultimaFilaZ = wsZ.Cells(wsZ.Rows.Count, "C").End(xlUp).Row
    Set rangoBusquedaZ = wsZ.Range("C2:C" & ultimaFilaZ)
   
    ws.Columns("AM:AM").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Range("AM2").FormulaR1C1 = "=VLOOKUP('[" & nombreArchivo & "]Grids'!RC[-37],R2C2:R" & ultimaFila & "C2,1,FALSE)"
    ws.Range("AM2").AutoFill Destination:=ws.Range("AM2:AM" & ultimaFilaZ), Type:=xlFillDefault
    ws.Range("AM1").Value = "Cal_Aux3"
    
    'Definir el rango de resultados de la búsqueda
    ultimaFila = ws.Cells(ws.Rows.Count, "AM").End(xlUp).Row
    Set rangoBusqueda = ws.Range("AM2:AM" & ultimaFila)
    
For Each celda In rangoBusqueda.Cells
        valorBuscado = celda.Value
        
        Me.ProgressBar1.Value = (celda.Row / rangoBusqueda.Rows.Count) * 100
        DoEvents  ' Actualiza la barra de progreso en tiempo real

            If IsError(valorBuscado) Then
                ultimaFila = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
                wsZ.Rows(celda.Row).Copy Destination:=ws.Rows(ultimaFila + 1)
                ws.Cells(ultimaFila + 1, 3).Interior.Color = RGB(255, 0, 0)
                ws.Cells(ultimaFila + 1, 3).Value = "VENCIDO"
            End If
        'a = a + 1 05621-1-70
Next celda

'DEPURACION YA VENCIDOS
    
 ' Encuentra la última fila con datos en la columna AL
    ultimaFila = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ws.Range("AL2").FormulaR1C1 = "=+VLOOKUP(RC[-36],Superados!C[-36],1,0)"
    ws.Range("AL2").AutoFill Destination:=Range("AL2:AL" & ultimaFila)
    ws.Range("AL1").Value = "Cal_Aux4"

  ' Recorrer cada celda en el rango de búsqueda desde la última fila hacia la primera
    For i = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row To 2 Step -1
        valorBuscado = ws.Cells(i, "AL").Value
        
        Me.ProgressBar1.Value = 100 - (i / ultimaFila) * 100
        DoEvents  ' Actualiza la barra de progreso en tiempo real
        
        ' Si el valor no es un error, eliminar la fila
    If Not IsError(valorBuscado) Then
        For j = wsD.Cells(wsD.Rows.Count, "B").End(xlUp).Row To 2 Step -1
            If LCase(wsD.Cells(j, "B").Value) Like LCase(valorBuscado) Then
                If wsD.Cells(j, "C").Value = "NO" Then
                    ws.Rows(i).Delete
                    Exit For
                End If
                If CDate(ws.Cells(i, "Q").Value) Like CDate(wsD.Cells(j, "Q").Value) Then
                    ws.Rows(i).Delete
                End If
            End If
        Next j
    End If
    Next i

    Me.ProgressBar1.Visible = False

    ' Cerrar el libro sin guardar cambios
    libro.Close SaveChanges:=False
    
    Me.Frame1.Caption = "Contratos Nuevos: " & Me.ListView1.ListItems.Count
    Call Cont_Act
    Me.Frame2.Caption = "Contratos Vencidos: " & Me.ListView2.ListItems.Count
    Call Cont_Venc
    
    On Error GoTo 0
    
On Error Resume Next
        Cells.Select
        Selection.AutoFilter
        ws.AutoFilter.Sort.SortFields.Clear
        ws.AutoFilter.Sort.SortFields.Add2 Key:=Range _
        ("C1:C" & ws.Cells(ws.Rows.Count, "C").End(xlUp).Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
        
    With ws.AutoFilter.Sort
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
On Error GoTo 0

ws.Range("AL:AN").Delete

End Sub


Private Sub CommandButton7_Click()
    ' Verificar si se seleccionó un archivo
    If rutaArchivo = "Falso" Or rutaArchivo = "" Then
        MsgBox "No se seleccionó ningún archivo.", vbExclamation
        Exit Sub
    End If
    
            ' Verificar si ya se realizo un gestión
    If ws.Cells(1, 3).Value = "Estado" Then
        MsgBox "Realice un nuevo Corte previamente", vbExclamation
        Exit Sub
    End If
        
    ' Verificar si se seleccionó una fecha de corte
    If ComboBox1.Value = "" Then
        MsgBox "No se seleccionó una fecha de corte.", vbExclamation, "Error Búsqueda"
        Exit Sub
    End If
    
    ' Llamar a la subrutina ActualizarGrids2
    Call ActualizarGrids(ws, wsZ, nombreArchivo, libro)
End Sub

Private Sub CommandButton8_Click()
    Dim wsZ57 As Worksheet
    Dim ws57 As Worksheet
    Dim wsB As Worksheet
    Dim wsDestino As Worksheet
    Dim nuevoLibro As Workbook
    Dim nombreArchivo As String
    Dim nombreHoja As String
    Dim hoja As Worksheet
    Dim Eleccion As Integer
    Dim Eleccion2 As Integer
    Dim lastRow As Long
    Dim rangoOrigen As Range
    Dim rutaArchivo As Variant

        
Eleccion = MsgBox("Desea cargar un archivo existente?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestión de Contratos")
If Eleccion = vbNo Then

    Call DescargaSAPzm57
        
    Eleccion2 = MsgBox("Desea guardar la descarga zm57", vbQuestion + vbYesNo + vbDefaultButton2, "Gestión de Contratos")
   
    If Eleccion2 = vbYes Then
        ' Establecer la hoja que deseas exportar
        Set wsZ57 = ThisWorkbook.Sheets(5)
        ' Crear un nuevo libro de trabajo
        Set nuevoLibro = Workbooks.Add
        ' Copiar la hoja al nuevo libro de trabajo
        wsZ57.Copy Before:=nuevoLibro.Sheets(1)
        
            ' Eliminar todas las hojas adicionales en el nuevo libro
            For Each hoja In nuevoLibro.Sheets
                If hoja.Name <> wsZ57.Name Then
                    hoja.Delete
                End If
            Next hoja
        
        ' Confirmación
        MsgBox "Por favor, verifique y guarde.", vbInformation, "Exportación Completa"
    End If
End If

If Eleccion = vbYes Then
        Set wsB = ThisWorkbook.Sheets("Base Trabajo")
    
        Set wsDestino = ThisWorkbook.Sheets(5)
        MsgBox "Por favor, seleccione el archivo ZM-57", vbInformation, "Actualización de Montos"
        ' Solicitar al usuario que seleccione el archivo
        rutaArchivo = ThisWorkbook.application.GetOpenFilename("Archivos de Excel (*.xls; *.xlsx ; *.xlsm), *.xls; *.xlsx; *.xlsm", , "Seleccione el archivo de Excel")
        
        If rutaArchivo = False Then
            MsgBox "No se seleccionó ningún archivo.", vbExclamation
            Exit Sub
        End If
              
        ' Abrir el archivo seleccionado
        Set libro = Workbooks.Open(rutaArchivo)
        ' Configurar hojas de trabajo
        Set ws57 = libro.Sheets(1)
        
        ' Verificar y mover la columna antes de copiar los datos
        If ws57.Cells(1, 16).Value = "Documento compras" Then
            ws57.Columns("P:P").Cut
            ws57.Columns("A:A").Insert Shift:=xlToRight
        End If

        ' Definir el rango de origen después de mover la columna
        Set rangoOrigen = ws57.Range("A:AJ")
        rangoOrigen.Copy Destination:=wsDestino.Range("A1")
        wsDestino.Name = "zm57"
 End If
 
    ' Establecer la hoja de trabajo activa o especificar la hoja si es necesario
    Set wsB = ThisWorkbook.Sheets("Base Trabajo")
    
    ' Encontrar la última fila con datos en la columna A
    lastRow = wsB.Cells(wsB.Rows.Count, "A").End(xlUp).Row
    
    ' Insertar la fórmula en la celdaa
    'wsB.Range("O2").formula = "=VLOOKUP(A2;'zm50 04-11-24'!$B:$Q;16;FALSE)"
    wsB.Range("N2").formula = "=VLOOKUP(A2,'zm57'!$A:$R,5,FALSE)"
    wsB.Range("P2").formula = "=VLOOKUP(A2,'zm57'!$A:$R,8,FALSE)"
    wsB.Range("Q2").formula = "=VLOOKUP(A2,'zm57'!$A:$R,10,FALSE)"
    
    ' Autocompletar la fórmula hasta la última fila
    'wsB.Range("O2").AutoFill Destination:=wsB.Range("O2:O" & lastRow)
    wsB.Range("N2").AutoFill Destination:=wsB.Range("N2:N" & lastRow)
    wsB.Range("P2").AutoFill Destination:=wsB.Range("P2:P" & lastRow)
    wsB.Range("Q2").AutoFill Destination:=wsB.Range("Q2:Q" & lastRow)
    
    Dim rng As Object
    Set rng = wsB.Range("A2:AC" & lastRow)
    
    ' Aplica bordes alrededor de todo el rango
    With rng.Borders
        .LineStyle = xlContinuous       ' Estilo de línea continua
        .Color = RGB(0, 0, 0)           ' Color negro (puedes cambiarlo)
        .Weight = xlThin                ' Grosor de la línea (puedes usar xlMedium, xlThick, etc.)
    End With
    
    Call ActualizarCoti
        
End Sub

Private Sub CommandButton13_Click()
    Call ActualizarCoti
End Sub

Sub ActualizarCoti()
    Dim wsB As Worksheet
    Dim Eleccion As Integer
    
    Set wsB = ThisWorkbook.Sheets("Base Trabajo")

    Eleccion = MsgBox("Desea actualizar cotizaciones?", vbQuestion + vbYesNo + vbDefaultButton2, "Actualización de Montos")
    
    If Eleccion = vbYes Then
        wsB.Range("AD1:AE1").Copy
        wsB.Range("AD3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End If

End Sub

Sub DescargaSAPzm57()
    On Error GoTo ErrorHandler
    Dim wsB As Worksheet
    Dim ws57 As Worksheet
    Dim processComplete As Boolean
    Dim attemptCounter As Integer
    Dim newWindowId As String
    Dim originalWindowId As String
    Dim lastRow As Long
    
    ' Acceder a la hoja de Excel y la columna deseada
    Set wsB = ThisWorkbook.Sheets("Base Trabajo")
    lastRow = wsB.Cells(wsB.Rows.Count, "A").End(xlUp).Row

    Call IniciarSAP
    
        ' Verificar si se pudo establecer la conexión
    If connection Is Nothing Then
        Set connection = Nothing
        Set session = Nothing
        Exit Sub
    End If
        
    ' Ejecutar la transacción
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "zm57"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[0]").press
 
    ' Verificación de errores y esperar
    If Not IsObject(session.findById("wnd[0]/usr/ctxtS_KDATE-LOW")) Then
        MsgBox "Error: No se pudo encontrar el campo S_KDATE-LOW.", vbCritical
        Exit Sub
    End If
    session.findById("wnd[0]/usr/ctxtS_KDATE-LOW").Text = ""
    session.findById("wnd[0]/usr/btn%_S_EBELN_%_APP_%-VALU_PUSH").press
    
    wsB.Range("A2:A" & lastRow).Copy
    session.findById("wnd[1]/tbar[0]/btn[24]").press ' Pegar en SAP
    
    ' Presionar botón para continuar con el proceso
    session.findById("wnd[1]/tbar[0]/btn[8]").press

    ' Esperar hasta que el proceso se complete
    originalWindowId = session.findById("wnd[0]").ID
    processComplete = False
    attemptCounter = 0
    Do While Not processComplete And attemptCounter < 10
        On Error Resume Next
        newWindowId = session.findById("wnd[0]").ID
        If newWindowId <> originalWindowId Then
            processComplete = True
        End If
        attemptCounter = attemptCounter + 1
        ThisWorkbook.application.Wait Now + TimeValue("00:00:01") ' Espera de 1 segundo entre intentos
        On Error GoTo ErrorHandler
    Loop
        
    ThisWorkbook.application.Wait Now + TimeValue("00:00:05") ' Espera de 5 segundos
    
    ' Finalizar el proceso
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[43]").press ' Abre documento Excel con la info procesada

    Set ws57 = ActiveWorkbook.Sheets(1)
       
    ' Verificar y mover la columna antes de copiar los datos
    If ws57.Cells(1, 16).Value = "Documento compras" Then
        ws57.Columns("P:P").Cut
        ws57.Columns("A:A").Insert Shift:=xlToRight
    End If
    
        ' Copiar la hoja al nuevo libro de trabajo
    ws57.Copy Before:=Workbooks(macro).Sheets("zm57")
    ThisWorkbook.application.DisplayAlerts = False
    Workbooks(macro).Sheets(6).Delete
    ThisWorkbook.application.DisplayAlerts = True
    Workbooks(macro).Sheets(5).Name = "zm57"
    
    With Workbooks(macro).Sheets(6).Tab
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
    End With
         
    If Workbooks(macro).Sheets(7).Cells(1, 16).Value = "Documento Compras" Then
        Workbooks(macro).Sheets(7).Columns("P:P").Cut
        Workbooks(macro).Sheets(7).Columns("A:A").Insert Shift:=xlToRight
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[15]").press 'Cierra la ventana con la info procesada y el excel abierto
Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

Private Sub CommandButton9_Click()
    Dim ws As Worksheet
    Dim nuevoLibro As Workbook
    Dim hoja As Worksheet
    Dim Eleccion As Integer
    Dim rutaArchivo As Variant
    
    Eleccion = MsgBox("Confirma guardar la nueva Gestión realizada?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestión de Contratos")
    If Eleccion <> vbYes Then Exit Sub
    
    ' Establecer la hoja que deseas exportar
    Set ws = ThisWorkbook.Sheets("Grids")
    
    ' Copiar la hoja a un nuevo libro (esto automáticamente crea un nuevo libro)
    ws.Copy
    Set nuevoLibro = ActiveWorkbook
    
    ' Eliminar todas las hojas adicionales en el nuevo libro
    ThisWorkbook.application.DisplayAlerts = False
    For Each hoja In nuevoLibro.Sheets
        If hoja.Name <> ws.Name Then
            hoja.Delete
        End If
    Next hoja
    ThisWorkbook.application.DisplayAlerts = True
    
    ' Abrir la ventana para seleccionar la ubicación y el nombre del archivo
    rutaArchivo = ThisWorkbook.application.GetSaveAsFilename( _
        InitialFileName:="Grids.xlsx", _
        FileFilter:="Archivos de Excel (*.xlsx), *.xlsx", _
        Title:="Guardar como")
    
    If rutaArchivo = False Then
        MsgBox "Operación cancelada. El archivo no fue guardado.", vbExclamation, "Cancelado"
        nuevoLibro.Close SaveChanges:=False
        Exit Sub
    End If
    
    If Dir(rutaArchivo) <> "" Then
        Dim sobreescribir As VbMsgBoxResult
        sobreescribir = MsgBox("El archivo ya existe. ¿Desea sobrescribirlo?", vbExclamation + vbYesNo, "Archivo existente")
        If sobreescribir <> vbYes Then
            MsgBox "No se guardó el archivo.", vbInformation, "Cancelado"
            nuevoLibro.Close SaveChanges:=False
            Exit Sub
        Else
            On Error Resume Next
            Kill rutaArchivo
            On Error GoTo 0
        End If
    End If
    
    ' Guardar el nuevo libro
    On Error Resume Next
    nuevoLibro.SaveAs Filename:=rutaArchivo, FileFormat:=xlOpenXMLWorkbook
    nuevoLibro.Close SaveChanges:=False
    On Error GoTo 0
    
    MsgBox "Archivo guardado correctamente.", vbInformation, "Éxito"
End Sub


Private Sub Image2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim coleccion As Collection
    Dim item As MSComctlLib.ListItem
    Dim index As Variant
    Set coleccion = New Collection
    Dim Election As Integer
    
    On Error Resume Next
    If Me.ListView2.selectedItem.Selected = False Then
        MsgBox "Por favor, seleccioná varios Contratos.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Contar los ítems seleccionados
    For Each item In Me.ListView2.ListItems
        If item.Selected Then
            coleccion.Add item.index
        End If
    Next item
    
If coleccion.Count = 1 Then
            MsgBox "Por favor, seleccioná varios Contratos.", vbExclamation
        Exit Sub
    Else
    
If Me.CheckBox1.Value = False Then
        Election = MsgBox("Desea Finalizar los Contratos?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestion de Contrato")
        
    If Election = vbYes Then
        For Each index In coleccion
            Bottom = True
            Me.ListView2.ListItems(index).Selected = True
            Call Fin_Cont(Me.ListView2.ListItems(ListView2.selectedItem.index).SubItems(1))
        Next index
        
    ElseIf Election = vbNo Then
        Election = MsgBox("¿Desea Mantenerlos Activos?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestion de Contrato")
        
        If Election = vbYes Then
            For Each index In coleccion
                Bottom = True
                Me.ListView2.ListItems(index).Selected = True
                Call Hold_Cont_Act(Me.ListView2.ListItems(ListView2.selectedItem.index).SubItems(1))
            Next index
            
            Else
                Exit Sub
        End If
    End If
End If

If Me.CheckBox1.Value = True Then
        Election = MsgBox("¿Desea Reactivar los Contratos?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestion de Contrato")
        
    If Election = vbYes Then
        For Each index In coleccion
            Bottom = True
            Me.ListView2.ListItems(index).Selected = True
            Call React_cont(Me.ListView2.ListItems(ListView2.selectedItem.index).SubItems(1))
        Next index
        
    ElseIf Election = vbNo Then
        Exit Sub
    End If
End If

End If

    Call Cont_Venc
    Call Cont_Act
    Bottom = False
    
End Sub


Private Sub Image4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim coleccion As Collection
Dim item As MSComctlLib.ListItem
Dim index As Variant
Set coleccion = New Collection
Dim Election As Integer
    
On Error Resume Next
If Me.ListView1.selectedItem.Selected = False Then
    MsgBox "Por favor, seleccioná varios Contratos.", vbExclamation
    Exit Sub
End If
On Error GoTo 0
    
' Contar los ítems seleccionados
For Each item In Me.ListView1.ListItems
    If item.Selected Then
        coleccion.Add item.index
    End If
Next item
    
If coleccion.Count = 1 Then
            MsgBox "Por favor, seleccioná varios Contratos.", vbExclamation
        Exit Sub
Else
    
If Me.CheckBox2.Value = False Then
    Election = MsgBox("¿Desea Activar los Contratos?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestion de Contrato")
    
        If Election = vbYes Then
        
        For Each index In coleccion
 
            Me.ListView1.ListItems(index).Selected = True
            Call Acti_Cont(Me.ListView1.ListItems(ListView1.selectedItem.index).SubItems(1))
        Next index

        Else
            Election = MsgBox("¿Desea Descartar los Contratos?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestion de Contrato")
        
        If Election = vbYes Then
        
         For Each index In coleccion

             Me.ListView1.ListItems(index).Selected = True
             Call Desc_Cont(Me.ListView1.ListItems(ListView1.selectedItem.index).SubItems(1))
         Next index
             
         Else
             Exit Sub
         End If
         End If
End If
    
If Me.CheckBox2.Value = True Then
    Election = MsgBox("¿Desea Finalizar los Contratos?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestion de Contrato")

        ' Si el usuario selecciona 'Sí', ejecutar SeleccionarArchivo
        If Election = vbYes Then
        
        For Each index In coleccion

            Me.ListView1.ListItems(index).Selected = True
            Call Fin_Cont(Me.ListView1.ListItems(ListView1.selectedItem.index).SubItems(1))
        Next index
                   
        Else
            Exit Sub
        End If
End If
End If
    
Call Cont_Act
Call Cont_Venc

End Sub

Private Sub ListView1_DblClick()
    Dim Election As Integer
    Dim wsDestino As Worksheet
    Dim wsB As Worksheet
    Dim nextRow As Long
    Dim i As Integer
    Dim formula As String
    Dim selectedItem As MSComctlLib.ListItem
    Dim selectedIndex As Integer

    ' Definir la hoja de destino
    Set wsDestino = ThisWorkbook.Sheets("Superados")
    Set ws = ThisWorkbook.Sheets("Grids")
    Set wsB = Workbooks(macro).Worksheets("Base Trabajo")

    If Me.ListView1.selectedItem Is Nothing Then
        Exit Sub
    End If
    
    'Obtener el ítem seleccionado
    selectedIndex = Me.ListView1.selectedItem.index
    numCont = Me.ListView1.selectedItem.SubItems(1) ' Suponiendo que el número de contrato está en la columna 2 (SubItem 1)


If Me.CheckBox2.Value = False Then
    Election = MsgBox("¿Desea Activar el Contrato " & numCont & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestion de Contrato")
    
        If Election = vbYes Then
            Call Acti_Cont(numCont)

        Else
            Election = MsgBox("¿Desea Descartar el Contrato" & numCont & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestion de Contrato")
        
           If Election = vbYes Then
                Call Desc_Cont(numCont)
            Else
                Exit Sub
            End If
        End If
End If
    
If Me.CheckBox2.Value = True Then
        Election = MsgBox("¿Desea Finalizar el Contrato " & numCont & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestion de Contrato")

        ' Si el usuario selecciona 'Sí', ejecutar SeleccionarArchivo
        If Election = vbYes Then
            Call Fin_Cont(numCont)
        Else
            Exit Sub
        End If
End If
    
Call Cont_Act
Call Cont_Venc
'Call Act_List1

End Sub

Private Sub Acti_Cont(ByVal numCont As Variant)
    Dim wsDestino As Worksheet
    Dim wsB As Worksheet
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim i As Long, j As Long
    Dim macro As String
    Dim formula As String
    Dim encontrado As Boolean

    ' Definir las hojas
    Set wsDestino = ThisWorkbook.Sheets("Superados")
    Set ws = ThisWorkbook.Sheets("Grids")
    Set wsB = ThisWorkbook.Sheets("Base Trabajo")

    ' Activar el contrato en hoja Grids
    For i = 1 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        If LCase(ws.Cells(i, 2).Value) Like LCase(numCont) Then
            ws.Cells(i, 3).Interior.Color = RGB(0, 255, 0)
            ws.Cells(i, 3).Value = "ACTIVO"
            Exit For
        End If
    Next i

    ' Buscar si ya existe en Base Trabajo
    encontrado = False
    For j = 1 To wsB.Cells(wsB.Rows.Count, "A").End(xlUp).Row
        If LCase(wsB.Cells(j, 1).Value) Like LCase(numCont) Then
            wsB.Cells(j, 2).Value = "ACTIVO"
            wsB.Cells(j, 2).Interior.Color = RGB(0, 255, 0)
            MsgBox "El contrato ya existía en la Base de Trabajo en la fila: " & j
            encontrado = True
            Exit For
        End If
    Next j

    ' Si no se encontró, copiar desde Grids
    If Not encontrado Then
        nextRow = wsB.Cells(wsB.Rows.Count, "A").End(xlUp).Row + 1
        ws.Cells(i, 2).Copy Destination:=wsB.Cells(nextRow, 1)
        ws.Cells(i, 3).Copy Destination:=wsB.Cells(nextRow, 2)
        ws.Cells(i, 4).Copy Destination:=wsB.Cells(nextRow, 3)
        ws.Cells(i, 8).Copy Destination:=wsB.Cells(nextRow, 12)

        formula = "=BUSCARV(L" & nextRow & ", Estructura!A:B, 2, 0)"
        wsB.Range("K" & nextRow).formula = formula
    End If

    ' Eliminar de Superados si existe
    For j = 1 To wsDestino.Cells(wsDestino.Rows.Count, "B").End(xlUp).Row
        If LCase(wsDestino.Cells(j, 2).Value) Like LCase(numCont) Then
            wsDestino.Rows(j).Delete
            Exit For
        End If
    Next j
End Sub


Private Sub Desc_Cont(ByVal numCont As Variant)
Dim wsDestino As Worksheet
Dim wsB As Worksheet
Dim ws As Worksheet
Dim nextRow As Integer
Dim i As Integer

' Definir la hoja de destino
Set wsDestino = ThisWorkbook.Sheets("Superados")
Set ws = ThisWorkbook.Sheets("Grids")
Set wsB = ThisWorkbook.Sheets("Base Trabajo")

nextRow = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row + 1
    For i = 1 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        If LCase(ws.Cells(i, 2).Value) Like LCase(numCont) Then
            ws.Rows(i).Copy Destination:=wsDestino.Rows(nextRow)
            wsDestino.Cells(nextRow, 3).Interior.Color = RGB(255, 0, 0)
            wsDestino.Cells(nextRow, 3).Value = "NO"
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
                
    For i = 1 To wsB.Cells(wsB.Rows.Count, "B").End(xlUp).Row
        If LCase(wsB.Cells(i, 1).Value) Like LCase(numCont) Then
            wsB.Rows(i).Delete
            Exit For
        End If
    Next i
    
End Sub

Private Sub Fin_Cont(ByVal numCont As Variant)
Dim wsDestino As Worksheet
Dim wsB As Worksheet
Dim ws As Worksheet
Dim nextRow As Integer
Dim i As Integer

' Definir la hoja de destino
Set wsDestino = ThisWorkbook.Sheets("Superados")
Set ws = ThisWorkbook.Sheets("Grids")
Set wsB = ThisWorkbook.Sheets("Base Trabajo")

' Encontrar la siguiente fila vacía en la columna B
nextRow = wsDestino.Cells(wsDestino.Rows.Count, "B").End(xlUp).Row + 1

    For i = 1 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        If LCase(ws.Cells(i, 2).Value) Like LCase(numCont) Then
            ws.Rows(i).Copy Destination:=wsDestino.Rows(nextRow)
            wsDestino.Cells(nextRow, 3).Interior.Color = RGB(255, 0, 0)
            wsDestino.Cells(nextRow, 3).Value = "FINALIZADO"
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
    
    For i = 1 To wsB.Cells(wsB.Rows.Count, "B").End(xlUp).Row
        If LCase(wsB.Cells(i, 1).Value) Like LCase(numCont) Then
            wsB.Cells(i, 2).Value = "FINALIZADO"
            wsB.Cells(i, 2).Interior.Color = RGB(255, 0, 0)
            Exit For
        End If
    Next i
    
End Sub

Private Sub Hold_Cont_Act(ByVal numCont As Variant)
Dim wsB As Worksheet
Dim ws As Worksheet
Dim nextRow As Integer
Dim formula As String
Dim i As Integer
Dim j As Integer
Dim encontrado As Boolean

Set wsB = ThisWorkbook.Sheets("Base Trabajo")
Set ws = ThisWorkbook.Sheets("Grids")

nextRow = wsB.Cells(wsB.Rows.Count, "B").End(xlUp).Row + 1

For i = 1 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    encontrado = False
    
    If LCase(ws.Cells(i, 2).Value) Like LCase(numCont) Then
        ws.Cells(i, 3).Value = "ACTIVO"
        ws.Cells(i, 3).Interior.Color = RGB(255, 150, 0)
        
        For j = 2 To wsB.Cells(wsB.Rows.Count, "A").End(xlUp).Row
            If LCase(wsB.Cells(j, 1).Value) = LCase(numCont) Then
                encontrado = True
                Exit For
            End If
        Next j
        
        If Not encontrado Then
            ws.Cells(i, 2).Copy Destination:=wsB.Cells(nextRow, 1)
            ws.Cells(i, 3).Copy Destination:=wsB.Cells(nextRow, 2)
            ws.Cells(i, 4).Copy Destination:=wsB.Cells(nextRow, 3)
            ws.Cells(i, 8).Copy Destination:=wsB.Cells(nextRow, 12)
            
            formula = "=BUSCARV(L" & nextRow & ", Estructura!A:B, 2, 0)"
            wsB.Range("K" & nextRow).formula = formula
        End If
        
    End If
Next i

End Sub

Private Sub React_cont(ByVal numCont As Variant)
Dim wsDestino As Worksheet
Dim wsB As Worksheet
Dim ws As Worksheet
Dim nextRow As Integer
Dim nextRowB As Integer
Dim i As Integer
Dim encontrado As Boolean
Dim formula As String

Set wsB = ThisWorkbook.Sheets("Base Trabajo")
Set wsDestino = ThisWorkbook.Sheets("Superados")
Set ws = ThisWorkbook.Sheets("Grids")
encontrado = False


For i = 1 To wsDestino.Cells(wsDestino.Rows.Count, "B").End(xlUp).Row
    If LCase(wsDestino.Cells(i, 2).Value) Like LCase(numCont) Then
        ' Encontrar la siguiente fila vacía en la columna A
        nextRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1
        wsDestino.Rows(i).Copy Destination:=ws.Rows(nextRow)
        wsDestino.Rows(i).Delete
        
        ws.Cells(nextRow, 3).Value = "ACTIVO"
        ws.Cells(nextRow, 3).Interior.Color = RGB(0, 255, 0)
        Exit For
    End If
Next i

For i = 1 To wsB.Cells(wsB.Rows.Count, "B").End(xlUp).Row
    If LCase(wsB.Cells(i, 1).Value) Like LCase(numCont) Then
        wsB.Cells(i, 2).Value = "ACTIVO"
        wsB.Cells(i, 2).Interior.Color = RGB(0, 255, 0)
        encontrado = True
        Exit For
    End If
Next i

If Not encontrado Then
For i = 1 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If LCase(ws.Cells(i, 2).Value) Like LCase(numCont) Then
        nextRow = wsB.Cells(wsB.Rows.Count, "B").End(xlUp).Row + 1
        ws.Cells(i, 2).Copy Destination:=wsB.Cells(nextRow, 1)
        ws.Cells(i, 3).Copy Destination:=wsB.Cells(nextRow, 2)
        ws.Cells(i, 4).Copy Destination:=wsB.Cells(nextRow, 3)
        ws.Cells(i, 8).Copy Destination:=wsB.Cells(nextRow, 12)
        
        formula = "=BUSCARV(L" & nextRow & ", Estructura!A:B, 2, 0)"
        wsB.Range("K" & nextRow).formula = formula
        Exit For
    End If
Next i
End If



End Sub

Private Sub ListView2_DblClick()
Dim wsDestino As Worksheet
Dim wsB As Worksheet
Dim ws As Worksheet
Dim Election As Integer
Dim selectedItem As MSComctlLib.ListItem
Dim selectedIndex As Integer
   
    If Me.ListView2.selectedItem Is Nothing Then
        Exit Sub
    End If
    
    If Not Bottom = True Then
        numCont = ListView2.selectedItem.index
    End If
    
    ' Definir la hoja de destino
    Set wsDestino = ThisWorkbook.Sheets("Superados")
    Set ws = ThisWorkbook.Sheets("Grids")
    Set wsB = ThisWorkbook.Sheets("Base Trabajo")

    ' Obtener el ítem seleccionado
    selectedIndex = Me.ListView2.selectedItem.index
    numCont = Me.ListView2.selectedItem.SubItems(1) ' Suponiendo que el número de contrato está en la columna 2 (SubItem 1)
           
If Me.ListView2.selectedItem.Selected Then
    
    If Me.CheckBox1.Value = False Then
        
        If Not Bottom Then
            Election = MsgBox("Desea Finalizar el Contrato " & numCont & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestion de Contrato")
        Else
            Election = vbYes
        End If
        
        If Election = vbYes Then
            Call Fin_Cont(numCont)
        End If

        If Election = vbNo Then
            Election = MsgBox("¿Desea Mantenerlo Activo?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestion de Contrato")
              
        If Election = vbYes Then
            Call Hold_Cont_Act(numCont)
        Else
            Exit Sub
        End If
        End If
    End If

    If Me.CheckBox1.Value = True Then
        
        If Not Bottom Then
            Election = MsgBox("¿Desea Reactivar el Contrato " & numCont & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Gestion de Contrato")
        Else
            Election = vbYes
        End If

        If Election = vbYes Then
           Call React_cont(numCont)
        Else
            Exit Sub
        End If
    End If
End If
    Call Cont_Venc
    Call Cont_Act
End Sub

Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Debug.Print "Tecla presionada: " & KeyCode
    Dim itemFound As Boolean
    Dim i As Integer
    Dim searchValue As String

    ' Verificar si se presionó la tecla Enter (KeyCode = 13)
    If KeyCode = vbKeyReturn Then
        ' Capturar el valor del TextBox2
        searchValue = Trim(Me.TextBox2.Text)
        itemFound = False

        ' Si el TextBox está vacío, salir del sub
        If searchValue = "" Then Exit Sub

        ' Buscar en ListView1
        For i = 1 To Me.ListView1.ListItems.Count
            If LCase(Me.ListView1.ListItems(i).SubItems(1)) = LCase(searchValue) Then
                ' Seleccionar el ítem y salir
                Me.ListView1.ListItems(i).Selected = True
                Me.ListView1.SetFocus
                Me.ListView1.ListItems(i).EnsureVisible ' Asegurarse de que sea visible
                itemFound = True
                Exit For
            End If
        Next i

        ' Si no se encontró en ListView1, buscar en ListView2
        If Not itemFound Then
            For i = 1 To Me.ListView2.ListItems.Count
                If LCase(Me.ListView2.ListItems(i).SubItems(1)) = LCase(searchValue) Then
                    ' Seleccionar el ítem y salir
                    Me.ListView2.ListItems(i).Selected = True
                    Me.ListView2.SetFocus
                    Me.ListView2.ListItems(i).EnsureVisible ' Asegurarse de que sea visible
                    itemFound = True
                    Exit For
                End If
            Next i
        End If

        ' Si no se encontró en ninguno de los ListView
        If Not itemFound Then
            'Debug.Print "Ítem no encontrado en ningún ListView"
            ' Opcional: mostrar un mensaje
            ' MsgBox "Contrato no encontrado", vbInformation, "Búsqueda"
        End If
    End If
End Sub



Private Sub ToggleButton1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Verifica si el botón izquierdo del mouse fue presionado
    If Button = 1 Then
        ' Alternar entre los anchos deseados y cambiar texto del botón
        If Me.Width = 787 Then
            Me.Width = 507 ' Ancho reducido
            ToggleButton1.Caption = "Expandir"
        Else
            Me.Width = 787 ' Ancho expandido
            ToggleButton1.Caption = "Contraer"
        End If
    End If
End Sub


Private Sub UserForm_Activate()
    Dim ws1 As Worksheet
    Dim wb As Object
    Dim HTMLDoc As Object
    Dim item As MSComctlLib.ListItem
    Dim numCont As String
    Dim lastRow As Long

    Me.CommandButton6.Enabled = False
    Me.CommandButton4.Enabled = False
    Me.CommandButton3.Enabled = False
    Me.CommandButton12.Enabled = False
    Me.ToggleButton1.Enabled = False
    Bottom = False
    
    Me.Label16.Caption = Date
    

    ' Deseleccionar todos los ítems en ambos ListView
    Me.ListView2.selectedItem = Nothing
    Me.ListView1.selectedItem = Nothing
    
    Me.ListView1.MultiSelect = True
    Me.ListView2.MultiSelect = True

    ' Suprimir los errores de script
    Me.WebBrowser1.Silent = True
    Me.WebBrowser1.Height = Me.InsideHeight
    Me.ProgressBar1.Visible = False

    ' Extraer el valor de la fila con la fecha de corte establecida
    macro = ActiveWorkbook.Name
    Set ws1 = Workbooks(macro).Worksheets(4)
    Set wsD = Workbooks(macro).Worksheets(2)

    ' Formato de las columnas como fecha y ordenarlas
    With ws1
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row

        .Columns("N:Q").NumberFormat = "dd/mm/yyyy"

        .Sort.SortFields.Clear
        .Sort.SortFields.Add2 Key:=.Range("P2:P" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortTextAsNumbers

        With .Sort
            .SetRange ws1.Range("A1:AJ" & lastRow)
            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With

    ' Establecer la lista de datos en el ComboBox
    Me.ComboBox1.List = ws1.Range("P2:P" & lastRow).Value
    Me.StartUpPosition = 1
    Me.Height = 542
    Me.Width = 507
    rutaArchivo = "Falso"

    Call MostrarCotizacion

    ' Establecer la referencia al control WebBrowser
    Set wb = Me.WebBrowser1
    Set HTMLDoc = CreateObject("HTMLFile")

    ' Manejo de errores para evitar mensajes de error de script
    On Error Resume Next
    wb.Silent = True
    wb.FullScreen = True
    HTMLDoc.parentWindow.execScript "window.onerror = function() { return true; }"

    ' Cargar una página web
    wb.Navigate "https://www.lanacion.com.ar/tema/euro-hoy-tid66142"

    ' Esperar a que la página se cargue
    Do While wb.Busy Or wb.ReadyState <> 4
        DoEvents
        Me.Label14.BackColor = IIf(Me.Label14.BackColor = RGB(255, 255, 255), RGB(0, 0, 0), RGB(255, 255, 255))
    Loop

    ' Rehabilitar manejo de errores
    On Error GoTo 0

    Me.CommandButton6.Enabled = True
    Me.CommandButton4.Enabled = True
    Me.CommandButton3.Enabled = True
    Me.CommandButton12.Enabled = True
    Me.Label14.Visible = False
    Me.ToggleButton1.Enabled = True
    'ThisWorkbook.Sheets(1).Activate

End Sub

Sub MostrarCotizacion()
    Dim pag As Object
    Dim html As Object
    Dim dolarCompra As String
    Dim dolarVenta As String
    Dim euroCompra As String
    Dim euroVenta As String
    
    ' Crear el objeto InternetExplorer
    Set pag = CreateObject("InternetExplorer.Application")
    pag.Visible = False ' Para hacerlo visible cambiar a True
    
    ' Cargar la página web
    pag.Navigate "https://www.lanacion.com.ar/tema/euro-hoy-tid66142/"
    
    ' Esperar a que la página se cargue
    Do While pag.Busy Or pag.ReadyState <> 4
        DoEvents
        If Me.Label14.BackColor = RGB(255, 255, 255) Then
            Me.Label14.BackColor = RGB(0, 0, 0)
        Else
            Me.Label14.BackColor = RGB(255, 255, 255)
        End If
    Loop
    
    ' Obtener el HTML de la página
    Set html = pag.Document
    
    ' Buscar y extraer los valores
    ExtractCurrencyValues html, "Dólar oficial", dolarCompra, dolarVenta
    ExtractCurrencyValues html, "Euro", euroCompra, euroVenta
    
    ' Cerrar el objeto InternetExplorer
    pag.Quit
    Set pag = Nothing
    
    Dim promedioDolar As Double
    Dim promedioEuro As Double
    Dim DolarEuro As Double
    promedioDolar = (Val(dolarVenta) + Val(dolarCompra)) / 2
    promedioEuro = (Val(euroVenta) + Val(euroCompra)) / 2
    DolarEuro = (promedioEuro / promedioDolar)
    
    Me.Label6.Caption = "Fuente: https://www.lanacion.com.ar/tema/euro-hoy-tid66142/"
    Me.Label7.Caption = "Compra: " & "$" & dolarCompra & " Venta: " & "$" & dolarVenta
    Me.Label8.Caption = "Compra: " & "$" & euroCompra & " Venta: " & "$" & euroVenta
    Workbooks(macro).Worksheets("Base Trabajo").Range("AD1") = promedioEuro
    Workbooks(macro).Worksheets("Base Trabajo").Range("AE1") = DolarEuro
    
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

