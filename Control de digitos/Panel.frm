VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Panel 
   Caption         =   "PANEL"
   ClientHeight    =   5412
   ClientLeft      =   12
   ClientTop       =   84
   ClientWidth     =   7848
   OleObjectBlob   =   "Panel.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Call ZCO9_certificados
End Sub

Private Sub CommandButton3_Click()
    Me.CommandButton1.Enabled = False
    Me.CommandButton3.Enabled = False
    Me.CommandButton4.Enabled = False
    
    Call VerResumen
    Call LlenarListView
    
    Me.CommandButton1.Enabled = True
    Me.CommandButton3.Enabled = True
    Me.CommandButton4.Enabled = True
    
    If numCont <> "" Then
        MsgBox "El Panel fue actualizado exitosamente.", vbInformation
    End If
    
End Sub

Private Sub VerResumen()
    Dim ht1 As Worksheet
    Dim ht2 As Worksheet
    Dim rng As Range
    Dim celda As Range
    Dim dict As Object
    Dim Valor As Variant
    Dim ultimoValor As Long
    Dim lastRow As Long
    Dim i As Long
    Dim numCont As String

    ' Asignar hojas
    Set ht2 = ThisWorkbook.Sheets("Contratos")
    Set ht1 = ThisWorkbook.Sheets("Imputaciones")
    Set rng = ht1.Range("A5:D" & ht1.Cells(ht1.Rows.Count, "A").End(xlUp).Row)
    
    ' Aplica bordes alrededor de todo el rango
    With rng.Borders
        .LineStyle = xlNone
    End With
    
    If ht1.AutoFilterMode Then ht1.AutoFilterMode = False
      
    'ht1.Sort.SortFields.Clear ' Limpiar criterios de ordenamiento
    
    With ht1
        .Range("B2").ClearContents
        .Range("A5:A" & ht1.Rows.Count).ClearContents
        .Range("B5:B" & ht1.Rows.Count).ClearContents
        .Range("C5:C" & ht1.Rows.Count).ClearContents
        .Range("D5:D" & ht1.Rows.Count).ClearContents
        .Cells(2, 2).value = Panel.ComboBox1.value
    End With

    ' Asegúrate de que el ComboBox tiene un valor seleccionado
    If Panel.ComboBox1.ListIndex <> -1 Then
        ' Asigna el valor seleccionado a la variable numCont
        numCont = Trim(Panel.ComboBox1.value)
    End If

    ' Comprobar que la hoja existe antes de continuar
    On Error Resume Next
    Set ht2 = ThisWorkbook.Sheets(numCont)
    On Error GoTo 0 ' Restablecer manejo de errores

    If ht2 Is Nothing Then
        MsgBox "La hoja '" & numCont & "' no existe.", vbExclamation
        Exit Sub
    End If
    
Call concatserpos(numCont)

    ht1.Activate
    ht1.Range("A4:D4").AutoFilter

End Sub

Sub concatserpos(numCont As String)
    Dim ht3 As Worksheet
    Dim ht1 As Worksheet
    Dim celdaC As Range
    Dim celdaD As Range
    Dim celdaE As Range
    Dim concatenacion As String
    Dim dict As Object
    Dim i As Long
    Dim ultimoValor As Long
    Dim ultimaFila As Long
    Dim key As Variant
    Dim valores As Variant

    ' Asignar la hoja
    'numCont = Trim(ComboBox1.Text)
    
    ' Verificar que el número de Solp tenga exactamente 10 dígitos
    If numCont = "" Then
        MsgBox "Ingrese un número de Contrato.", vbExclamation
        Exit Sub
    End If
    
    Set ht3 = ThisWorkbook.Sheets(numCont)
    
    Set ht1 = ThisWorkbook.Sheets("Imputaciones")
    
    ' Crear el diccionario para almacenar las concatenaciones únicas
    Set dict = CreateObject("Scripting.Dictionary")
    ' Limpiar el diccionario antes de usarlo nuevamente
    dict.RemoveAll

    ' Obtener la última fila con datos en la columna C
    ultimoValor = ht3.Cells(ht3.Rows.Count, "C").End(xlUp).Row
    
    ' Configurar el ProgressBar
    With Panel.ProgressBar1
        .Visible = True
        .Min = 1
        .Max = ultimoValor - 1
        .value = 1
    End With

    With ThisWorkbook.application
        .ScreenUpdating = False
        .Calculation = xlCalculationAutomatic
        .EnableEvents = False
    End With

    ' Recorrer las celdas de las columnas C y D
    For i = 2 To ultimoValor
        Set celdaC = ht3.Cells(i, "C")
        Set celdaD = ht3.Cells(i, "D")
        Set celdaE = ht3.Cells(i, "E")
        
        ' Concatenar los valores de las celdas C y D
        concatenacion = celdaD.value & " " & celdaC.value
        
        ' Verificar que la concatenación no esté vacía
        If concatenacion <> " " Then
            ' Si el valor concatenado no está en el diccionario, agregarlo junto con el valor de celdaE
            If Not dict.exists(concatenacion) Then
                ' Almacenar los valores concatenados y celdaE en una matriz
                valores = Array(concatenacion, celdaE.value)
                dict.Add concatenacion, valores
            End If
        End If
        
        ' Actualizar el ProgressBar
        Panel.ProgressBar1.value = i - 1
        DoEvents
    Next i

    ' Escribir los valores del diccionario en la hoja de "Certificados"
    i = 5
    For Each key In dict.Keys
        ht1.Cells(i, 1).value = dict(key)(0) ' Escribir el valor concatenado en la columna A
        ht1.Cells(i, 2).value = dict(key)(1) ' Escribir el valor de celdaE en la columna B
        i = i + 1
    Next key
    
    Call CalcularSumas
    
    ultimaFila = ht1.Cells(ht1.Rows.Count, 1).End(xlUp).Row
     ' Ordenar la tabla de acuerdo a la columna "TOTAL" (suponiendo que está en la columna C)
    With ht1.Sort
        .SortFields.Clear
        .SortFields.Add key:=ht1.Range("C4:C" & ultimaFila), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        .SetRange ht1.Range("A4:D" & ultimaFila) ' Ajusta el rango según tu tabla
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Set rng = ht1.Range("A4:D" & ultimaFila)
    
    ' Aplica bordes alrededor de todo el rango
    With rng.Borders
        .LineStyle = xlContinuous       ' Estilo de línea continua
        .Color = RGB(0, 0, 0)           ' Color negro (puedes cambiarlo)
        .Weight = xlThin                ' Grosor de la línea (puedes usar xlMedium, xlThick, etc.)
    End With
    

    With ThisWorkbook.application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
      
    ' Ocultar el ProgressBar una vez terminado el llenado
    Panel.ProgressBar1.Visible = False
End Sub

Sub CargarContratos()
Dim dict As Object
Dim ws As Worksheet
Dim nombreHoja As String
Dim key As Variant
Dim certificadosIndex As Long
Dim encontrado As Boolean
Dim i As Long

' Inicializar el diccionario
Set dict = CreateObject("Scripting.Dictionary")

' Inicializar variables
encontrado = False
certificadosIndex = 0

' Buscar la hoja "Certificados" y obtener su índice
For i = 1 To ThisWorkbook.Worksheets.Count
    If ThisWorkbook.Worksheets(i).Name = "Imputaciones" Then
        certificadosIndex = i
        encontrado = True
        Exit For
    End If
Next i

' Si la hoja "Certificados" no se encuentra, salir del procedimiento
If Not encontrado Then
    MsgBox "La hoja 'Certificados' no se encontró en el libro.", vbExclamation
    Exit Sub
End If

' Recorrer las hojas que están después de "Certificados" y agregarlas al diccionario
For i = certificadosIndex + 1 To ThisWorkbook.Worksheets.Count
    nombreHoja = ThisWorkbook.Worksheets(i).Name
    ' Solo agrega si el nombre de la hoja no está en el diccionario
    If Not dict.exists(nombreHoja) Then
        dict.Add nombreHoja, Nothing
    End If
Next i

' Limpiar el ComboBox antes de llenarlo
Panel.ComboBox1.Clear

' Llenar el ComboBox con los nombres de las hojas del diccionario
For Each key In dict.Keys
    Panel.ComboBox1.AddItem key
Next key
End Sub

Sub CalcularSumas()
    Dim ht1 As Worksheet
    Dim ht3 As Worksheet
    Dim numCont As String
    Dim lastRow As Long
    Dim ultimaFila As Long
    Dim i As Long
    Dim sumaTotal As Double
    Dim izqParte As String
    Dim derParte As Variant
    Dim saldo As Double
    Dim data As Variant
    Dim resultados() As Double
    Dim diccionario As Object
    Dim clave As String
    Const valorFijo As Double = 1000000000
    
    ' Asignar las hojas
    Set ht1 = ThisWorkbook.Sheets("Imputaciones")
    
    ' Obtener el valor seleccionado del ComboBox
    If Panel.ComboBox1.ListIndex <> -1 Then
        numCont = Trim(Panel.ComboBox1.value)
    Else
        MsgBox "Por favor, selecciona un Contrato.", vbExclamation
        Exit Sub
    End If
    
    ' Comprobar que la hoja existe antes de continuar
    On Error Resume Next
    Set ht3 = ThisWorkbook.Sheets(numCont)
    On Error GoTo 0
    
    If ht3 Is Nothing Then
        MsgBox "La hoja '" & numCont & "' no existe.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener la última fila con datos
    lastRow = ht3.Cells(ht3.Rows.Count, "F").End(xlUp).Row
    ultimaFila = ht1.Cells(ht1.Rows.Count, "B").End(xlUp).Row
    
    ' Leer los datos de ht3 en un diccionario
    Set diccionario = CreateObject("Scripting.Dictionary")
    data = ht3.Range("C2:F" & lastRow).value
    
    For i = LBound(data, 1) To UBound(data, 1)
        clave = data(i, 2) & "|" & data(i, 1) ' Combinar izqParte y derParte
        If Not diccionario.exists(clave) Then
            diccionario.Add clave, data(i, 4)
        Else
            diccionario(clave) = diccionario(clave) + data(i, 4)
        End If
    Next i

On Error Resume Next
    ' Configurar el ProgressBar
    With Panel.ProgressBar1
        .Visible = True
        .Min = 1
        .Max = ultimaFila - 4
        .value = 1
    End With
On Error GoTo 0

    ' Preparar resultados
    ReDim resultados(1 To ultimaFila - 4, 1 To 2)
    
    ' Calcular sumas usando el diccionario
    For i = 5 To ultimaFila
        sumaTotal = 0
        izqParte = Left(ht1.Cells(i, "A").value, InStr(1, ht1.Cells(i, "A").value, " ") - 1)
        derParte = ObtenerParteDerecha(ht1.Cells(i, "A").value)
        
        If IsNumeric(derParte) Then
            clave = izqParte & "|" & derParte
            If diccionario.exists(clave) Then
                sumaTotal = diccionario(clave)
            End If
        End If
        
        resultados(i - 4, 1) = sumaTotal
        saldo = valorFijo - sumaTotal
        resultados(i - 4, 2) = saldo
        
        ' Actualizar el ProgressBar
        If i Mod 100 = 0 Then
            On Error Resume Next
            Panel.ProgressBar1.value = i
            On Error GoTo 0
            DoEvents
        End If
    Next i
    
    ' Volcar resultados a la hoja
    ht1.Range("C5:D" & ultimaFila).value = resultados
    
    ' Ocultar el ProgressBar
    Panel.ProgressBar1.Visible = False
End Sub

Function ObtenerParteDerecha(Valor As String) As Variant
    Dim derParte As String
    derParte = Trim(Right(Valor, Len(Valor) - InStrRev(Valor, " ")))

    ' Verificar si es numérico y devolver el valor como número
    If IsNumeric(derParte) Then
        ObtenerParteDerecha = CLng(derParte)
    Else
        ObtenerParteDerecha = 0
    End If
End Function

Sub LlenarListView()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim itm As listItem
    Dim listView As listView
    
    ' Referencia a la hoja que contiene los datos
    Set ws = ThisWorkbook.Sheets("Imputaciones")
    
    ' Referencia al ListView en el formulario
    Set listView = Panel.ListView1
    
    ' Configuración de las columnas
    With listView
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False
        .ColumnHeaders.Clear
        .ListItems.Clear

        ' Configurar las columnas, incluyendo la nueva columna "Saldo"
        .ColumnHeaders.Add , , "Ser-Pos", 50
        .ColumnHeaders.Add , , "Descripción", 150
        .ColumnHeaders.Add , , "Total", 70
        .ColumnHeaders.Add , , "Saldo", 70
    End With
    
    ' Determinar la última fila con datos en la hoja
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Llenar el ListView con los datos
    For i = 5 To lastRow ' Asumiendo que las primeras 4 filas son encabezados
        Set itm = listView.ListItems.Add(, , ws.Cells(i, 1).value) ' Columna "Servicio y Posición"
        itm.SubItems(1) = ws.Cells(i, 2).value ' Columna "Descripción"
        itm.SubItems(2) = Format(ws.Cells(i, 3).value, "$#,##0.00") ' Columna "Total"
        itm.SubItems(3) = Format(ws.Cells(i, 4).value, "$#,##0.00") ' Columna "Saldo"
    Next i
    
listView.Refresh                  ' Refresca el control
listView.SetFocus          ' Regresa el foco al ListView

End Sub

Private Sub CommandButton4_Click()


Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Desea guardar los cambios antes de salir?", vbYesNoCancel + vbQuestion, "Confirmación de salida")

Select Case respuesta
    Case vbYes
        ' Guardar el libro y mostrar mensaje al finalizar
        ThisWorkbook.Save
        Unload Panel  ' Cierra el formulario
        
    Case vbNo
        ' Cierra el formulario sin guardar
        Unload Panel

    Case vbCancel
        ' No hace nada; el formulario permanece abierto
        Exit Sub
End Select


End Sub

Private Sub CommandButton5_Click()

    If Panel.ListView1.SelectedItem Is Nothing Then
        MsgBox "No hay elementos cargados en el Panel."
    Else: Call EnviarCorreo
    End If
    

End Sub

Sub EnviarCorreo()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim htmlBody As String
    Dim listItem As MSComctlLib.listItem
    Dim listSubItem As MSComctlLib.listSubItem
    Dim columnHeader As MSComctlLib.columnHeader
    Dim contrato As String
    Dim enviarTodasFilas As Boolean
    
    If Panel.ListView1.SelectedItem Is Nothing Then
        Exit Sub
    End If

    ' Inicializar el objeto de correo
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    contrato = Panel.ComboBox1.value
    enviarTodasFilas = Panel.CheckBox1.value ' True si el CheckBox1 está seleccionado, False si no lo está

    ' Comenzar el cuerpo del correo HTML
    htmlBody = "<html><head>"
    htmlBody = htmlBody & "<style>"
    htmlBody = htmlBody & "table { border-collapse: collapse; width: 100%; }"
    htmlBody = htmlBody & "th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; }"
    htmlBody = htmlBody & "th { background-color: #4CAF50; color: white; }"
    htmlBody = htmlBody & "tr:nth-child(even) { background-color: #f2f2f2; }"
    htmlBody = htmlBody & "tr:hover { background-color: #ddd; }"
    htmlBody = htmlBody & "h2 { color: #4CAF50; }"
    htmlBody = htmlBody & "</style></head><body>"
    htmlBody = htmlBody & "<h1>Contrato: " & contrato & "</h1>"
    htmlBody = htmlBody & "<h2>Certificación total por Servicios y Posición</h2>"
    htmlBody = htmlBody & "<p>Se adjunta listado con el total de certificaciones imputadas a los servicios en la posicion del contrato. El Saldo se calcula como la diferencia entre el Total imputado al servico en la posición, menos los $1000 millones que permite SAP. </p>"
    htmlBody = htmlBody & "<table>"

    ' Agregar encabezados de la tabla usando For Each
    htmlBody = htmlBody & "<tr>"
    For Each columnHeader In Panel.ListView1.ColumnHeaders
        htmlBody = htmlBody & "<th>" & columnHeader.Text & "</th>"
    Next columnHeader
    htmlBody = htmlBody & "</tr>"

    ' Verificar si se deben enviar todas las filas o solo las seleccionadas
    If enviarTodasFilas Then
        ' Iterar sobre todas las filas del ListView
        For Each listItem In Panel.ListView1.ListItems
            htmlBody = htmlBody & "<tr>"
            htmlBody = htmlBody & "<td>" & listItem.Text & "</td>" ' Columna principal
            For Each listSubItem In listItem.ListSubItems
                htmlBody = htmlBody & "<td>" & listSubItem.Text & "</td>"
            Next listSubItem
            htmlBody = htmlBody & "</tr>"
        Next listItem
    Else
        ' Iterar solo sobre las filas seleccionadas del ListView
        For Each listItem In Panel.ListView1.ListItems
            If listItem.Selected Then ' Verifica si la fila está seleccionada
                htmlBody = htmlBody & "<tr>"
                htmlBody = htmlBody & "<td>" & listItem.Text & "</td>" ' Columna principal
                For Each listSubItem In listItem.ListSubItems
                    htmlBody = htmlBody & "<td>" & listSubItem.Text & "</td>"
                Next listSubItem
                htmlBody = htmlBody & "</tr>"
            End If
        Next listItem
    End If

    ' Cerrar la tabla y el cuerpo del correo
    htmlBody = htmlBody & "</table>"
    htmlBody = htmlBody & "</body></html>"

    ' Configurar y enviar el correo
    With OutMail
        .To = ""
        .Subject = "Certificación total por Servicio - Contrato: " & contrato & ""
        .htmlBody = htmlBody
        .Display
    End With

    ' Limpiar objetos
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Private Sub UserForm_Initialize()
Dim ht1 As Worksheet
Set ht1 = ThisWorkbook.Sheets("Imputaciones")

Me.Width = 405
Me.Height = 300

Call CargarContratos
Call LlenarListView

Me.ListView1.ListItems.Clear
'Me.ListView1.ListItems = Nothing
Me.ProgressBar1.Visible = False
Me.ListView1.MultiSelect = True
Me.ComboBox1.value = ht1.Cells(2, 2).value
End Sub
