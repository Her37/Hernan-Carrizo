Attribute VB_Name = "Módulo1"
Function ExtraerNombre(celda As String) As String
    Dim partes() As String
    Dim nombreCompleto As String
    
    ' Dividir el texto en partes usando la coma
    partes = Split(celda, ".")
    
    ' Si hay al menos dos partes, obtener el nombre completo
    If UBound(partes) >= 1 Then
        nombreCompleto = partes(0) ' Tomamos la segunda parte (después de la coma)
    Else
        nombreCompleto = "Nombre no encontrado"
    End If
    
    ' Devolver el nombre completo sin el correo
    ExtraerNombre = Trim(nombreCompleto)
End Function

Sub EnviarCorreo()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim destinatario As String
    Dim asunto As String
    Dim cuerpo As String
    Dim contrato As String
    Dim ws As Worksheet
    Dim solpNumber As String
    Dim usuario As String
    Dim textBrv As String
    Dim texto As String
    Dim proveedor As String
    Dim nombre As String
    Dim ListBox As Object
    Dim actCel As Range
    Dim mail As String
    Dim firstRow As Long
    Dim lastRow As Long
    Dim i As Long
    Dim tablaHTML As String
    
    ' Asignar la hoja de cálculo
    Set ws = ThisWorkbook.Sheets("Solpes")
    ' Obtener la fila activa
    Set actCel = Selection
    
    firstRow = Selection.Rows(1).Row ' Primera fila de la selección
    lastRow = Selection.Rows(Selection.Rows.Count).Row ' Última fila de la selección
        
    If ws.Range("L" & firstRow).Value = "PR" Then
        mail = "ControlPresupuestarioDX@enel.com"
    ElseIf ws.Range("L" & firstRow).Value = "JD" Then
        mail = ws.Range("I" & firstRow).Value
    ElseIf ws.Range("L" & firstRow).Value = "DI" Then
        mail = "leonardo.bednarik@enel.com"
    ElseIf ws.Range("L" & firstRow).Value = "OC" Then
        mail = "jorge.fado@enel.com"
    Else
        mail = "INGRESE DESTINATARIO"
    End If
        
    If actCel.Rows.Count > 1 Then

        ' Configurar destinatario y asunto
        destinatario = mail
        asunto = "Solicitud de Monto por MMC - Resumen"

        ' Construir la tabla HTML
        tablaHTML = "<table border='1' style='border-collapse: collapse; width: 100%;'>" & _
                    "<tr><th>Contrato</th><th>Empresa</th><th>Monto</th><th>Solp</th></tr>"
        
        ' Recorrer las filas para llenar la tabla
        For i = firstRow To lastRow
            tablaHTML = tablaHTML & "<tr>" & _
                        "<td>" & ws.Cells(i, 2).Value & "</td>" & _
                        "<td>" & ws.Cells(i, 3).Value & "</td>" & _
                        "<td>" & ws.Cells(i, 7).Value & "</td>" & _
                        "<td>" & ws.Cells(i, 10).Value & "</td>" & _
                        "</tr>"
        Next i

        tablaHTML = tablaHTML & "</table>"

        ' Construir el cuerpo del correo
        cuerpo = "<p><strong>Estimado/a,</strong></p>" & _
                 "<p>Se adjunta la información consolidada de las solicitudes:</p>" & _
                 tablaHTML & _
                 "<p>Desde ya muchas gracias.</p>" & _
                 "<p>Saludos cordiales,</p>" & _
                 "<p><b>Hernán Carrizo</b><br>" & _
                 "Supply Chain - Work & Services</p>" & _
                 "</body></html>"
        
        ' Crear el objeto Outlook
        Set OutlookApp = CreateObject("Outlook.Application")
        Set OutlookMail = OutlookApp.CreateItem(0)
        
        With OutlookMail
            .To = destinatario
            .Subject = asunto
            .htmlBody = cuerpo
            .Display ' Cambia a .Send para enviar directamente
        End With
        
    Else
    
        ' Obtener valores de la hoja de Excel
        solpNumber = Trim(actCel.Rows.Value)
        
        contrato = ws.Range("B" & firstRow).Value
        proveedor = ws.Range("C" & firstRow).Value
        textBrv = ws.Range("D" & firstRow).Value
        texto = ws.Range("F" & firstRow).Value
        usuario = ws.Range("L" & firstRow).Value
        nombre = ExtraerNombre(mail)
        
        destinatario = mail
        asunto = textBrv & " - Contrato " & contrato & " - " & proveedor
        
                mail = "leonardo.bednarik@enel.com"
    If ws.Range("L" & firstRow).Value = "OC" Then
         cuerpo = "<html><body style='font-family: Aptos; color:#000000; font-size:12pt;'>" & _
                 "<p>Hola " & nombre & " Buendía,</p>" & _
                 "<p>Ya se encuentra en Me-Sing la <b>" & "CO" & "</b> del asunto a disposición y la SOLP N° <b>" & solpNumber & "</b> liberada con los archivos adjuntos y firmados. Correspondiente a:</p>" & _
                 "<blockquote style='margin-left:30px; font-style:italic; color:#666666;'>" & texto & "</blockquote>" & _
                 "<p>Desde ya muchas gracias.</p>" & _
                 "<p>Saludos cordiales,</p>" & _
                 "<p><b>Hernán Carrizo</b><br>" & _
                 "Supply Chain - Work & Services</p>" & _
                 "</body></html>"
    Else
        cuerpo = "<html><body style='font-family: Aptos; color:#000000; font-size:12pt;'>" & _
                "<p>Hola " & nombre & " Buendía,</p>" & _
                "<p>Se encuentran disponibles para liberar en <b>" & usuario & "</b> la SOLP N° <b>" & solpNumber & "</b> correspondiente a:</p>" & _
                "<blockquote style='margin-left:30px; font-style:italic; color:#666666;'>" & texto & "</blockquote>" & _
                "<p>Desde ya muchas gracias.</p>" & _
                "<p>Saludos,</p>" & _
                "<p><b>Ing. Hernán Carrizo</b><br>" & _
                "Supply Chain - Work & Services</p>" & _
                "</body></html>"
    End If
    
        ' Crear el objeto Outlook
        Set OutlookApp = CreateObject("Outlook.Application")
        Set OutlookMail = OutlookApp.CreateItem(0)
        
        ' Configurar los parámetros del correo
        With OutlookMail
            .To = destinatario
            .CC = "joaquin.o.sanchez@enel.com"
            .Subject = asunto
            .htmlBody = cuerpo
            .Display
        End With
        
    End If

    ' Liberar objetos
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub

