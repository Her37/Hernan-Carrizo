Attribute VB_Name = "Módulo5"
'*************************************
'|||REALIZADO POR HERNAN F. CARRIZO|||
'************AGO-SEP 2024*************

' Función para codificar en Base64
Function EncodeBase64(ByVal inputText As String) As String
    Dim bytes() As Byte
    Dim base64Encoded As String
    Dim xml As Object

    Set xml = CreateObject("MSXML2.DOMDocument")
    bytes = StrConv(inputText, vbFromUnicode)
    
    With xml.createElement("root")
        .DataType = "bin.base64"
        .nodeTypedValue = bytes
        base64Encoded = .Text
    End With
    
    Set xml = Nothing

    EncodeBase64 = base64Encoded
End Function

' Función para decodificar desde Base64
Function DecodeBase64(ByVal encodedText As String) As String
    Dim xml As Object
    Dim decodedBytes() As Byte
    Dim decodedString As String

    Set xml = CreateObject("MSXML2.DOMDocument")
    
    With xml.createElement("root")
        .DataType = "bin.base64"
        .Text = encodedText
        decodedBytes = .nodeTypedValue
    End With
    
    decodedString = StrConv(decodedBytes, vbUnicode)
    Set xml = Nothing

    DecodeBase64 = decodedString
End Function

' Función para encriptar texto
Function EncriptarTexto(ByVal texto As String, ByVal clave As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim resultado As String
    Dim claveLength As Integer
    
    resultado = ""
    claveLength = Len(clave)
    
    For i = 1 To Len(texto)
        j = (i - 1) Mod claveLength + 1
        resultado = resultado & Chr(Asc(Mid(texto, i, 1)) Xor Asc(Mid(clave, j, 1)))
    Next i
    
    ' Codificar el resultado en Base64 para manejar caracteres binarios
    EncriptarTexto = EncodeBase64(resultado)
End Function

' Función para desencriptar texto
Function DesencriptarTexto(ByVal texto As String, ByVal clave As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim resultado As String
    Dim claveLength As Integer
    Dim textoDecodificado As String
    
    ' Decodificar el texto desde Base64
    textoDecodificado = DecodeBase64(texto)
    
    resultado = ""
    claveLength = Len(clave)
    
    For i = 1 To Len(textoDecodificado)
        j = (i - 1) Mod claveLength + 1
        resultado = resultado & Chr(Asc(Mid(textoDecodificado, i, 1)) Xor Asc(Mid(clave, j, 1)))
    Next i
    
    DesencriptarTexto = resultado
End Function

' Subrutina para guardar las credenciales en un archivo
Sub GuardarCredenciales(usuario As String, contrasena As String, recordar As Boolean)
    Dim archivo As Integer
    Dim rutaArchivo As String
    Dim datosEncriptados As String
    Dim claveEncriptacion As String
    
    ' Ruta del archivo de credenciales
    rutaArchivo = ThisWorkbook.Path & "\credenciales.dat"
    claveEncriptacion = "MiClaveSecreta"
    
    If recordar Then
        ' Encriptar las credenciales
        datosEncriptados = EncriptarTexto(usuario & "|" & contrasena, claveEncriptacion)
        
        ' Guardar las credenciales encriptadas en un archivo
        archivo = FreeFile
        Open rutaArchivo For Output As archivo
        Print #archivo, datosEncriptados
        Close archivo
    Else
        ' Eliminar el archivo de credenciales si existe
        If Dir(rutaArchivo) <> "" Then
            Kill rutaArchivo
        End If
    End If
End Sub


' Función para cargar las credenciales desde un archivo
Function CargarCredenciales() As Variant
    Dim archivo As Integer
    Dim rutaArchivo As String
    Dim datosEncriptados As String
    Dim datosDesencriptados As String
    Dim partes() As String
    Dim claveEncriptacion As String
    
    rutaArchivo = ThisWorkbook.Path & "\credenciales.dat"
    claveEncriptacion = "MiClaveSecreta"
    
    ' Verificar si el archivo existe
    If Dir(rutaArchivo) <> "" Then
        archivo = FreeFile
        Open rutaArchivo For Input As archivo
        Line Input #archivo, datosEncriptados
        Close archivo
        
        ' Desencriptar los datos
        datosDesencriptados = DesencriptarTexto(datosEncriptados, claveEncriptacion)
        partes = Split(datosDesencriptados, "|")
        
        ' Devolver el usuario y la contraseña
        CargarCredenciales = Array(partes(0), partes(1))
    Else
        ' Si no existe el archivo, devolver vacío
        CargarCredenciales = Array("", "")
    End If
End Function


