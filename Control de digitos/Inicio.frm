VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Inicio 
   Caption         =   "Inicio de Sesi�n SAP"
   ClientHeight    =   4032
   ClientLeft      =   -252
   ClientTop       =   -1644
   ClientWidth     =   5796
   OleObjectBlob   =   "Inicio.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Public usuario As String
    Public contrasena As String

Private Sub btnConfirmar_Click()
    
    usuario = Inicio.txtUsuario.Text ' Asumiendo que tienes un TextBox llamado txtUsuario
    contrasena = Inicio.txtContrase�a.Text ' Asumiendo que tienes un TextBox llamado txtContrasena
    
    ' Verificar que los campos no est�n vac�os
    If usuario = "" Or contrasena = "" Then
        MsgBox "Por favor, ingrese ambos, usuario y contrase�a."
        Exit Sub
    End If
    
    ' Guardar o eliminar las credenciales seg�n el estado del checkbox
    Call GuardarCredenciales(usuario, contrasena, chkRecordar.value)
    
    ' Cerrar el formulario
    Inicio.Hide
    
    ' Indicar que el formulario se cerr� correctamente
    Inicio.Tag = "OK"
End Sub


Private Sub CommandButton1_Click()
Unload Inicio
Inicio.Tag = "Exit"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Inicio.Tag = "Exit"
        
    ' Si el formulario se cierra de manera inesperada, cancelar el proceso
    If CloseMode = vbFormControlMenu Then
        'MsgBox "El formulario se ha cerrado sin guardar las credenciales. El proceso se detendr�."
        Cancel = True
        Inicio.Tag = "Cancel"
    End If
    
End Sub

Private Sub UserForm_Initialize()
    Dim credenciales As Variant
    
    ' Cargar las credenciales si est�n guardadas
    credenciales = CargarCredenciales()
    
    If credenciales(0) <> "" Then
        txtUsuario.Text = credenciales(0)
        txtContrase�a.Text = credenciales(1)
        chkRecordar.value = True
    End If
    
    Inicio.Width = 230
    Inicio.Height = 180
    
    
End Sub


    
