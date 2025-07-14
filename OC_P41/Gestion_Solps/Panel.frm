VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Panel 
   Caption         =   "Panel"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
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

Private Sub CommandButton2_Click()
    cancelarProceso = True
End Sub

Private Sub UserForm_Click()

End Sub
