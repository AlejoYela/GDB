VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Inicio 
   Caption         =   "Inicio"
   ClientHeight    =   2124
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8412.001
   OleObjectBlob   =   "Inicio.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CrearHv_Click()
    Unload Me
    Hv.Show
End Sub

Private Sub CrearMtto_Click()
    Unload Me
    Mttos.Show
End Sub

Private Sub Lista_Click()
    Dim codigo As Integer
    codigo = Me.Lista.List(Lista.ListIndex, 0)
    MsgBox (codigo)
End Sub

Private Sub Salir_Click()
    Unload Me
End Sub


