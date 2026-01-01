VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditar 
   Caption         =   "Editar"
   ClientHeight    =   5173
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11064
   OleObjectBlob   =   "frmEditar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClientes_Click()
    Unload Me
    frmClientes.Show
End Sub

Private Sub cmdEncomendas_Click()
    Unload Me
    frmEncomendas.Show
End Sub

Private Sub cmdFábricas_Click()
   Unload Me
    frmFábricas.Show
End Sub

Private Sub cmdFuncionários_Click()
    Unload Me
    frmFuncionários.Show
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
    frmPrincipal.Show
End Sub

