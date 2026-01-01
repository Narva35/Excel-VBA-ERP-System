VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEncomendas 
   Caption         =   "Editar Encomendas"
   ClientHeight    =   4403
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9264.001
   OleObjectBlob   =   "frmEncomendas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEncomendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdicionarEncomenda_Click()
Unload Me
frmAdicionarEncomenda.Show
End Sub

Private Sub cmdAlterarDados_Click()
Unload Me
frmDadosEncomendas.Show
End Sub

Private Sub cmdRemoverEncomenda_Click()
Unload Me
frmRemoverEncomenda.Show
End Sub

Private Sub cmdVoltar_Click()
Unload Me
frmEditar.Show
End Sub

