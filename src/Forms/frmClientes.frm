VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClientes 
   Caption         =   "Editar Clientes"
   ClientHeight    =   4116
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8880.001
   OleObjectBlob   =   "frmClientes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdicionarCliente_Click()
Unload Me
frmAdicionarCliente.Show

End Sub


Private Sub cmdAlterarDados_Click()
Unload Me
frmDadosClientes.Show
End Sub

Private Sub cmdRemoverCliente_Click()
Unload Me
frmRemoverClient.Show

End Sub

Private Sub cmdVoltar_Click()
Unload Me
frmEditar.Show
End Sub
