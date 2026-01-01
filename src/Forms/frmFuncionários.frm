VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFuncionários 
   Caption         =   "Editar Funcionários"
   ClientHeight    =   4431
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10080
   OleObjectBlob   =   "frmFuncionários.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFuncionários"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdicionarFuncionário_Click()
Unload Me
frmAdicionarFuncionário.Show
End Sub

Private Sub cmdAlterarDados_Click()
Unload Me
frmDadosFuncionários.Show
End Sub

Private Sub cmdRemoverFuncionário_Click()
Unload Me
frmRemoverFuncionário.Show
End Sub

Private Sub cmdVoltar_Click()
Unload Me
frmEditar.Show
End Sub

