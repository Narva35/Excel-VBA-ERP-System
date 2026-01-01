VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFábricas 
   Caption         =   "Editar Fábricas"
   ClientHeight    =   4144
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9228.001
   OleObjectBlob   =   "frmFábricas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFábricas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdicionarFábrica_Click()
Unload Me
frmAdicionarFábrica.Show
End Sub

Private Sub cmdAlterarDados_Click()
Unload Me
frmDadosFábricas.Show
End Sub

Private Sub cmdRemoverFábrica_Click()
Unload Me
frmRemoverFábrica.Show
End Sub

Private Sub cmdVoltar_Click()
Unload Me
frmEditar.Show
End Sub

