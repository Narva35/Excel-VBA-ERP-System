VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstatisticas 
   Caption         =   "Estatísticas"
   ClientHeight    =   5509
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9468.001
   OleObjectBlob   =   "frmEstatisticas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstatisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdVoltar_Click()
  Unload Me
    frmPrincipal.Show
End Sub

Private Sub cmdClientes_Click()
  Unload Me
    frmEstatisticasClientes2.Show
End Sub

Private Sub cmdFabricas_Click()
  Unload Me
    frmEstatisticasFabricas.Show
End Sub

Private Sub cmdEncomendas_Click()
  Unload Me
    frmEstatisticasEncomendas.Show
End Sub

Private Sub cmdPercentagens_Click()
  Unload Me
    frmEstatisticasPercentagens.Show
End Sub

Private Sub cmdFuncionarios_Click()
  Unload Me
    frmEstatisticasFuncionarios.Show
End Sub

Private Sub cmdMedias_Click()
  Unload Me
    frmEstatisticasMedias.Show
End Sub

Private Sub cmdQuantidades_Click()
  Unload Me
    frmEstatisticasQuantidades.Show
End Sub


