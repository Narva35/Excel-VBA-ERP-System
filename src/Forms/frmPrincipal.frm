VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrincipal 
   Caption         =   "Trabalho Prático 2"
   ClientHeight    =   10766
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17160
   OleObjectBlob   =   "frmPrincipal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdFechar_Click()

Me.Hide

End Sub

Private Sub cmdVisualizar_Click()

    Me.Hide
    frmVisualizar.Show

End Sub

Private Sub cmdEditar_Click()

    Me.Hide
    frmAcesso.Show

End Sub

Private Sub cmdEstatisticas_Click()

    Me.Hide
    frmEstatisticas.Show

End Sub

