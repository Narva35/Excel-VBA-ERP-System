VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAcesso 
   Caption         =   "Login"
   ClientHeight    =   6396
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8568.001
   OleObjectBlob   =   "frmAcesso.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEntrar_Click()

    If (txtUtilizador.Text = "Paula" And txtPasse.Text = "paula123") Or (txtUtilizador.Text = "Maria" And txtPasse.Text = "maria123") Then
        MsgBox "Acesso concedido. Bem-vinda, " & txtUtilizador.Text & "!", vbInformation
        Unload Me
        frmEditar.Show
    ElseIf (txtUtilizador.Text = "Gonçalo" And txtPasse.Text = "goncalo123") Or _
           (txtUtilizador.Text = "Ekumby" And txtPasse.Text = "ekumby123") Then
        MsgBox "Acesso concedido. Bem-vindo, " & txtUtilizador.Text & "!", vbInformation
        Unload Me
        frmEditar.Show
    Else
        MsgBox "Dados de acesso errados.", vbExclamation
        txtUtilizador.SetFocus
    End If

    txtUtilizador.Text = ""
    txtPasse.Text = ""

End Sub

Private Sub cmdVoltar_Click()

    Unload Me
    frmPrincipal.Show

End Sub

Private Sub CommandButton1_Click()

frmCredenciais.Show

End Sub

Private Sub UserForm_Click()

End Sub
