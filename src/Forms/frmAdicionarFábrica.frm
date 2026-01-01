VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdicionarFábrica 
   Caption         =   "Adicionar Fábrica"
   ClientHeight    =   10311
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8724.001
   OleObjectBlob   =   "frmAdicionarFábrica.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdicionarFábrica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAdicionar_Click()

Dim novalinha As Integer
Dim ws As Worksheet

Set ws = ThisWorkbook.Sheets("Fábricas")

If txtNome.Value = "" Or txtID.Value = "" Or txtTelefone.Value = "" Or txtClientes.Value = "" Or txtMorada.Value = "" Or txtPaís.Value = "" Or txtFundação.Value = "" Or txtIDDiretor.Value = "" Or txtÁrea.Value = "" Or txtDespesas.Value = "" Or txtFaturação.Value = "" Or txtResultadoLíquido.Value = "" Or txtFuncionários.Value = "" Or txtCapacidade = "" Then
    MsgBox ("Deve preencher todos os campos.")
Else
    novalinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(novalinha, 1) = novalinha - 1
    ws.Cells(novalinha, 2) = txtNome.Value
    ws.Cells(novalinha, 3) = txtID.Value
    ws.Cells(novalinha, 4) = txtTelefone.Value
    ws.Cells(novalinha, 5) = txtClientes.Value
    ws.Cells(novalinha, 6) = txtMorada.Value
    ws.Cells(novalinha, 7) = txtPaís.Value
    ws.Cells(novalinha, 8) = txtFundação.Value
    ws.Cells(novalinha, 9) = txtIDDiretor.Value
    ws.Cells(novalinha, 10) = txtÁrea.Value
    ws.Cells(novalinha, 11) = txtDespesas.Value
    ws.Cells(novalinha, 12) = txtFaturação.Value
    ws.Cells(novalinha, 13) = txtResultadoLíquido.Value
    ws.Cells(novalinha, 14) = txtFuncionários.Value
    ws.Cells(novalinha, 15) = txtCapacidade.Value
    
    MsgBox ("Fábrica adicionada com sucesso!")
    MsgBox ("Ao adicionar a fábrica certifique-se que adiciona também os funcionários referidos e também os clientes.")
End If
    
    txtNome.Value = ""
    txtID.Value = ""
    txtTelefone.Value = ""
    txtClientes.Value = ""
    txtMorada.Value = ""
    txtPaís.Value = ""
    txtFundação.Value = ""
    txtIDDiretor.Value = ""
    txtÁrea.Value = ""
    txtDespesas.Value = ""
    txtFaturação.Value = ""
    txtResultadoLíquido.Value = ""
    txtFuncionários.Value = ""
    txtCapacidade.Value = ""
    
End Sub

Private Sub cmdVoltar_Click()
Unload Me
frmFábricas.Show
End Sub


Private Sub UserForm_Initialize()

txtFundação.ControlTipText = ("Introduza a data no formato dd/mm/aaaa. Exemplo: 07/12/2017")
txtDespesas.ControlTipText = ("Introduza o valor em milhões. Exemplo: 3,1 refere-se a 3,1 milhões de euros.")
txtFaturação.ControlTipText = ("Introduza o valor em milhões. Exemplo: 3,1 refere-se a 3,1 milhões de euros.")
txtResultadoLíquido.ControlTipText = ("Introduza o valor em milhões. Exemplo: 3,1 refere-se a 3,1 milhões de euros.")
txtÁrea.ControlTipText = ("O valor refere-se a metros quadrados.")
txtCapacidade.ControlTipText = ("O valor refere-se a toneladas.")

End Sub
