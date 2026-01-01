VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdicionarEncomenda 
   Caption         =   "Adicionar Encomenda"
   ClientHeight    =   8435.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6024
   OleObjectBlob   =   "frmAdicionarEncomenda.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdicionarEncomenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdAdicionar_Click()
    
    Dim ws As Worksheet
    Dim novalinha As Long
    
    Set ws = ThisWorkbook.Sheets("Encomendas")
    
If txtUnidades = "" Or txtCustoProduto = "" Or txtIVA = "" Or txtProduto.Value = "" Or txtDataCompra.Value = "" Or txtDataEnvio.Value = "" Or txtDataChegada.Value = "" Or txtIDCliente.Value = "" Then
MsgBox ("Introduza os dados necessários para adicionar uma encomenda.")
Else
    novalinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(novalinha, 3).NumberFormat = "dd/mm/yyyy"
    ws.Cells(novalinha, 4).NumberFormat = "dd/mm/yyyy"
    ws.Cells(novalinha, 5).NumberFormat = "dd/mm/yyyy"
    ws.Cells(novalinha, 2).Value = txtProduto.Value
    ws.Cells(novalinha, 3).Value = txtDataCompra.Value
    ws.Cells(novalinha, 4).Value = txtDataEnvio.Value
    ws.Cells(novalinha, 5).Value = txtDataChegada.Value
    ws.Cells(novalinha, 8).Value = txtIDCliente.Value
    ws.Cells(novalinha, 9).Value = ComboBox1.Value
    ws.Cells(novalinha, 10).Value = txtUnidades.Value
    ws.Cells(novalinha, 11).Value = txtCustoProduto.Value
    ws.Cells(novalinha, 12).Value = txtIVA.Value / 100
   
    MsgBox ("Encomenda adicionada com sucesso!")
 End If
    txtProduto.Value = ""
    txtDataCompra.Value = ""
    txtDataEnvio.Value = ""
    txtDataChegada.Value = ""
    txtIDCliente.Value = ""
    ComboBox1.Value = ""
    txtUnidades.Value = ""
    txtCustoProduto.Value = ""
    txtIVA.Value = ""
    
Unload Me
frmEncomendas.Show

End Sub

Private Sub cmdVoltar_Click()
Unload Me
frmEncomendas.Show
End Sub


Private Sub UserForm_Initialize()
Dim ws As Worksheet
Dim colFabs As New Collection
Dim i As Long
Dim fab As String

txtDataEnvio.ControlTipText = "Deve introduzir a data no formato dd/mm/aaaa. Exemplo: 06/11/2017"
txtDataCompra.ControlTipText = "Deve introduzir a data no formato dd/mm/aaaa. Exemplo: 11/05/2015"
txtDataChegada.ControlTipText = "Deve introduzir a data no formato dd/mm/aaaa. Exemplo: 07/08/2018"
txtCustoProduto.ControlTipText = "O valor é referente a euros. Exemplo: se introduzir 17,1 refere-se a 17,10%"""
txtIVA.ControlTipText = "Deve introduzir o número decimal. Exemplo: 0,23 refere-se a 23%"

Set ws = ThisWorkbook.Sheets("Fábricas")

fab = ws.Cells(2, 3).Value

For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If Not IsInCollection(colFabs, fab) Then
        colFabs.Add fab, fab
    End If

    fab = ws.Cells(i + 1, 3).Value
Next i

Dim vec() As String
ReDim vec(1 To colFabs.Count)
For i = 1 To colFabs.Count
    vec(i) = colFabs(i)
Next i

ComboBox1.List = vec

End Sub

Function IsInCollection(col As Collection, key As String) As Boolean
    On Error Resume Next
    IsInCollection = Not col(key) Is Nothing
    On Error GoTo 0
End Function

Private Sub ComboBox1_Change()
txtIDCliente.Value = ComboBox1.Value & "-C"
End Sub
