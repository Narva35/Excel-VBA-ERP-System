VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdicionarCliente 
   Caption         =   "Adicionar Cliente"
   ClientHeight    =   10968
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   6420
   OleObjectBlob   =   "frmAdicionarCliente.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdicionarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAdicionar_Click()
    
    Dim ws As Worksheet
    Dim novalinha As Long
    Dim DataValida As Boolean
    
    Set ws = ThisWorkbook.Sheets("Clientes")
    
    If txtNome.Value = "" Or txtCEO.Value = "" Or txtIDFabrica.Value = "" Or txtIDCliente.Value = "" Or txtNIF.Value = "" Or txtLocalizacao.Value = "" Or txtTelefone.Value = "" Or txtData1Encomenda.Value = "" Or txtEmail.Value = "" Or txtFeedback.Value = "" Then
        MsgBox ("Deve preencher todos os campos obrigatórios.")
    Else
        DataValida = VerificarFormatoData(txtData1Encomenda.Value)
        
        If Not DataValida Then
            MsgBox "Por favor, insira a data no formato dd/mm/aaaa.", vbExclamation
            txtData1Encomenda.SetFocus
            Exit Sub
        End If
        
        novalinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        
        ws.Cells(novalinha, 2).Value = txtNome.Value
        ws.Cells(novalinha, 3).Value = txtCEO.Value
        ws.Cells(novalinha, 4).Value = txtIDFabrica.Value
        ws.Cells(novalinha, 5).Value = txtIDCliente.Value
        ws.Cells(novalinha, 6).Value = txtNIF.Value
        ws.Cells(novalinha, 7).Value = txtLocalizacao.Value
        ws.Cells(novalinha, 8).Value = txtTelefone.Value
        ws.Cells(novalinha, 9).Value = txtData1Encomenda.Value
        ws.Cells(novalinha, 10).Value = txtEmail.Value
        ws.Cells(novalinha, 11).Value = txtFeedback.Value
        ws.Cells(novalinha, 11).Value = CDbl(ws.Cells(novalinha, 11).Value)
        ws.Cells(novalinha, 12).Value = txtComentarios.Value

        MsgBox ("Cliente adicionado com sucesso!")
        MsgBox ("Por favor, altere os dados da fábrica referente ao cliente adicionado. Adicione +1 no número de clientes da fábrica")
    End If
    
    txtNome.Value = ""
    txtCEO.Value = ""
    txtIDFabrica.Value = ""
    txtIDCliente.Value = ""
    txtNIF.Value = ""
    txtLocalizacao.Value = ""
    txtTelefone.Value = ""
    txtData1Encomenda.Value = ""
    txtEmail.Value = ""
    txtFeedback.Value = ""
    txtComentarios.Value = ""

End Sub


Private Sub cmdVoltar_Click()
    
    Unload Me
    frmClientes.Show

End Sub

Private Sub UserForm_Initialize()

txtCEO.ControlTipText = ("Deve inserir o primeiro e último nomes do CEO.")
txtTelefone.ControlTipText = ("Digite no formato (indicativo) número. Exemplo: (351) 919195674")
txtData1Encomenda.ControlTipText = ("Digite no formato dd/mm/aaaa. Exemplo: 05/11/2016")
txtFeedback.ControlTipText = ("Entre 1,0 (mínimo) e 5,0 (máximo). Exemplo: 4,8")

End Sub
