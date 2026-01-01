VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRemoverEncomenda 
   Caption         =   "Remover Encomenda"
   ClientHeight    =   9807.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7020
   OleObjectBlob   =   "frmRemoverEncomenda.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRemoverEncomenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRemover_Click()
Dim ws As Worksheet
    Dim EncomendaSelecionado As String
    Dim linha As Long
    Dim lastRow As Long
    

    Set ws = ThisWorkbook.Sheets("Encomendas")
    
    If lstEncomenda.ListIndex = -1 Then
        MsgBox ("Por favor, selecione uma encomenda para remover.")
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For linha = 2 To lastRow
        If ws.Cells(linha, 2).Value = lstEncomenda.Value Then
            ws.Rows(linha).Delete
            Exit For
        End If
    Next linha
    
    EncomendaSelecionado = lstEncomenda.ListIndex
    lstEncomenda.RemoveItem EncomendaSelecionado
    
    MsgBox ("Encomenda removida com sucesso.")
End Sub

Private Sub cmdVoltar_Click()
Unload Me
frmEncomendas.Show
End Sub


Private Sub UserForm_Initialize()
Dim ws As Worksheet
    Dim linha As Long
    Dim EncomendaInfo As String
    
    Set ws = ThisWorkbook.Sheets("Encomendas")
    
    lstEncomenda.clear
    
    For linha = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        EncomendaInfo = "Nome: " & ws.Cells(linha, 2).Value
        lstEncomenda.AddItem EncomendaInfo
    Next linha
End Sub
