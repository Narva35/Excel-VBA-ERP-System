VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRemoverClient 
   Caption         =   "Remover Cliente"
   ClientHeight    =   10192
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   6900
   OleObjectBlob   =   "frmRemoverClient.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRemoverClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVoltar_Click()

    Unload Me
    frmClientes.Show
End Sub

Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Dim linha As Integer
    Dim clienteInfo As String
    
    Set ws = ThisWorkbook.Sheets("Clientes")
    
    lstClientes.clear
    
    For linha = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        clienteInfo = "Nome: " & ws.Cells(linha, 2).Value & " - ID: " & ws.Cells(linha, 5).Value
        lstClientes.AddItem clienteInfo
    Next linha
    
End Sub

Private Sub cmdRemover_Click()

    Dim ws As Worksheet
    Dim clienteSelecionado As String
    Dim clienteID As String
    Dim linha As Long
    Dim lastRow As Long
    Dim pos As Integer

    Set ws = ThisWorkbook.Sheets("Clientes")
    
    If lstClientes.ListIndex = -1 Then
        MsgBox ("Por favor, selecione um cliente para remover.")
        Exit Sub
    End If
    
    clienteSelecionado = lstClientes.Value
    
    pos = InStr(clienteSelecionado, "ID: ") + 4
    clienteID = Mid(clienteSelecionado, pos)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For linha = 2 To lastRow
        If ws.Cells(linha, 5).Value = clienteID Then
            ws.Rows(linha).Delete
            Exit For
        End If
    Next linha
    
    lstClientes.RemoveItem lstClientes.ListIndex
    
    MsgBox ("Cliente removido com sucesso.")
    
End Sub



