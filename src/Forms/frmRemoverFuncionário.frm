VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRemoverFuncionário 
   Caption         =   "Remover Funcionário"
   ClientHeight    =   10668
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7392
   OleObjectBlob   =   "frmRemoverFuncionário.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRemoverFuncionário"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRemover_Click()
Dim ws As Worksheet
    Dim FuncionárioSelecionado As String
    Dim FuncionárioID As String
    Dim linha As Long
    Dim lastRow As Long
    Dim pos As Integer

    Set ws = ThisWorkbook.Sheets("Funcionários")
    
    If lstFuncionários.ListIndex = -1 Then
        MsgBox ("Por favor, selecione um funcionário para remover.")
        Exit Sub
    End If
    
    FuncionárioSelecionado = lstFuncionários.Value
    
    pos = InStr(FuncionárioSelecionado, "ID: ") + 4
    FuncionárioID = Mid(FuncionárioSelecionado, pos)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For linha = 2 To lastRow
        If ws.Cells(linha, 4).Value = FuncionárioID Then
            ws.Rows(linha).Delete
            Exit For
        End If
    Next linha
    
    lstFuncionários.RemoveItem lstFuncionários.ListIndex
    
    MsgBox ("Funcionário removido com sucesso.")
    
End Sub

Private Sub cmdVoltar_Click()
Unload Me
frmFuncionários.Show
End Sub

Private Sub UserForm_Initialize()
Dim ws As Worksheet
    Dim linha As Long
    Dim FuncionárioInfo As String
    
    Set ws = ThisWorkbook.Sheets("Funcionários")
    
    lstFuncionários.clear
    
    For linha = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        FuncionárioInfo = ws.Cells(linha, 2).Value
        lstFuncionários.AddItem FuncionárioInfo
    Next linha
End Sub

Private Sub txtprocurar_Change()

 Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim filterText As String

    filterText = Me.txtProcurar.Text
    Me.lstFuncionários.clear

    Set ws = ThisWorkbook.Sheets("Funcionários")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For i = 1 To lastRow
        If InStr(1, ws.Cells(i + 1, 2).Value, filterText, vbTextCompare) > 0 Then
            Me.lstFuncionários.AddItem ws.Cells(i + 1, 2).Value
        End If
    Next i
    
End Sub

