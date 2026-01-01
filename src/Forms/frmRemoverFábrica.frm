VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRemoverFábrica 
   Caption         =   "Remover Fábrica"
   ClientHeight    =   9058.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7752
   OleObjectBlob   =   "frmRemoverFábrica.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRemoverFábrica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRemover_Click()
Dim ws As Worksheet
    Dim FábricaSelecionado As String
    Dim FábricaID As String
    Dim linha As Long
    Dim lastRow As Long
    Dim pos As Integer

    Set ws = ThisWorkbook.Sheets("Fábricas")
    
    If lstFábrica.ListIndex = -1 Then
        MsgBox ("Por favor, selecione uma fábrica para remover.")
        Exit Sub
    End If
    
    FábricaSelecionado = lstFábrica.Value
    
    pos = InStr(FábricaSelecionado, "ID: ") + 4
    FábricaID = Mid(FábricaSelecionado, pos)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For linha = 2 To lastRow
        If ws.Cells(linha, 3).Value = FábricaID Then
            ws.Rows(linha).Delete
            Exit For
        End If
    Next linha
    
    lstFábrica.RemoveItem lstFábrica.ListIndex
    
    MsgBox ("Fábrica removida com sucesso.")
End Sub

Private Sub cmdVoltar_Click()
Unload Me
frmFábricas.Show
End Sub


Private Sub UserForm_Initialize()
Dim ws As Worksheet
    Dim linha As Long
    Dim FábricaInfo As String
    
    Set ws = ThisWorkbook.Sheets("Fábricas")
    
    lstFábrica.clear
    
    For linha = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        FábricaInfo = "Nome: " & ws.Cells(linha, 2).Value & " - ID: " & ws.Cells(linha, 3).Value
        lstFábrica.AddItem FábricaInfo
    Next linha
End Sub


