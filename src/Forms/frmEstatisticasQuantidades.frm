VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstatisticasQuantidades 
   Caption         =   "EstatisticasQuantidades"
   ClientHeight    =   4263
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8256.001
   OleObjectBlob   =   "frmEstatisticasQuantidades.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstatisticasQuantidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbQtdPais_Change()
    txbRespostaQuantidades = ""
    Me.cmdnumfabricas.Value = False
    Me.cmdnumfuncionarios.Value = False
    Me.cmdnumencomendas.Value = False
    Me.cmdnumclientes.Value = False
End Sub

Private Sub cmdnumclientes_Click()
    Dim ws As Worksheet, linha As Integer
    Set ws = ThisWorkbook.Sheets("Fábricas")
    linha = ws.Cells(Rows.Count, "G").End(xlUp).Row
    txbRespostaQuantidades = WorksheetFunction.SumIf(ws.Range("G2:G" & linha), cmbQtdPais.Value, ws.Range("E2:E" & linha))
End Sub

Private Sub cmdnumencomendas_Click()
    Dim tbl As ListObject, lookup As Range, id As String
    Dim ws As Worksheet, linha As Integer
    Set ws = ThisWorkbook.Sheets("Fábricas")
    Set tbl = ws.ListObjects(1)
    Set lookup = tbl.ListColumns(7).DataBodyRange
    linha = WorksheetFunction.Match(cmbQtdPais.Value, lookup, 0)
    id = WorksheetFunction.Index(tbl.ListColumns(3).DataBodyRange, linha)
    id = Left(id, Len(id) - 2)
    Dim ws2 As Worksheet, tabela As ListObject, coluna As Range, c As Range, tirar As String, soma As Integer
    Set ws2 = ThisWorkbook.Sheets("Encomendas")
    Set tabela = ws2.ListObjects(1)
    Set coluna = tabela.ListColumns(9).DataBodyRange
    soma = 0
    For Each c In coluna
        tirar = Left(c.Value, Len(c.Value) - 2)
        If tirar = id Then
            soma = soma + 1
        End If
    Next c
    txbRespostaQuantidades = soma
End Sub

Private Sub cmdnumfabricas_Click()
    Dim ws As Worksheet, linha As Integer
    Set ws = ThisWorkbook.Sheets("Fábricas")
    linha = ws.Cells(Rows.Count, "G").End(xlUp).Row
    txbRespostaQuantidades = WorksheetFunction.CountIf(ws.Range("G2:G" & linha), cmbQtdPais.Value)
End Sub

Private Sub cmdnumfuncionarios_Click()
    Dim ws As Worksheet, linha As Integer
    Set ws = ThisWorkbook.Sheets("Fábricas")
    linha = ws.Cells(Rows.Count, "G").End(xlUp).Row
    txbRespostaQuantidades = WorksheetFunction.SumIf(ws.Range("G2:G" & linha), cmbQtdPais.Value, ws.Range("N2:N" & linha))
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet, i As Integer, linha As Integer
    Set ws = ThisWorkbook.Sheets("Fábricas")
    linha = ws.Cells(Rows.Count, "G").End(xlUp).Row
    For i = 2 To linha
        Dim valor As String
        valor = ws.Cells(i, "G")
        If WorksheetFunction.CountIf(ws.Range("G2:G" & i), valor) = 1 Then
            Me.cmbQtdPais.AddItem valor
        End If
    Next

End Sub

Private Sub cmbVoltar_Click()
  Unload Me
    frmEstatisticas.Show
End Sub
