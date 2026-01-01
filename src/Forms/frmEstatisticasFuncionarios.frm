VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstatisticasFuncionarios 
   Caption         =   "Estatisticas Funcionários"
   ClientHeight    =   7462
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12804
   OleObjectBlob   =   "frmEstatisticasFuncionarios.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstatisticasFuncionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbFabricasFuncionarios_Change()
    Dim i As Integer
    For i = 1 To 11
        Me.Controls("TextBox" & i).Text = ""
    Next i
    
    Me.opbFunAntg.Value = False
    Me.opbFuncRecente.Value = False
    Me.opbFuncVelho.Value = False
    Me.opbFuncNovo.Value = False
    Me.opbFuncmaiorVencimento.Value = False
    Me.opbFuncmenosVenc.Value = False
End Sub

Private Sub cmbVoltar_Click()
    Unload Me
    frmEstatisticas.Show
End Sub

Private Sub opbFunAntg_Click()
    Dim ws As Worksheet, tabela As ListObject, linha As Long, id As String
    Dim ws2 As Worksheet, tbl As ListObject, coluna As Range, c As Range
    Dim novo As Date, pos As Range

    Set ws = ThisWorkbook.Sheets("Fábricas")
    Set tabela = ws.ListObjects(1)
    linha = Application.WorksheetFunction.Match(cmbFabricasFuncionarios.Value, tabela.ListColumns(2).DataBodyRange, 0)
    id = Application.WorksheetFunction.Index(tabela.ListColumns(3).DataBodyRange, linha)
    Set ws2 = ThisWorkbook.Sheets("Funcionários")
    Set tbl = ws2.ListObjects(1)
    Set coluna = tbl.ListColumns(3).DataBodyRange

    novo = DateSerial(9999, 12, 31)

    For Each c In coluna
        If id = c.Value And IsDate(c.Offset(0, 5).Value) And c.Offset(0, 5).Value < novo Then
            novo = c.Offset(0, 5).Value
            Set pos = c
        End If
    Next c

    TextBox1.Text = pos.Offset(0, -1).Value
    TextBox2.Text = pos.Offset(0, 1).Value
    TextBox3.Text = pos.Offset(0, 9).Value
    TextBox4.Text = pos.Offset(0, 3).Value
    TextBox7.Text = Format(pos.Offset(0, 5).Value, "dd/mm/yyyy")
    TextBox6.Text = pos.Offset(0, 7).Value
    TextBox5.Text = pos.Offset(0, 6).Value
    TextBox8.Text = pos.Offset(0, 4).Value
    TextBox9.Text = pos.Value
    TextBox10.Text = pos.Offset(0, 2).Value
    TextBox11.Text = pos.Offset(0, 8).Value
End Sub

Private Sub opbFuncmaiorVencimento_Click()
    Dim ws As Worksheet, tabela As ListObject, linha As Single, id As String
    Set ws = ThisWorkbook.Sheets("Fábricas")
    Set tabela = ws.ListObjects(1)
    linha = Application.WorksheetFunction.Match(cmbFabricasFuncionarios.Value, tabela.ListColumns(2).DataBodyRange, 0)
    id = Application.WorksheetFunction.Index(tabela.ListColumns(3).DataBodyRange, linha)
    Dim ws2 As Worksheet, tbl As ListObject, coluna As Range, c As Range, venc As Single, pos As Range
    Set ws2 = ThisWorkbook.Sheets("Funcionários")
    Set tabela = ws2.ListObjects(1)
    Set coluna = tabela.ListColumns(3).DataBodyRange
    venc = 0
    For Each c In coluna
        If id = c.Value And c.Offset(0, 2).Value > venc Then
            venc = c.Offset(0, 2).Value
            Set pos = c.Offset(0, 1)
        End If
    Next c
    TextBox1.Text = pos.Offset(0, -2)
    TextBox2.Text = pos
    TextBox3.Text = pos.Offset(0, 8)
    TextBox4.Text = pos.Offset(0, 2)
    TextBox5.Text = pos.Offset(0, 5)
    TextBox6.Text = pos.Offset(0, 6)
    TextBox7.Text = pos.Offset(0, 4)
    TextBox8.Text = pos.Offset(0, 3)
    TextBox9.Text = pos.Offset(0, -1)
    TextBox10.Text = pos.Offset(0, 1)
    TextBox11.Text = pos.Offset(0, 7)
End Sub

Private Sub opbFuncmenosVenc_Click()
    Dim ws As Worksheet, tabela As ListObject, linha As Single, id As String
    Set ws = ThisWorkbook.Sheets("Fábricas")
    Set tabela = ws.ListObjects(1)
    linha = Application.WorksheetFunction.Match(cmbFabricasFuncionarios.Value, tabela.ListColumns(2).DataBodyRange, 0)
    id = Application.WorksheetFunction.Index(tabela.ListColumns(3).DataBodyRange, linha)
    Dim ws2 As Worksheet, tbl As ListObject, coluna As Range, c As Range, venc As Single, pos As Range
    Set ws2 = ThisWorkbook.Sheets("Funcionários")
    Set tabela = ws2.ListObjects(1)
    Set coluna = tabela.ListColumns(3).DataBodyRange
    venc = 9999
    For Each c In coluna
        If id = c.Value And c.Offset(0, 2).Value < venc Then
            venc = c.Offset(0, 2).Value
            Set pos = c.Offset(0, 1)
        End If
    Next c
    TextBox1.Text = pos.Offset(0, -2)
    TextBox2.Text = pos
    TextBox3.Text = pos.Offset(0, 8)
    TextBox4.Text = pos.Offset(0, 2)
    TextBox5.Text = pos.Offset(0, 5)
    TextBox6.Text = pos.Offset(0, 6)
    TextBox7.Text = pos.Offset(0, 4)
    TextBox8.Text = pos.Offset(0, 3)
    TextBox9.Text = pos.Offset(0, -1)
    TextBox10.Text = pos.Offset(0, 1)
    TextBox11.Text = pos.Offset(0, 7)
End Sub

Private Sub opbFuncNovo_Click()
    Dim ws As Worksheet, tabela As ListObject, linha As Single, id As String
    Set ws = ThisWorkbook.Sheets("Fábricas")
    Set tabela = ws.ListObjects(1)
    linha = Application.WorksheetFunction.Match(cmbFabricasFuncionarios.Value, tabela.ListColumns(2).DataBodyRange, 0)
    id = Application.WorksheetFunction.Index(tabela.ListColumns(3).DataBodyRange, linha)
    Dim ws2 As Worksheet, tbl As ListObject, coluna As Range, c As Range, novo As Integer, pos As Range
    Set ws2 = ThisWorkbook.Sheets("Funcionários")
    Set tabela = ws2.ListObjects(1)
    Set coluna = tabela.ListColumns(3).DataBodyRange
    novo = 100
    For Each c In coluna
        If id = c.Value And c.Offset(0, 7).Value < novo Then
            novo = c.Offset(0, 7).Value
            Set pos = c.Offset(0, 1)
        End If
    Next c
    TextBox1.Text = pos.Offset(0, -2)
    TextBox2.Text = pos
    TextBox3.Text = pos.Offset(0, 8)
    TextBox4.Text = pos.Offset(0, 2)
    TextBox5.Text = pos.Offset(0, 5)
    TextBox6.Text = pos.Offset(0, 6)
    TextBox7.Text = pos.Offset(0, 4)
    TextBox8.Text = pos.Offset(0, 3)
    TextBox9.Text = pos.Offset(0, -1)
    TextBox10.Text = pos.Offset(0, 1)
    TextBox11.Text = pos.Offset(0, 7)
End Sub

Private Sub opbFuncRecente_Click()
    Dim ws As Worksheet, tabela As ListObject, linha As Long, id As String
    Dim ws2 As Worksheet, tbl As ListObject, coluna As Range, c As Range
    Dim novo As Date, pos As Range

    Set ws = ThisWorkbook.Sheets("Fábricas")
    Set tabela = ws.ListObjects(1)
    linha = Application.WorksheetFunction.Match(cmbFabricasFuncionarios.Value, tabela.ListColumns(2).DataBodyRange, 0)
    id = Application.WorksheetFunction.Index(tabela.ListColumns(3).DataBodyRange, linha)
    Set ws2 = ThisWorkbook.Sheets("Funcionários")
    Set tbl = ws2.ListObjects(1)
    Set coluna = tbl.ListColumns(3).DataBodyRange

    novo = #1/1/1900#

    For Each c In coluna
        If id = c.Value And IsDate(c.Offset(0, 5).Value) And c.Offset(0, 5).Value > novo Then
            novo = c.Offset(0, 5).Value
            Set pos = c
        End If
    Next c

    TextBox1.Text = pos.Offset(0, -1).Value
    TextBox2.Text = pos.Offset(0, 1).Value
    TextBox3.Text = pos.Offset(0, 9).Value
    TextBox4.Text = pos.Offset(0, 3).Value
    TextBox7.Text = Format(pos.Offset(0, 5).Value, "dd/mm/yyyy")
    TextBox6.Text = pos.Offset(0, 7).Value
    TextBox5.Text = pos.Offset(0, 6).Value
    TextBox8.Text = pos.Offset(0, 4).Value
    TextBox9.Text = pos.Value
    TextBox10.Text = pos.Offset(0, 2).Value
    TextBox11.Text = pos.Offset(0, 8).Value
End Sub

Private Sub opbFuncVelho_Click()
    Dim ws As Worksheet, tabela As ListObject, linha As Single, id As String
    Set ws = ThisWorkbook.Sheets("Fábricas")
    Set tabela = ws.ListObjects(1)
    linha = Application.WorksheetFunction.Match(cmbFabricasFuncionarios.Value, tabela.ListColumns(2).DataBodyRange, 0)
    id = Application.WorksheetFunction.Index(tabela.ListColumns(3).DataBodyRange, linha)
    Dim ws2 As Worksheet, tbl As ListObject, coluna As Range, c As Range, velho As Integer, pos As Range
    Set ws2 = ThisWorkbook.Sheets("Funcionários")
    Set tabela = ws2.ListObjects(1)
    Set coluna = tabela.ListColumns(3).DataBodyRange
    velho = 0
    For Each c In coluna
        If id = c.Value And c.Offset(0, 7).Value > velho Then
            velho = c.Offset(0, 7).Value
            Set pos = c.Offset(0, 1)
        End If
    Next c
    TextBox1.Text = pos.Offset(0, -2)
    TextBox2.Text = pos
    TextBox3.Text = pos.Offset(0, 8)
    TextBox4.Text = pos.Offset(0, 2)
    TextBox5.Text = pos.Offset(0, 5)
    TextBox6.Text = pos.Offset(0, 6)
    TextBox7.Text = pos.Offset(0, 4)
    TextBox8.Text = pos.Offset(0, 3)
    TextBox9.Text = pos.Offset(0, -1)
    TextBox10.Text = pos.Offset(0, 1)
    TextBox11.Text = pos.Offset(0, 7)
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim fabrica As String
    Dim linha As Integer, ultLinha As Integer
    Set ws = ThisWorkbook.Sheets("Fábricas")
    ultLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For linha = 2 To ultLinha
        fabrica = ws.Cells(linha, 2).Value
        cmbFabricasFuncionarios.AddItem fabrica
    Next linha
End Sub



