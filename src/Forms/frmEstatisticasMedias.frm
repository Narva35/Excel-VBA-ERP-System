VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstatisticasMedias 
   Caption         =   "UserForm1"
   ClientHeight    =   10731
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9456.001
   OleObjectBlob   =   "frmEstatisticasMedias.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstatisticasMedias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub lstbMedias_Initialize()

lstbMedias.AddItem "Todos os Países"
lstbMedias.AddItem "Número de Funcionários"
lstbMedias.AddItem "Número de Encomendas"
lstbMedias.AddItem "Número de Clientes"

End Sub


Private Sub cmbVoltar_Click()
  Unload Me
    frmEstatisticas.Show
End Sub

Private Sub cmdEstMediaFuncFab_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Fábricas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(14).DataBodyRange
media = 0
soma = 0
i = 0
For Each c In coluna
    soma = soma + c.Value
    i = i + 1
Next c
media = soma / i

txbRespostaEstMedias.Value = FormatNumber(media, 0) & " Funcionários"
End Sub

Private Sub cmdMedAreaFabrica_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Fábricas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(10).DataBodyRange
media = 0
soma = 0
i = 0
For Each c In coluna
    soma = soma + c.Value
    i = i + 1
Next c
media = soma / i

txbRespostaEstMedias.Value = FormatNumber(media, 1) & " Metros quadrados"
End Sub

Private Sub cmdMedCapacidadeProducao_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Fábricas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(15).DataBodyRange
media = 0
soma = 0
i = 0
For Each c In coluna
    soma = soma + c.Value
    i = i + 1
Next c
media = soma / i

txbRespostaEstMedias.Value = FormatNumber(media, 2) & " Toneladas"
End Sub

Private Sub cmdMedDespesasFabrica_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Fábricas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(11).DataBodyRange
media = 0
soma = 0
i = 0
For Each c In coluna
    soma = soma + c.Value
    i = i + 1
Next c
media = soma / i

txbRespostaEstMedias.Value = FormatNumber(media, 2) & " Milhões de Euros"
End Sub

Private Sub cmdMedFeedback_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Clientes")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(11).DataBodyRange
media = 0
soma = 0
i = 0
For Each c In coluna
    soma = soma + c.Value
    i = i + 1
Next c
media = soma / i

txbRespostaEstMedias.Value = FormatNumber(media, 1)
End Sub

Private Sub cmdMediaFaturacaoFabrica_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Fábricas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(12).DataBodyRange
media = 0
soma = 0
i = 0
For Each c In coluna
    soma = soma + c.Value
    i = i + 1
Next c
media = soma / i

txbRespostaEstMedias.Value = FormatNumber(media, 2) & " Milhões de Euros"
End Sub

Private Sub cmdMediaSalarioDiretor_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Funcionários")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(11).DataBodyRange
soma = 0
i = 0
For Each c In coluna
    If c.Value = "Diretor" Then
        soma = soma + c.Offset(0, -6).Value
        i = i + 1
    End If
Next c
media = soma / i
txbRespostaEstMedias.Value = FormatNumber(media, 2) & " Euros"
End Sub

Private Sub cmdMedIdadeFunc_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Funcionários")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(10).DataBodyRange
media = 0
soma = 0
i = 0
For Each c In coluna
    soma = soma + c.Value
    i = i + 1
Next c
media = soma / i

txbRespostaEstMedias.Value = FormatNumber(media, 1) & " Anos"
End Sub

Private Sub cmdMedIdadeGestores_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Funcionários")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(11).DataBodyRange
soma = 0
i = 0
For Each c In coluna
    If c.Value = "Diretor" Then
        soma = soma + c.Offset(0, -1).Value
        i = i + 1
    End If
Next c
media = soma / i
txbRespostaEstMedias.Value = FormatNumber(media, 1) & " Anos"
End Sub

Private Sub cmdMedMargemLucro_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Encomendas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(14).DataBodyRange
media = 0
soma = 0
i = 0
For Each c In coluna
    soma = soma + c.Value
    i = i + 1
Next c
media = soma / i

txbRespostaEstMedias.Value = FormatNumber(media, 1) & " Euros"
End Sub

Private Sub cmdMedOperadorFabrica_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Funcionários")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(11).DataBodyRange
soma = 0
i = 0
For Each c In coluna
    If c.Value = "Operador de Máquina" Then
        soma = soma + c.Offset(0, -6).Value
        i = i + 1
    End If
Next c
media = soma / i
txbRespostaEstMedias.Value = FormatNumber(media, 2) & " Euros"
End Sub

Private Sub cmdMedSalarioEngenheiro_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Funcionários")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(11).DataBodyRange
soma = 0
i = 0
For Each c In coluna
    If c.Value = "Engenheiro" Then
        soma = soma + c.Offset(0, -6).Value
        i = i + 1
    End If
Next c
media = soma / i
txbRespostaEstMedias.Value = FormatNumber(media, 2) & " Euros"
End Sub

Private Sub cmdMedSalarioGestores_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Funcionários")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(11).DataBodyRange
soma = 0
i = 0
For Each c In coluna
    If c.Value = "Gestor" Then
        soma = soma + c.Offset(0, -6).Value
        i = i + 1
    End If
Next c
media = soma / i
txbRespostaEstMedias.Value = FormatNumber(media, 2) & " Euros"
End Sub

Private Sub cmdMedTempEncomenda_Click()
Dim ws As Worksheet, tbl As ListObject, coluna As Range, media As Single, c As Range, soma As Single, i As Integer
Set ws = ThisWorkbook.Sheets("Encomendas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(6).DataBodyRange
media = 0
soma = 0
i = 0
For Each c In coluna
    soma = soma + c.Value
    i = i + 1
Next c
media = soma / i

txbRespostaEstMedias.Value = FormatNumber(media, 1) & " Dias"
End Sub


