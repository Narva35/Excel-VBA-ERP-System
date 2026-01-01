VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstatisticasEncomendas 
   Caption         =   "Estatísticas Encomendas"
   ClientHeight    =   6734
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9336.001
   OleObjectBlob   =   "frmEstatisticasEncomendas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstatisticasEncomendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmbVoltar_Click()
  Unload Me
    frmEstatisticas.Show
End Sub



Private Sub cmdEstEncMaisRecente_Click()
Dim ws As Worksheet, data As Date, c As Range, tbl As ListObject, coluna As Range, pos As Range
Set ws = ThisWorkbook.Sheets("Encomendas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(3).DataBodyRange
data = DateSerial(1900, 1, 1)
For Each c In coluna
    If c.Value > data Then
    data = c.Value
    Set pos = c
    End If
Next c
TextBox1.Text = pos.Offset(0, -1).Value
TextBox3.Text = pos.Offset(0, 5).Value
TextBox4.Text = pos.Offset(0, -2).Value
TextBox5.Text = pos.Value
TextBox8.Text = pos.Offset(0, 1).Value
TextBox7.Text = pos.Offset(0, 2).Value
TextBox9.Text = pos.Offset(0, 4).Value
TextBox10.Text = pos.Offset(0, 10).Value
TextBox11.Text = pos.Offset(0, 11).Value
End Sub

Private Sub CommandButton1_Click()
Dim ws As Worksheet, data As Date, c As Range, tbl As ListObject, coluna As Range, pos As Range
Set ws = ThisWorkbook.Sheets("Encomendas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(3).DataBodyRange
data = DateSerial(2030, 1, 1)
For Each c In coluna
    If c.Value < data Then
    data = c.Value
    Set pos = c
    End If
Next c
TextBox1.Text = pos.Offset(0, -1).Value
TextBox3.Text = pos.Offset(0, 5).Value
TextBox4.Text = pos.Offset(0, -2).Value
TextBox5.Text = pos.Value
TextBox8.Text = pos.Offset(0, 1).Value
TextBox7.Text = pos.Offset(0, 2).Value
TextBox9.Text = pos.Offset(0, 4).Value
TextBox10.Text = pos.Offset(0, 10).Value
TextBox11.Text = pos.Offset(0, 11).Value
End Sub

Private Sub CommandButton2_Click()
Dim ws As Worksheet, data As Date, c As Range, tbl As ListObject, coluna As Range, pos As Range
Set ws = ThisWorkbook.Sheets("Encomendas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(4).DataBodyRange
data = DateSerial(1900, 1, 1)
For Each c In coluna
    If c.Value > data Then
    data = c.Value
    Set pos = c
    End If
Next c
TextBox1.Text = pos.Offset(0, -2).Value
TextBox3.Text = pos.Offset(0, 4).Value
TextBox4.Text = pos.Offset(0, -3).Value
TextBox5.Text = pos.Offset(0, -1).Value
TextBox8.Text = pos.Value
TextBox7.Text = pos.Offset(0, 1).Value
TextBox9.Text = pos.Offset(0, 3).Value
TextBox10.Text = pos.Offset(0, 9).Value
TextBox11.Text = pos.Offset(0, 10).Value
End Sub

Private Sub CommandButton3_Click()
Dim ws As Worksheet, data As Date, c As Range, tbl As ListObject, coluna As Range, pos As Range
Set ws = ThisWorkbook.Sheets("Encomendas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(4).DataBodyRange
data = DateSerial(2030, 1, 1)
For Each c In coluna
    If c.Value < data Then
    data = c.Value
    Set pos = c
    End If
Next c
TextBox1.Text = pos.Offset(0, -2).Value
TextBox3.Text = pos.Offset(0, 4).Value
TextBox4.Text = pos.Offset(0, -3).Value
TextBox5.Text = pos.Offset(0, -1).Value
TextBox8.Text = pos.Value
TextBox7.Text = pos.Offset(0, 1).Value
TextBox9.Text = pos.Offset(0, 3).Value
TextBox10.Text = pos.Offset(0, 9).Value
TextBox11.Text = pos.Offset(0, 10).Value
End Sub

Private Sub CommandButton4_Click()
Dim ws As Worksheet, data As Date, c As Range, tbl As ListObject, coluna As Range, pos As Range
Set ws = ThisWorkbook.Sheets("Encomendas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(5).DataBodyRange
data = DateSerial(1900, 1, 1)
For Each c In coluna
    If c.Value > data Then
    data = c.Value
    Set pos = c
    End If
Next c
TextBox1.Text = pos.Offset(0, -3).Value
TextBox3.Text = pos.Offset(0, 3).Value
TextBox4.Text = pos.Offset(0, -4).Value
TextBox5.Text = pos.Offset(0, -2).Value
TextBox8.Text = pos.Offset(0, -1).Value
TextBox7.Text = pos.Value
TextBox9.Text = pos.Offset(0, 2).Value
TextBox10.Text = pos.Offset(0, 8).Value
TextBox11.Text = pos.Offset(0, 9).Value
End Sub

Private Sub CommandButton5_Click()
Dim ws As Worksheet, data As Date, c As Range, tbl As ListObject, coluna As Range, pos As Range
Set ws = ThisWorkbook.Sheets("Encomendas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(5).DataBodyRange
data = DateSerial(2030, 1, 1)
For Each c In coluna
    If c.Value < data Then
    data = c.Value
    Set pos = c
    End If
Next c
TextBox1.Text = pos.Offset(0, -3).Value
TextBox3.Text = pos.Offset(0, 3).Value
TextBox4.Text = pos.Offset(0, -4).Value
TextBox5.Text = pos.Offset(0, -2).Value
TextBox8.Text = pos.Offset(0, -1).Value
TextBox7.Text = pos.Value
TextBox9.Text = pos.Offset(0, 2).Value
TextBox10.Text = pos.Offset(0, 8).Value
TextBox11.Text = pos.Offset(0, 9).Value
End Sub

Private Sub CommandButton6_Click()
Dim ws As Worksheet, valor As Single, c As Range, tbl As ListObject, coluna As Range, pos As Range
Set ws = ThisWorkbook.Sheets("Encomendas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(14).DataBodyRange
valor = 0
For Each c In coluna
    If c.Value > valor Then
    valor = c.Value
    Set pos = c
    End If
Next c
TextBox1.Text = pos.Offset(0, -12).Value
TextBox3.Text = pos.Offset(0, -6).Value
TextBox4.Text = pos.Offset(0, -13).Value
TextBox5.Text = pos.Offset(0, -11).Value
TextBox8.Text = pos.Offset(0, -10).Value
TextBox7.Text = pos.Offset(0, -9).Value
TextBox9.Text = pos.Offset(0, -7).Value
TextBox10.Text = pos.Offset(0, -1).Value
TextBox11.Text = pos.Value
End Sub

Private Sub CommandButton7_Click()
Dim ws As Worksheet, valor As Single, c As Range, tbl As ListObject, coluna As Range, pos As Range
Set ws = ThisWorkbook.Sheets("Encomendas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(14).DataBodyRange
valor = 999999
For Each c In coluna
    If c.Value < valor Then
    valor = c.Value
    Set pos = c
    End If
Next c
TextBox1.Text = pos.Offset(0, -12).Value
TextBox3.Text = pos.Offset(0, -6).Value
TextBox4.Text = pos.Offset(0, -13).Value
TextBox5.Text = pos.Offset(0, -11).Value
TextBox8.Text = pos.Offset(0, -10).Value
TextBox7.Text = pos.Offset(0, -9).Value
TextBox9.Text = pos.Offset(0, -7).Value
TextBox10.Text = pos.Offset(0, -1).Value
TextBox11.Text = pos.Value
End Sub

