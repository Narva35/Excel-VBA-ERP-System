VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstatisticasClientes2 
   Caption         =   "Estatisticas Clientes"
   ClientHeight    =   7455
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11580
   OleObjectBlob   =   "frmEstatisticasClientes2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstatisticasClientes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub lstEstatisticasClientes_Click()
Dim valor As String
Dim ws As Worksheet, tabela As ListObject
Set ws = ThisWorkbook.Sheets("Clientes")
Set tabela = ws.ListObjects(1)
valor = lstEstatisticasClientes.Value
Select Case valor
Case "Cliente mais antigo"
Call clear
Dim data As Date, c As Range, coluna As Range, pos As Range
data = DateSerial(9999, 12, 31)
Set coluna = tabela.ListColumns(9).DataBodyRange
For Each c In coluna
    If data > c.Value Then
    data = c.Value
    Set pos = c
    End If
Next c
TextBox7.Text = pos
TextBox1.Text = pos.Offset(0, -7)
TextBox3.Text = pos.Offset(0, -4)
TextBox4.Text = pos.Offset(0, -3)
TextBox5.Text = pos.Offset(0, -6)
TextBox8.Text = pos.Offset(0, -1)
TextBox9.Text = pos.Offset(0, -5)
TextBox11.Text = pos.Offset(0, 1)
Case "Cliente mais recente"
Call clear
data = DateSerial(1900, 12, 31)
Set coluna = tabela.ListColumns(9).DataBodyRange
For Each c In coluna
    If data < c.Value Then
    data = c.Value
    Set pos = c
    End If
Next c
TextBox7.Text = pos
TextBox1.Text = pos.Offset(0, -7)
TextBox3.Text = pos.Offset(0, -4)
TextBox4.Text = pos.Offset(0, -3)
TextBox5.Text = pos.Offset(0, -6)
TextBox8.Text = pos.Offset(0, -1)
TextBox9.Text = pos.Offset(0, -5)
TextBox11.Text = pos.Offset(0, 1)
Case "Cliente com melhor feedback"
Call clear
Dim a As Single
a = 0
Set coluna = tabela.ListColumns(11).DataBodyRange
For Each c In coluna
    If a < c.Value Then
    a = c.Value
    Set pos = c
    End If
Next c
TextBox7.Text = pos.Offset(0, -2)
TextBox1.Text = pos.Offset(0, -9)
TextBox3.Text = pos.Offset(0, -6)
TextBox4.Text = pos.Offset(0, -5)
TextBox5.Text = pos.Offset(0, -8)
TextBox8.Text = pos.Offset(0, -3)
TextBox9.Text = pos.Offset(0, -7)
TextBox11.Text = pos.Offset(0, -1)
Case "Cliente com pior feedback"
Call clear
a = 9999
Set coluna = tabela.ListColumns(11).DataBodyRange
For Each c In coluna
    If a > c.Value Then
    a = c.Value
    Set pos = c
    End If
Next c
TextBox7.Text = pos.Offset(0, -2)
TextBox1.Text = pos.Offset(0, -9)
TextBox3.Text = pos.Offset(0, -6)
TextBox4.Text = pos.Offset(0, -5)
TextBox5.Text = pos.Offset(0, -8)
TextBox8.Text = pos.Offset(0, -3)
TextBox9.Text = pos.Offset(0, -7)
TextBox11.Text = pos.Offset(0, -1)
Case "Cliente da encomenda com maior margem de lucro"
Call clear
Dim work As Worksheet, table As ListObject, novo As Single
Set work = ThisWorkbook.Sheets("Encomendas")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(14).DataBodyRange
novo = 0
Dim posi As Range, ws2 As Worksheet
For Each c In coluna
    If c.Value > novo Then
        novo = c.Value
        Set pos = c
    End If
Next c
Set posi = pos.Offset(0, -6)
Set work = ThisWorkbook.Sheets("Clientes")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(5).DataBodyRange
For Each c In coluna
    If posi = c.Value Then
    Set pos = c
    End If
Next c
TextBox7.Text = pos.Offset(0, 4)
TextBox1.Text = pos.Offset(0, -3)
TextBox3.Text = pos
TextBox4.Text = pos.Offset(0, 1)
TextBox5.Text = pos.Offset(0, -2)
TextBox8.Text = pos.Offset(0, 3)
TextBox9.Text = pos.Offset(0, -1)
TextBox11.Text = pos.Offset(0, 5)
Case "Cliente da encomenda com menor margem de lucro"
Call clear
Set work = ThisWorkbook.Sheets("Encomendas")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(14).DataBodyRange
novo = 9999
For Each c In coluna
    If c.Value < novo Then
        novo = c.Value
        Set pos = c
    End If
Next c
Set posi = pos.Offset(0, -6)
Set work = ThisWorkbook.Sheets("Clientes")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(5).DataBodyRange
For Each c In coluna
    If posi = c.Value Then
    Set pos = c
    End If
Next c
TextBox7.Text = pos.Offset(0, 4)
TextBox1.Text = pos.Offset(0, -3)
TextBox3.Text = pos
TextBox4.Text = pos.Offset(0, 1)
TextBox5.Text = pos.Offset(0, -2)
TextBox8.Text = pos.Offset(0, 3)
TextBox9.Text = pos.Offset(0, -1)
TextBox11.Text = pos.Offset(0, 5)
Case "Cliente da encomenda com maior duração"
Call clear
Set work = ThisWorkbook.Sheets("Encomendas")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(6).DataBodyRange
novo = 0
For Each c In coluna
    If c.Value > novo Then
        novo = c.Value
        Set pos = c
    End If
Next c
Set posi = pos.Offset(0, 2)
Set work = ThisWorkbook.Sheets("Clientes")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(5).DataBodyRange
For Each c In coluna
    If posi = c.Value Then
    Set pos = c
    End If
Next c
TextBox7.Text = pos.Offset(0, 4)
TextBox1.Text = pos.Offset(0, -3)
TextBox3.Text = pos
TextBox4.Text = pos.Offset(0, 1)
TextBox5.Text = pos.Offset(0, -2)
TextBox8.Text = pos.Offset(0, 3)
TextBox9.Text = pos.Offset(0, -1)
TextBox11.Text = pos.Offset(0, 5)
Case "Cliente da encomenda com menor duração"
Call clear
Set work = ThisWorkbook.Sheets("Encomendas")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(6).DataBodyRange
novo = 9999
For Each c In coluna
    If c.Value < novo Then
        novo = c.Value
        Set pos = c
    End If
Next c
Set posi = pos.Offset(0, 2)
Set ws2 = ThisWorkbook.Sheets("Clientes")
Set table = ws.ListObjects(1)
Set coluna = table.ListColumns(5).DataBodyRange
For Each c In coluna
    If posi = c.Value Then
    Set pos = c
    End If
Next c
TextBox7.Text = pos.Offset(0, 4)
TextBox1.Text = pos.Offset(0, -3)
TextBox3.Text = pos
TextBox4.Text = pos.Offset(0, 1)
TextBox5.Text = pos.Offset(0, -2)
TextBox8.Text = pos.Offset(0, 3)
TextBox9.Text = pos.Offset(0, -1)
TextBox11.Text = pos.Offset(0, 5)
Case "Cliente da encomenda com maior valor"
Call clear
Set work = ThisWorkbook.Sheets("Encomendas")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(7).DataBodyRange
novo = 0
For Each c In coluna
    If c.Value > novo Then
        novo = c.Value
        Set pos = c
    End If
Next c
Set posi = pos.Offset(0, 1)
Set work = ThisWorkbook.Sheets("Clientes")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(5).DataBodyRange
For Each c In coluna
    If posi = c.Value Then
    Set pos = c
    End If
Next c
TextBox7.Text = pos.Offset(0, 4)
TextBox1.Text = pos.Offset(0, -3)
TextBox3.Text = pos
TextBox4.Text = pos.Offset(0, 1)
TextBox5.Text = pos.Offset(0, -2)
TextBox8.Text = pos.Offset(0, 3)
TextBox9.Text = pos.Offset(0, -1)
TextBox11.Text = pos.Offset(0, 5)
Case "Cliente da encomenda com menor valor"
Call clear
Set work = ThisWorkbook.Sheets("Encomendas")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(7).DataBodyRange
novo = 9999
For Each c In coluna
    If c.Value < novo Then
        novo = c.Value
        Set pos = c
    End If
Next c
Set posi = pos.Offset(0, 1)
Set work = ThisWorkbook.Sheets("Clientes")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(5).DataBodyRange
For Each c In coluna
    If posi = c.Value Then
    Set pos = c
    End If
Next c
TextBox7.Text = pos.Offset(0, 4)
TextBox1.Text = pos.Offset(0, -3)
TextBox3.Text = pos
TextBox4.Text = pos.Offset(0, 1)
TextBox5.Text = pos.Offset(0, -2)
TextBox8.Text = pos.Offset(0, 3)
TextBox9.Text = pos.Offset(0, -1)
TextBox11.Text = pos.Offset(0, 5)
End Select
End Sub


Private Sub cmbVoltar_Click()
  Unload Me
    frmEstatisticas.Show
End Sub

Private Sub UserForm_Initialize()
   
    Me.lstEstatisticasClientes.AddItem "Cliente mais antigo"
    Me.lstEstatisticasClientes.AddItem "Cliente mais recente"
    Me.lstEstatisticasClientes.AddItem "Cliente com melhor feedback"
    Me.lstEstatisticasClientes.AddItem "Cliente com pior feedback"
    Me.lstEstatisticasClientes.AddItem "Cliente da encomenda com maior margem de lucro"
    Me.lstEstatisticasClientes.AddItem "Cliente da encomenda com menor margem de lucro"
    Me.lstEstatisticasClientes.AddItem "Cliente da encomenda com maior duração"
    Me.lstEstatisticasClientes.AddItem "Cliente da encomenda com menor duração"
    Me.lstEstatisticasClientes.AddItem "Cliente da encomenda com maior valor"
    Me.lstEstatisticasClientes.AddItem "Cliente da encomenda com menor valor"
    
End Sub

Public Sub clear()
TextBox1.Text = ""
TextBox3.Text = ""
TextBox4.Text = ""
TextBox5.Text = ""
TextBox7.Text = ""
TextBox8.Text = ""
TextBox11.Text = ""
End Sub
Public Sub ola()
TextBox7.Text = pos.Offset(0, 4)
TextBox1.Text = pos.Offset(0, -3)
TextBox3.Text = pos
TextBox4.Text = pos.Offset(0, 1)
TextBox5.Text = pos.Offset(0, -2)
TextBox8.Text = pos.Offset(0, 3)
TextBox9.Text = pos.Offset(0, -1)
TextBox11.Text = pos.Offset(0, 5)
End Sub

