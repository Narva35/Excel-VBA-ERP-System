VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstatisticasFabricas 
   Caption         =   "Estatísticas Fábricas"
   ClientHeight    =   6664
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10068
   OleObjectBlob   =   "frmEstatisticasFabricas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstatisticasFabricas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmbVoltar_Click()
  Unload Me
    frmEstatisticas.Show
End Sub

Private Sub lstEstatisticasFabricas_Click()
Dim valor As String
Dim ws As Worksheet, tabela As ListObject
Set ws = ThisWorkbook.Sheets("Fábricas")
Set tabela = ws.ListObjects(1)
valor = lstEstatisticasFabricas.Value
Select Case valor
    Case "Fábrica mais antiga"
Call clear
Dim data As Date, c As Range, coluna As Range, pos As Range
data = DateSerial(9999, 12, 31)
Set coluna = tabela.ListColumns(8).DataBodyRange
For Each c In coluna
    If data > c.Value Then
    data = c.Value
    Set pos = c
    End If
Next c
TextBox7.Text = pos
TextBox1.Text = pos.Offset(0, -6)
TextBox3.Text = pos.Offset(0, 1)
TextBox4.Text = pos.Offset(0, -5)
Dim ws2 As Worksheet, tabe As ListObject, linha As Long, lookup As Range
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = pos.Offset(0, -4)
TextBox11.Text = pos.Offset(0, -2)
    Case "Fábrica mais recente"
Call clear
data = #1/1/1900#
Dim posi As Range, cell As Range
Set coluna = tabela.ListColumns(8).DataBodyRange
For Each cell In coluna
    If data < cell.Value Then
    data = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi
TextBox1.Text = posi.Offset(0, -6)
TextBox3.Text = posi.Offset(0, 1)
TextBox4.Text = posi.Offset(0, -5)
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -4)
TextBox11.Text = posi.Offset(0, -2)
    Case "Fábrica com maior número de clientes"
Call clear
Dim num As Double
num = 0
Set coluna = tabela.ListColumns(5).DataBodyRange
For Each cell In coluna
    If num < cell.Value Then
    num = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi.Offset(0, 3)
TextBox1.Text = posi.Offset(0, -3)
TextBox3.Text = posi.Offset(0, 4)
TextBox4.Text = posi.Offset(0, -2)
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -1)
TextBox11.Text = posi.Offset(0, 1)
 Case "Fábrica com menor número de clientes"
Call clear
num = 999
Set coluna = tabela.ListColumns(5).DataBodyRange
For Each cell In coluna
    If num > cell.Value Then
    num = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi.Offset(0, 3)
TextBox1.Text = posi.Offset(0, -3)
TextBox3.Text = posi.Offset(0, 4)
TextBox4.Text = posi.Offset(0, -2)
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -1)
TextBox11.Text = posi.Offset(0, 1)
 Case "Fábrica com maior número de funcionários"
Call clear
num = 0
Set coluna = tabela.ListColumns(14).DataBodyRange
For Each cell In coluna
    If num < cell.Value Then
    num = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi.Offset(0, -6)
TextBox1.Text = posi.Offset(0, -12)
TextBox3.Text = posi.Offset(0, -5)
TextBox4.Text = posi.Offset(0, -11)
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -10)
TextBox11.Text = posi.Offset(0, -8)
 Case "Fábrica com menor número de funcionários"
Call clear
 num = 999
Set coluna = tabela.ListColumns(14).DataBodyRange
For Each cell In coluna
    If num > cell.Value Then
    num = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi.Offset(0, -6)
TextBox1.Text = posi.Offset(0, -12)
TextBox3.Text = posi.Offset(0, -5)
TextBox4.Text = posi.Offset(0, -11)
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -10)
TextBox11.Text = posi.Offset(0, -8)
 Case "Fábrica de maior área"
Call clear
num = 0
Set coluna = tabela.ListColumns(10).DataBodyRange
For Each cell In coluna
    If num < cell.Value Then
    num = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi.Offset(0, -2)
TextBox1.Text = posi.Offset(0, -8)
TextBox3.Text = posi.Offset(0, -1)
TextBox4.Text = posi.Offset(0, -7)
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -6)
TextBox11.Text = posi.Offset(0, -4)
 Case "Fábrica de menor área"
Call clear
num = 999999
Set coluna = tabela.ListColumns(10).DataBodyRange
For Each cell In coluna
    If cell.Value < num Then
    num = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi.Offset(0, -2).Value
TextBox1.Text = posi.Offset(0, -8).Value
TextBox3.Text = posi.Offset(0, -1).Value
TextBox4.Text = posi.Offset(0, -7).Value
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -6).Value
TextBox11.Text = posi.Offset(0, -4).Value
    Case "Fábrica com mais despesas"
Call clear
num = 0
Set coluna = tabela.ListColumns(11).DataBodyRange
For Each cell In coluna
    If num < cell.Value Then
    num = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi.Offset(0, -3)
TextBox1.Text = posi.Offset(0, -9)
TextBox3.Text = posi.Offset(0, -2)
TextBox4.Text = posi.Offset(0, -8)
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -7)
TextBox11.Text = posi.Offset(0, -5)
 Case "Fábrica com menos despesas"
Call clear
num = 9999
Set coluna = tabela.ListColumns(11).DataBodyRange
For Each cell In coluna
    If num > cell.Value Then
    num = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi.Offset(0, -3)
TextBox1.Text = posi.Offset(0, -9)
TextBox3.Text = posi.Offset(0, -2)
TextBox4.Text = posi.Offset(0, -8)
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -7)
TextBox11.Text = posi.Offset(0, -5)
Case "Fábrica com cliente mais antigo"
Call clear
Dim work As Worksheet, table As ListObject, linhas As Long, id As String
Dim novo As Date
    
Set work = ThisWorkbook.Sheets("Clientes")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(9).DataBodyRange
novo = DateSerial(9999, 12, 31)
    
For Each c In coluna
    If c.Value < novo Then
        novo = c.Value
        Set pos = c
    End If
Next c
Dim aux As Range
Set posi = pos.Offset(0, -5)
Set work = ThisWorkbook.Sheets("Fábricas")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(3).DataBodyRange
For Each c In coluna
    If posi = c.Value Then
    Set aux = c
    End If
Next c
TextBox1.Text = aux.Offset(0, -1).Value
TextBox3.Text = aux.Offset(0, 6).Value
TextBox4.Text = aux.Value
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox7.Text = aux.Offset(0, 5).Value
TextBox8.Text = aux.Offset(0, 1).Value
TextBox11.Text = aux.Offset(0, 4).Value
Case "Fábrica com cliente mais recente"
Call clear

    
Set work = ThisWorkbook.Sheets("Clientes")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(9).DataBodyRange
novo = DateSerial(1900, 1, 1)
    
For Each c In coluna
    If c.Value > novo Then
        novo = c.Value
        Set pos = c
    End If
Next c
Set posi = pos.Offset(0, -5)
Set work = ThisWorkbook.Sheets("Fábricas")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(3).DataBodyRange
For Each c In coluna
    If posi = c.Value Then
    Set aux = c
    End If
Next c
TextBox1.Text = aux.Offset(0, -1).Value
TextBox3.Text = aux.Offset(0, 6).Value
TextBox4.Text = aux.Value
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox7.Text = aux.Offset(0, 5).Value
TextBox8.Text = aux.Offset(0, 1).Value
TextBox11.Text = aux.Offset(0, 4).Value
Case "Fábrica com funcionário mais recente"
Call clear
  
Set work = ThisWorkbook.Sheets("Funcionários")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(8).DataBodyRange
novo = DateSerial(1900, 1, 1)
    
For Each c In coluna
    If c.Value > novo Then
        novo = c.Value
        Set pos = c
    End If
Next c
Set posi = pos.Offset(0, -5)
Set work = ThisWorkbook.Sheets("Fábricas")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(3).DataBodyRange
For Each c In coluna
    If posi = c.Value Then
    Set aux = c
    End If
Next c
TextBox1.Text = aux.Offset(0, -1).Value
TextBox3.Text = aux.Offset(0, 6).Value
TextBox4.Text = aux.Value
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox7.Text = aux.Offset(0, 5).Value
TextBox8.Text = aux.Offset(0, 1).Value
TextBox11.Text = aux.Offset(0, 4).Value
Case "Fábrica com funcionário mais antigo"
Call clear
  
Set work = ThisWorkbook.Sheets("Funcionários")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(8).DataBodyRange
novo = DateSerial(2030, 12, 31)
    
For Each c In coluna
    If c.Value < novo Then
        novo = c.Value
        Set pos = c
    End If
Next c
Set posi = pos.Offset(0, -5)
Set work = ThisWorkbook.Sheets("Fábricas")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(3).DataBodyRange
For Each c In coluna
    If posi = c.Value Then
    Set aux = c
    End If
Next c
TextBox1.Text = aux.Offset(0, -1).Value
TextBox3.Text = aux.Offset(0, 6).Value
TextBox4.Text = aux.Value
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox7.Text = aux.Offset(0, 5).Value
TextBox8.Text = aux.Offset(0, 1).Value
TextBox11.Text = aux.Offset(0, 4).Value
Case "Fábrica com maior faturação"
Call clear
num = 0
Set coluna = tabela.ListColumns(12).DataBodyRange
For Each cell In coluna
    If num < cell.Value Then
    num = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi.Offset(0, -4)
TextBox1.Text = posi.Offset(0, -10)
TextBox3.Text = posi.Offset(0, -3)
TextBox4.Text = posi.Offset(0, -9)
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -8)
TextBox11.Text = posi.Offset(0, -6)
Case "Fábrica com menor faturação"
Call clear
num = 99999
Set coluna = tabela.ListColumns(12).DataBodyRange
For Each cell In coluna
    If num > cell.Value Then
    num = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi.Offset(0, -4)
TextBox1.Text = posi.Offset(0, -10)
TextBox3.Text = posi.Offset(0, -3)
TextBox4.Text = posi.Offset(0, -9)
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -8)
TextBox11.Text = posi.Offset(0, -6)
 Case "Fábrica com maior capacidade de produção"
Call clear
num = 0
Set coluna = tabela.ListColumns(15).DataBodyRange
For Each cell In coluna
    If num < cell.Value Then
    num = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi.Offset(0, -7)
TextBox1.Text = posi.Offset(0, -13)
TextBox3.Text = posi.Offset(0, -6)
TextBox4.Text = posi.Offset(0, -12)
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -11)
TextBox11.Text = posi.Offset(0, -9)
Case "Fábrica com menor capacidade de produção"
Call clear
num = 99999
Set coluna = tabela.ListColumns(15).DataBodyRange
For Each cell In coluna
    If num > cell.Value Then
    num = cell.Value
    Set posi = cell
    End If
Next cell
TextBox7.Text = posi.Offset(0, -7)
TextBox1.Text = posi.Offset(0, -13)
TextBox3.Text = posi.Offset(0, -6)
TextBox4.Text = posi.Offset(0, -12)
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox8.Text = posi.Offset(0, -11)
TextBox11.Text = posi.Offset(0, -9)
Case "Fábrica com funcionário mais novo"
Call clear
  
Set work = ThisWorkbook.Sheets("Funcionários")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(9).DataBodyRange
novo = DateSerial(1900, 1, 1)
    
For Each c In coluna
    If c.Value > novo Then
        novo = c.Value
        Set pos = c
    End If
Next c
Set posi = pos.Offset(0, -6)
Set work = ThisWorkbook.Sheets("Fábricas")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(3).DataBodyRange
For Each c In coluna
    If posi = c.Value Then
    Set aux = c
    End If
Next c
TextBox1.Text = aux.Offset(0, -1).Value
TextBox3.Text = aux.Offset(0, 6).Value
TextBox4.Text = aux.Value
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox7.Text = aux.Offset(0, 5).Value
TextBox8.Text = aux.Offset(0, 1).Value
TextBox11.Text = aux.Offset(0, 3).Value
Case "Fábrica com funcionário mais velho"
Call clear
  
Set work = ThisWorkbook.Sheets("Funcionários")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(9).DataBodyRange
novo = DateSerial(2030, 12, 31)
    
For Each c In coluna
    If c.Value < novo Then
        novo = c.Value
        Set pos = c
    End If
Next c
Set posi = pos.Offset(0, -6)
Set work = ThisWorkbook.Sheets("Fábricas")
Set table = work.ListObjects(1)
Set coluna = table.ListColumns(3).DataBodyRange
For Each c In coluna
    If posi = c.Value Then
    Set aux = c
    End If
Next c
TextBox1.Text = aux.Offset(0, -1).Value
TextBox3.Text = aux.Offset(0, 6).Value
TextBox4.Text = aux.Value
Set ws2 = ThisWorkbook.Sheets("Funcionários")
Set tabe = ws2.ListObjects(1)
Set lookup = tabe.ListColumns(4).DataBodyRange
linha = Application.WorksheetFunction.Match(TextBox3.Value, lookup, 0)
TextBox5.Text = Application.WorksheetFunction.Index(tabe.ListColumns(2).DataBodyRange, linha)
TextBox7.Text = aux.Offset(0, 5).Value
TextBox8.Text = aux.Offset(0, 1).Value
TextBox11.Text = aux.Offset(0, 3).Value
End Select
End Sub

Private Sub UserForm_Initialize()
   
    Me.lstEstatisticasFabricas.AddItem "Fábrica mais antiga"
    Me.lstEstatisticasFabricas.AddItem "Fábrica mais recente"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com maior número de clientes"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com menor número de clientes"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com maior número de funcionários"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com menor número de funcionários"
    Me.lstEstatisticasFabricas.AddItem "Fábrica de maior área"
    Me.lstEstatisticasFabricas.AddItem "Fábrica de menor área"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com mais despesas"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com menos despesas"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com cliente mais antigo"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com cliente mais recente"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com funcionário mais recente"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com funcionário mais antigo"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com maior faturação"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com menor faturação"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com maior capacidade de produção"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com menor capacidade de produção"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com mais operadores de máquina"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com funcionário mais novo"
    Me.lstEstatisticasFabricas.AddItem "Fábrica com funcionário mais velho"
    
  
    
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
