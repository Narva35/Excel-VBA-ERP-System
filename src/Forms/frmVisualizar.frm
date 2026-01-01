VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVisualizar 
   Caption         =   "Visualizar"
   ClientHeight    =   11123
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17544
   OleObjectBlob   =   "frmVisualizar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVisualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim folha As String, original As Variant

Private Sub cmdFábricas_Click()
Dim c As Range
Call Module1.limpartxt
ListBox1.clear
Dim ws As Worksheet
Dim ulti As Long
Dim ende As String
Set ws = ThisWorkbook.Sheets("Fábricas")
ulti = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
If ws.Cells(ulti, "B").Value = "" Then
    ende = "B2:B" & ulti - 1
Else
    ende = "B2:B" & ulti
End If
With ListBox1
For Each c In Worksheets("Fábricas").Range(ende)
.AddItem c
Next c
End With
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label13.Visible = True
Label14.Visible = True
TextBox10.Visible = True
TextBox11.Visible = True
TextBox12.Visible = True
TextBox13.Visible = True
TextBox14.Visible = True


folha = "Fábricas"
Dim linha As Range, i As Integer
Set linha = Worksheets(folha).Rows(1)
For i = 1 To 14
Me.Controls("Label" & i).Caption = linha.Cells(1, i + 1).Value
Next i

End Sub

Private Sub cmdFechar_Click()
Call Module1.limpartxt
ListBox1.clear
Unload Me
frmPrincipal.Show
End Sub

Private Sub cmdFuncionários_Click()
Dim c As Range
Call Module1.limpartxt
ListBox1.clear
Dim ws As Worksheet
Dim ulti As Long
Dim ende As String
Set ws = ThisWorkbook.Sheets("Funcionários")
ulti = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
If ws.Cells(ulti, "B").Value = "" Then
    ende = "B2:B" & ulti - 1
Else
    ende = "B2:B" & ulti
End If
With ListBox1
For Each c In Worksheets("Funcionários").Range(ende)
.AddItem c
Next c
End With
Label10.Visible = True
Label11.Visible = True
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
TextBox10.Visible = True
TextBox11.Visible = True
TextBox12.Visible = False
TextBox13.Visible = False
TextBox14.Visible = False

folha = "Funcionários"
Dim linha As Range, i As Integer
Set linha = Worksheets(folha).Rows(1)
For i = 1 To 14
Me.Controls("Label" & i).Caption = linha.Cells(1, i + 1).Value
Next i

End Sub

Private Sub cmdClientes_Click()
Dim c As Range
Call Module1.limpartxt
ListBox1.clear
Dim ws As Worksheet
Dim ulti As Long
Dim ende As String
Set ws = ThisWorkbook.Sheets("Clientes")
ulti = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
If ws.Cells(ulti, "B").Value = "" Then
    ende = "B2:B" & ulti - 1
Else
    ende = "B2:B" & ulti
End If
With ListBox1
For Each c In Worksheets("Clientes").Range(ende)
.AddItem c
Next c
End With
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
TextBox11.Visible = False
TextBox12.Visible = False
TextBox13.Visible = False
TextBox14.Visible = False

folha = "Clientes"
Dim linha As Range, i As Integer
Set linha = Worksheets(folha).Rows(1)
For i = 1 To 14
Me.Controls("Label" & i).Caption = linha.Cells(1, i + 1).Value
Next i
End Sub

Private Sub cmdEncomendas_Click()
Dim c As Range
Call Module1.limpartxt
ListBox1.clear
Dim ws As Worksheet
Dim ulti As Long
Dim ende As String
Set ws = ThisWorkbook.Sheets("Encomendas")
ulti = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
If ws.Cells(ulti, "B").Value = "" Then
    ende = "B2:B" & ulti - 1
Else
    ende = "B2:B" & ulti
End If
With ListBox1
For Each c In Worksheets("Encomendas").Range(ende)
.AddItem c
Next c
End With
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label13.Visible = True
Label14.Visible = False

TextBox10.Visible = True
TextBox11.Visible = True
TextBox12.Visible = True
TextBox13.Visible = True
TextBox14.Visible = False


folha = "Encomendas"
Dim linha As Range, i As Integer
Set linha = Worksheets(folha).Rows(1)
For i = 1 To 14
Me.Controls("Label" & i).Caption = linha.Cells(1, i + 1).Value
Next i
End Sub

Private Sub cmdVisualizar()
frmVisualizar.Hide
frmPrincipal.Show

End Sub

    Private Sub ListBox1_Click()
        Dim tbl As ListObject
        Dim ws As Worksheet
        Dim lookup As Range, i As Integer
     Set ws = ThisWorkbook.Sheets(folha)
    With ws
     Set tbl = .ListObjects(1)
    End With
Set lookup = tbl.ListColumns(2).DataBodyRange

TextBox1.Text = Application.WorksheetFunction.Index(tbl.ListColumns(2).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox2.Text = Application.WorksheetFunction.Index(tbl.ListColumns(3).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox3.Text = Application.WorksheetFunction.Index(tbl.ListColumns(4).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox4.Text = Application.WorksheetFunction.Index(tbl.ListColumns(5).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox5.Text = Application.WorksheetFunction.Index(tbl.ListColumns(6).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox6.Text = Application.WorksheetFunction.Index(tbl.ListColumns(7).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox7.Text = Application.WorksheetFunction.Index(tbl.ListColumns(8).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox8.Text = Application.WorksheetFunction.Index(tbl.ListColumns(9).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox9.Text = Application.WorksheetFunction.Index(tbl.ListColumns(10).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox10.Text = Application.WorksheetFunction.Index(tbl.ListColumns(11).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox11.Text = Application.WorksheetFunction.Index(tbl.ListColumns(12).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
If tbl.ListColumns.Count > 12 Then
TextBox12.Text = Application.WorksheetFunction.Index(tbl.ListColumns(13).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox13.Text = Application.WorksheetFunction.Index(tbl.ListColumns(14).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
End If
If tbl.ListColumns.Count = 15 Then
TextBox12.Text = Application.WorksheetFunction.Index(tbl.ListColumns(13).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox13.Text = Application.WorksheetFunction.Index(tbl.ListColumns(14).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))
TextBox14.Text = Application.WorksheetFunction.Index(tbl.ListColumns(15).DataBodyRange, Application.WorksheetFunction.Match(ListBox1.Value, lookup, 0))

End If

 
    End Sub

Private Sub txtprocurar_Change()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim filterText As String

    filterText = Me.txtProcurar.Text
    Me.ListBox1.clear

    Set ws = ThisWorkbook.Sheets(folha)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For i = 1 To lastRow
        If InStr(1, ws.Cells(i + 1, 2).Value, filterText, vbTextCompare) > 0 Then
            Me.ListBox1.AddItem ws.Cells(i + 1, 2).Value
        End If
    Next i
End Sub

