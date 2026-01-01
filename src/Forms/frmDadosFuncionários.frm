VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDadosFuncionários 
   Caption         =   "Dados Funcionários"
   ClientHeight    =   10682
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15252
   OleObjectBlob   =   "frmDadosFuncionários.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDadosFuncionários"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlterar_Click()
    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim lookup As Range
    Dim linha As Long
    Dim cell As Range

    Set ws = ThisWorkbook.Sheets("Funcionários")
    With ws
        Set tbl = .ListObjects(1)
    End With
    Set lookup = tbl.ListColumns(2).DataBodyRange

    linha = Application.WorksheetFunction.Match(lstFuncionários.Value, lookup, 0)

    If Not IsNumeric(Me.TextBox4.Text) Then
        MsgBox "Por favor, insira um valor numérico de vencimento.", vbExclamation
        TextBox4.SetFocus
        TextBox4.Text = ""
        Exit Sub
    End If

    If Not IsNumeric(Me.TextBox5.Text) Then
        MsgBox "Por favor, insira um NIF válido (números).", vbExclamation
        TextBox5.SetFocus
        TextBox5.Text = ""
        Exit Sub
    End If

    If Not IsNumeric(Me.TextBox9.Text) Then
        MsgBox "Por favor, insira um valor numérico de idade.", vbExclamation
        TextBox9.SetFocus
        TextBox9.Text = ""
        Exit Sub
    End If

    ' Verificações de formato de data
    If Not VerificarFormatoData(Me.TextBox7.Text) Then
        MsgBox "Por favor, insira uma data de admissão no formato dd/mm/aaaa.", vbExclamation
        TextBox7.SetFocus
        TextBox7.Text = ""
        Exit Sub
    End If

    If Not VerificarFormatoData(Me.TextBox8.Text) Then
        MsgBox "Por favor, insira uma data de saída no formato dd/mm/aaaa.", vbExclamation
        TextBox8.SetFocus
        TextBox8.Text = ""
        Exit Sub
    End If

    For i = 1 To 10
    If i <> 4 Or i <> 9 Then
    Set cell = tbl.ListColumns(i + 1).DataBodyRange.Cells(linha)
    cell.Value = Me.Controls("TextBox" & i).Text
    ElseIf i = 4 Then
    Set cell = tbl.ListColumns(5).DataBodyRange.Cells(linha)
    cell.Value = TextBox4.Text
    cell.Value = CDbl(cell.Value)
    cell.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    End If
Next i
    
    MsgBox "Dados alterados com sucesso!", vbInformation
    
    Call clear
    lstFuncionários.clear

    Dim linha1 As Range
    Set linha1 = Worksheets("Funcionários").Rows(1)
    Dim c As Range
    lstFuncionários.clear
    Dim ulti As Long
    Dim ende As String
    ulti = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    If ws.Cells(ulti, "B").Value = "" Then
        ende = "B2:B" & ulti - 1
    Else
        ende = "B2:B" & ulti
    End If

    With lstFuncionários
        For Each c In Worksheets("Funcionários").Range(ende)
            .AddItem c
        Next c
    End With
End Sub


Private Sub cmdVoltar_Click()
Unload Me
frmFuncionários.Show
End Sub

Private Sub ComboBox1_Change()
ComboBox1.List = Worksheets("Funcionários").Range("K2:K6").Value
End Sub

Private Sub lstFuncionários_Click()
Dim tbl As ListObject, i As Integer
        Dim ws As Worksheet
        Dim lookup As Range, linha As Long
        Dim cell As String
     Set ws = ThisWorkbook.Sheets("Funcionários")
    With ws
     Set tbl = .ListObjects(1)
    End With
Set lookup = tbl.ListColumns(2).DataBodyRange
linha = Application.WorksheetFunction.Match(lstFuncionários.Value, lookup, 0)

For i = 1 To 9
Me.Controls("TextBox" & i).Text = Application.WorksheetFunction.Index(tbl.ListColumns(i + 1).DataBodyRange, linha)
Next i
TextBox10.Text = Application.WorksheetFunction.Index(tbl.ListColumns(12).DataBodyRange, linha)
ComboBox1.Text = Application.WorksheetFunction.Index(tbl.ListColumns(11).DataBodyRange, linha)
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

Private Sub UserForm_Initialize()
Dim ws As Worksheet, i As Integer, linha As Range
Set ws = ThisWorkbook.Sheets("Funcionários")
Set linha = Worksheets("Funcionários").Rows(1)
Dim c As Range
lstFuncionários.clear
Dim ulti As Long
Dim ende As String
ulti = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
If ws.Cells(ulti, "B").Value = "" Then
    ende = "B2:B" & ulti - 1
Else
    ende = "B2:B" & ulti
End If
With lstFuncionários
For Each c In Worksheets("Funcionários").Range(ende)
.AddItem c
Next c
End With
For i = 1 To 11
Me.Controls("lbl" & i).Caption = linha.Cells(1, i + 1).Value
Next i
End Sub

Public Sub clear()
Dim i As Integer
For i = 1 To 10
Me.Controls("TextBox" & i).Text = ""
Next i
ComboBox1.Text = ""
txtProcurar.Text = ""
lstFuncionários.ListIndex = -1
End Sub

