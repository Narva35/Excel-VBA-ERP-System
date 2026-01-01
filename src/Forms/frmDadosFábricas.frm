VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDadosFábricas 
   Caption         =   "Dados Fábricas"
   ClientHeight    =   11039
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14844
   OleObjectBlob   =   "frmDadosFábricas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDadosFábricas"
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

    Set ws = ThisWorkbook.Sheets("Fábricas")
    With ws
        Set tbl = .ListObjects(1)
    End With
    Set lookup = tbl.ListColumns(2).DataBodyRange

    linha = Application.WorksheetFunction.Match(lstFábricas.Value, lookup, 0)

    If Not IsNumeric(Me.TextBox4.Text) Then
        MsgBox "Por favor, insira um valor numérico de clientes.", vbExclamation
        TextBox4.SetFocus
        TextBox4.Text = ""
        Exit Sub
    End If

    If Not IsNumeric(Me.TextBox9.Text) Then
        MsgBox "Por favor, insira um valor numérico de área.", vbExclamation
        TextBox9.SetFocus
        TextBox9.Text = ""
        Exit Sub
    End If

    If Not IsNumeric(Me.TextBox10.Text) Then
        MsgBox "Por favor, insira um valor numérico de despesas.", vbExclamation
        TextBox10.SetFocus
        TextBox10.Text = ""
        Exit Sub
    End If

    If Not IsNumeric(Me.TextBox11.Text) Then
        MsgBox "Por favor, insira um valor numérico de faturação.", vbExclamation
        TextBox11.SetFocus
        TextBox11.Text = ""
        Exit Sub
    End If

    If Not IsNumeric(Me.TextBox12.Text) Then
        MsgBox "Por favor, insira um valor numérico de resultado líquido.", vbExclamation
        TextBox12.SetFocus
        TextBox12.Text = ""
        Exit Sub
    End If

    If Not IsNumeric(Me.TextBox13.Text) Then
        MsgBox "Por favor, insira um valor numérico de funcionários.", vbExclamation
        TextBox13.SetFocus
        TextBox13.Text = ""
        Exit Sub
    End If

    If Not IsNumeric(Me.TextBox14.Text) Then
        MsgBox "Por favor, insira um valor numérico de capacidade de produção anual.", vbExclamation
        TextBox14.SetFocus
        TextBox14.Text = ""
        Exit Sub
    End If

    For i = 1 To 14
        If i < 9 Then
            Set cell = tbl.ListColumns(i + 1).DataBodyRange.Cells(linha)
            cell.Value = Me.Controls("TextBox" & i).Text
        ElseIf i = 9 Or i = 10 Or i = 11 Or i = 13 Or i = 14 Then
            Set cell = tbl.ListColumns(i + 1).DataBodyRange.Cells(linha)
            cell.Value = Me.Controls("TextBox" & i).Text
            cell.Value = CDbl(cell.Value)
        End If
    Next i
    
    MsgBox "Dados alterados com sucesso!", vbInformation
    
    Call clear
    lstFábricas.clear

    Dim linha1 As Range
    Set linha1 = Worksheets("Fábricas").Rows(1)
    Dim c As Range
    lstFábricas.clear
    Dim ulti As Long
    Dim ende As String
    ulti = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    If ws.Cells(ulti, "B").Value = "" Then
        ende = "B2:B" & ulti - 1
    Else
        ende = "B2:B" & ulti
    End If

    With lstFábricas
        For Each c In Worksheets("Fábricas").Range(ende)
            .AddItem c
        Next c
    End With
End Sub


Private Sub cmdVoltar_Click()
Unload Me
frmFábricas.Show
End Sub


Private Sub lstFábricas_Click()
Dim tbl As ListObject, i As Integer
        Dim ws As Worksheet
        Dim lookup As Range, linha As Long
        Dim cell As String
     Set ws = ThisWorkbook.Sheets("Fábricas")
    With ws
     Set tbl = .ListObjects(1)
    End With
Set lookup = tbl.ListColumns(2).DataBodyRange
linha = Application.WorksheetFunction.Match(lstFábricas.Value, lookup, 0)

For i = 1 To 14
Me.Controls("TextBox" & i).Text = Application.WorksheetFunction.Index(tbl.ListColumns(i + 1).DataBodyRange, linha)
Next i
End Sub

Private Sub UserForm_Initialize()
Dim ws As Worksheet, i As Integer, linha As Range
Set ws = ThisWorkbook.Sheets("Fábricas")
Set linha = Worksheets("Fábricas").Rows(1)
Dim c As Range
lstFábricas.clear
Dim ulti As Long
Dim ende As String
ulti = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
If ws.Cells(ulti, "B").Value = "" Then
    ende = "B2:B" & ulti - 1
Else
    ende = "B2:B" & ulti
End If
With lstFábricas
For Each c In Worksheets("Fábricas").Range(ende)
.AddItem c
Next c
End With
For i = 1 To 14
Me.Controls("lbl" & i).Caption = linha.Cells(1, i + 1).Value
Next i
End Sub


Public Sub clear()
Dim i As Integer
For i = 1 To 14
Me.Controls("TextBox" & i).Text = ""
Next i
lstFábricas.ListIndex = -1
End Sub
