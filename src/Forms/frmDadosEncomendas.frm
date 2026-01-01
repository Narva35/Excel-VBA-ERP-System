VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDadosEncomendas 
   Caption         =   "Dados Encomendas"
   ClientHeight    =   10444
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15276
   OleObjectBlob   =   "frmDadosEncomendas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDadosEncomendas"
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
    Dim DataValida As Boolean

    Set ws = ThisWorkbook.Sheets("Encomendas")
    With ws
        Set tbl = .ListObjects(1)
    End With
    Set lookup = tbl.ListColumns(2).DataBodyRange

    linha = Application.WorksheetFunction.Match(lstEncomendas.Value, lookup, 0)

    If Not VerificarFormatoData(Me.TextBox2.Text) Then
        MsgBox "Por favor, insira uma data de compra no formato dd/mm/aaaa.", vbExclamation
        TextBox2.Text = ""
        TextBox2.SetFocus
        Exit Sub
    End If
    
    If Not VerificarFormatoData(Me.TextBox3.Text) Then
        MsgBox "Por favor, insira uma data de envio no formato dd/mm/aaaa.", vbExclamation
        TextBox3.SetFocus
        TextBox3.Text = ""
        Exit Sub
    End If
    
    If Not VerificarFormatoData(Me.TextBox4.Text) Then
        MsgBox "Por favor, insira uma data de chegada no formato dd/mm/aaaa.", vbExclamation
        TextBox4.SetFocus
        TextBox4.Text = ""
        Exit Sub
    End If

    If Not IsNumeric(Me.TextBox9.Text) Then
        MsgBox "Por favor, insira um valor numérico de unidades.", vbExclamation
        TextBox9.SetFocus
        TextBox9.Text = ""
        Exit Sub
    End If

    If Not IsNumeric(Me.TextBox10.Text) Then
        MsgBox "Por favor, insira um valor numérico de custo.", vbExclamation
        TextBox10.SetFocus
        TextBox10.Text = ""
        Exit Sub
    End If

    If Not IsNumeric(Me.TextBox11.Text) Or Me.TextBox11.Text < 0 Or Me.TextBox11.Text > 1 Then
        MsgBox "Por favor, insira um valor numérico de IVA entre 0 e 1.", vbExclamation
        TextBox11.SetFocus
        TextBox11.Text = ""
        Exit Sub
    End If

    For i = 1 To 13
        If i <> 11 And i <> 10 And i <> 5 And i <> 6 And i <> 12 And i <> 13 Then
            Set cell = tbl.ListColumns(i + 1).DataBodyRange.Cells(linha)
            cell.Value = Me.Controls("TextBox" & i).Text
        ElseIf i = 10 Then
            Set cell = tbl.ListColumns(11).DataBodyRange.Cells(linha)
            cell.Value = TextBox10.Text
            cell.Value = CDbl(cell.Value)
        ElseIf i = 11 Then
            Set cell = tbl.ListColumns(12).DataBodyRange.Cells(linha)
            cell.Value = TextBox11.Text
            cell.Value = CDbl(cell.Value)
        End If
    Next i
    
    MsgBox "Dados alterados com sucesso!", vbInformation
    
    Call clear
    lstEncomendas.clear

    Dim linha1 As Range
    Set linha1 = Worksheets("Encomendas").Rows(1)
    Dim c As Range
    lstEncomendas.clear
    Dim ulti As Long
    Dim ende As String
    ulti = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    If ws.Cells(ulti, "B").Value = "" Then
        ende = "B2:B" & ulti - 1
    Else
        ende = "B2:B" & ulti
    End If

    With lstEncomendas
        For Each c In Worksheets("Encomendas").Range(ende)
            .AddItem c
        Next c
    End With
End Sub

Private Sub cmdVoltar_Click()
Unload Me
frmEncomendas.Show
End Sub

Private Sub lstEncomendas_Click()
Dim tbl As ListObject, i As Integer
        Dim ws As Worksheet
        Dim lookup As Range, linha As Long
        Dim cell As String
     Set ws = ThisWorkbook.Sheets("Encomendas")
    With ws
     Set tbl = .ListObjects(1)
    End With
Set lookup = tbl.ListColumns(2).DataBodyRange
linha = Application.WorksheetFunction.Match(lstEncomendas.Value, lookup, 0)

For i = 1 To 13
Me.Controls("TextBox" & i).Text = Application.WorksheetFunction.Index(tbl.ListColumns(i + 1).DataBodyRange, linha)
Next i
End Sub

Private Sub UserForm_Initialize()
Dim ws As Worksheet, i As Integer, linha As Range
Set ws = ThisWorkbook.Sheets("Encomendas")
Set linha = Worksheets("Encomendas").Rows(1)
Dim c As Range
lstEncomendas.clear
Dim ulti As Long
Dim ende As String
ulti = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
If ws.Cells(ulti, "B").Value = "" Then
    ende = "B2:B" & ulti - 1
Else
    ende = "B2:B" & ulti
End If
With lstEncomendas
For Each c In Worksheets("Encomendas").Range(ende)
.AddItem c
Next c
End With
TextBox5.Enabled = False
TextBox6.Enabled = False
TextBox12.Enabled = False
TextBox13.Enabled = False
For i = 1 To 13
Me.Controls("lbl" & i).Caption = linha.Cells(1, i + 1).Value
Next i

End Sub

Public Sub clear()
Dim i As Integer
For i = 1 To 13
Me.Controls("TextBox" & i).Text = ""
Next i
lstEncomendas.ListIndex = -1
End Sub

