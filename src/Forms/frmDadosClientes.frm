VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDadosClientes 
   Caption         =   "Dados Clientes"
   ClientHeight    =   9772.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13080
   OleObjectBlob   =   "frmDadosClientes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDadosClientes"
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

    Set ws = ThisWorkbook.Sheets("Clientes")
    With ws
        Set tbl = .ListObjects(1)
    End With
    Set lookup = tbl.ListColumns(2).DataBodyRange

    linha = Application.WorksheetFunction.Match(lstClientes.Value, lookup, 0)

    If Me.TextBox1.Value = "" Or Me.TextBox2.Value = "" Or Me.TextBox3.Value = "" Or Me.TextBox4.Value = "" Or _
       Me.TextBox5.Value = "" Or Me.TextBox6.Value = "" Or Me.TextBox7.Value = "" Or Me.TextBox8.Value = "" Or _
       Me.TextBox9.Value = "" Or Me.TextBox10.Value = "" Then
        MsgBox "Deve preencher todos os campos obrigatórios.", vbExclamation
        Exit Sub
    End If

    DataValida = VerificarFormatoData(Me.TextBox8.Text)
    If Not DataValida Then
        MsgBox "Por favor, insira uma data no formato dd/mm/aaaa.", vbExclamation
        TextBox8.Text = ""
        Me.TextBox8.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(Me.TextBox10.Text) Or Me.TextBox10.Text < 1 Or Me.TextBox10.Text > 5 Then
        MsgBox "O valor do Feedback deve ser um número entre 1,0 e 5,0.", vbExclamation
        TextBox10.Text = ""
        TextBox10.SetFocus
        Exit Sub
    End If

    For i = 1 To 11
        If i <> 10 Then
            Set cell = tbl.ListColumns(i + 1).DataBodyRange.Cells(linha)
            cell.Value = Me.Controls("TextBox" & i).Text
        Else
            Set cell = tbl.ListColumns(11).DataBodyRange.Cells(linha)
            cell.Value = TextBox10.Text
            cell.Value = CDbl(cell.Value)
        End If
    Next i
    
    MsgBox "Dados alterados com sucesso!", vbInformation
    
    Call clear
    lstClientes.clear

    Dim linha1 As Range
    Set linha1 = Worksheets("Clientes").Rows(1)
    Dim c As Range
    lstClientes.clear
    Dim ulti As Long
    Dim ende As String
    ulti = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    If ws.Cells(ulti, "B").Value = "" Then
        ende = "B2:B" & ulti - 1
    Else
        ende = "B2:B" & ulti
    End If

    With lstClientes
        For Each c In Worksheets("Clientes").Range(ende)
            .AddItem c
        Next c
    End With
    
End Sub

Private Sub cmdVoltar_Click()
Unload Me
frmClientes.Show
End Sub

Private Sub lstClientes_Click()
       Dim tbl As ListObject, i As Integer
        Dim ws As Worksheet
        Dim lookup As Range, linha As Long
        Dim cell As String
     Set ws = ThisWorkbook.Sheets("Clientes")
    With ws
     Set tbl = .ListObjects(1)
    End With
Set lookup = tbl.ListColumns(2).DataBodyRange
linha = Application.WorksheetFunction.Match(lstClientes.Value, lookup, 0)

For i = 1 To 11
Me.Controls("TextBox" & i).Text = Application.WorksheetFunction.Index(tbl.ListColumns(i + 1).DataBodyRange, linha)
Next i

End Sub

Private Sub UserForm_Initialize()
Dim ws As Worksheet, i As Integer, linha As Range
Set ws = ThisWorkbook.Sheets("Clientes")
Set linha = Worksheets("Clientes").Rows(1)
Dim c As Range
lstClientes.clear
Dim ulti As Long
Dim ende As String
ulti = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
If ws.Cells(ulti, "B").Value = "" Then
    ende = "B2:B" & ulti - 1
Else
    ende = "B2:B" & ulti
End If
With lstClientes
For Each c In Worksheets("Clientes").Range(ende)
.AddItem c
Next c
End With
For i = 1 To 11
Me.Controls("lbl" & i).Caption = linha.Cells(1, i + 1).Value
Next i
End Sub

Public Sub clear()
Dim i As Integer
For i = 1 To 11
Me.Controls("TextBox" & i).Text = ""
Next i
lstClientes.ListIndex = -1
End Sub

