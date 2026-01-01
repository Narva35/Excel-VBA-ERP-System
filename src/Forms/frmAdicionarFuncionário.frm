VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdicionarFuncionário 
   Caption         =   "Adicionar Funcionário"
   ClientHeight    =   8771.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10296
   OleObjectBlob   =   "frmAdicionarFuncionário.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdicionarFuncionário"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAdicionar_Click()
Dim novalinha As Integer
Dim ws As Worksheet

Set ws = ThisWorkbook.Sheets("Funcionários")

If TextBox1.Value = "" Or ComboBox2.Text = "" Or TextBox3.Value = "" Or TextBox4.Value = "" Or TextBox5.Value = "" Or TextBox6.Value = "" Or TextBox7.Value = "" Or TextBox8.Value = "" Or TextBox9.Value = "" Then
    MsgBox ("Deve preencher todos os campos.")
Else
    novalinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(novalinha, 2) = TextBox1.Value
    ws.Cells(novalinha, 3) = ComboBox2.Text
    ws.Cells(novalinha, 4) = TextBox3.Value
    ws.Cells(novalinha, 5) = TextBox4.Value
    ws.Cells(novalinha, 6) = TextBox5.Value
    ws.Cells(novalinha, 7) = TextBox6.Value
    ws.Cells(novalinha, 8) = TextBox7.Value
    ws.Cells(novalinha, 9) = TextBox8.Value
    ws.Cells(novalinha, 11) = ComboBox1.Text
    ws.Cells(novalinha, 12) = TextBox9.Value
    ws.Cells(novalinha, 10).Formula = "=DATADIF(RC[-1],HOJE(),""Y"")"
    
    MsgBox ("Funcionário adicionado com sucesso!")
    MsgBox ("Por favor atualize a informação nas fábricas, +1 funcionário na fábrica referida.")
End If
    
Dim i As Integer
For i = 1 To 9
If i <> 2 Then
Me.Controls("Textbox" & i).Text = ""
End If
Next i
Dim id As String, linha As Integer, tbl As ListObject, coluna As Range
ComboBox2.Text = id
Set ws = ThisWorkbook.Sheets("Fábricas")
Set tbl = ws.ListObjects(1)
Set coluna = tbl.ListColumns(3).DataBodyRange
linha = Application.WorksheetFunction.Match(id, coluna, 0)


ComboBox1.Text = ""
ComboBox2.Text = ""

End Sub

Private Sub cmdVoltar_Click()
    Unload Me
    frmFuncionários.Show
End Sub

Public Sub ComboBox2_Change()
TextBox3.Text = Left(ComboBox2.Value, Len(ComboBox2.Value) - 2)
End Sub

Private Sub UserForm_Initialize()
Dim vec(4) As String
vec(0) = "Diretor"
vec(1) = "Gestor"
vec(2) = "Engenheiro"
vec(4) = "Operador de Máquina"
vec(3) = "Supervisor"
ComboBox1.List = vec
Dim ws As Worksheet
Dim colFabs As New Collection
Dim i As Long
Dim fab As String

Set ws = ThisWorkbook.Sheets("Fábricas")

' Inicializar a variável de fábrica com o valor da célula C2
fab = ws.Cells(2, 3).Value

' Loop para percorrer as linhas preenchidas na coluna A
For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If Not IsInCollection(colFabs, fab) Then
        colFabs.Add fab, fab
    End If
    ' Atualizar o valor da variável fab para a próxima linha
    fab = ws.Cells(i + 1, 3).Value
Next i

' Converter a coleção em um array de strings
Dim veca() As String
ReDim veca(1 To colFabs.Count)
For i = 1 To colFabs.Count
    veca(i) = colFabs(i)
Next i

' Preencher o ComboBox1 com o array de fábricas únicas
ComboBox2.List = veca

End Sub

Function IsInCollection(col As Collection, key As String) As Boolean
    On Error Resume Next
    IsInCollection = Not col(key) Is Nothing
    On Error GoTo 0
End Function
