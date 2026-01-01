Attribute VB_Name = "extra"
Private Sub txtprocurar_Change()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim filterText As String

    filterText = Me.txtProcurar.Text
    lstClientes.clear
    Set ws = ThisWorkbook.Sheets("Clientes")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For i = 1 To lastRow
        If InStr(1, ws.Cells(i + 1, 2).Value, filterText, vbTextCompare) > 0 Then
            Me.lstClientes.AddItem ws.Cells(i + 1, 2).Value
        End If
    Next i
End Sub
