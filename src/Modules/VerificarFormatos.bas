Attribute VB_Name = "VerificarFormatos"
Public Function VerificarFormatoData(DataTexto As String) As Boolean

    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    
    With RegEx
        .Pattern = "^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[012])/(19|20)\d\d$"
        .IgnoreCase = True
        .Global = False
    End With

    If RegEx.Test(DataTexto) Then
        VerificarFormatoData = True
    Else
        VerificarFormatoData = False
    End If
    
End Function

