



Sub c()
Dim coluna As String
Dim linha As Integer
MsgBox "teste", vbApplicationModal
coluna = InputBox("teste")
linha = 7
Range("I7").Value = Range(coluna & linha).Value
End Sub
