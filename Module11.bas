Attribute VB_Name = "Module1"
Sub ifTest()
    
    ' = < > <= >= <> and not or
    If Range("a1").Value > 80 Then
        Range("a2").Value = "OK"
    ElseIf Range("a1").Value > 60 Then
        Range("a2").Value = "soso..."
    Else
        Range("a2").Value = "NG!"
    End If
    
    
End Sub
