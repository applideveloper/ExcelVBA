Attribute VB_Name = "Module1"
Sub SelectTest()
    Dim signal As String
    signal = Range("a1").Value
    
    Dim result As Range
    Set result = Range("a2")
    
    Select Case signal
    
    Case "red"
        result.Value = "STOP"
    Case "green"
        result.Value = "GO!"
    Case "yellow"
        result.Value = "CAUTION!"
    Case Else
        result.Value = "n.a."
    End Select
End Sub
