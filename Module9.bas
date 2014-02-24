Attribute VB_Name = "Module1"
Sub VariableTest()
    Dim y As Double
    Dim s As String
    Dim d As Date
    Dim z As Variant
    Dim f As Boolean
    Dim r As Range
    
    y = 10.5
    s = "hello"
    d = "2012/04/23"
    f = True
    Set r = Range("A1")
    
    Debug.Print y / 3
    Debug.Print s & "world"
    r.Value = d + 7
    Debug.Print r.Value
End Sub
