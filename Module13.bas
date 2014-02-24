Attribute VB_Name = "Module1"
Sub WhileTest()
    Dim i As Integer
    i = 1
    Do While i < 10
        Cells(i, 1).Value = i
        i = i + 1
    Loop
End Sub
Sub ForTest()
    Dim i As Integer
    
    For i = 1 To 9 Step 2
        Cells(i, 1).Value = i
    Next i
End Sub


