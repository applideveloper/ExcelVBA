Attribute VB_Name = "Module1"
Sub WithTest()
    Range("A1").Value = "hello"
    Range("A1").Font.Bold = True
    Range("A1").Font.Size = 16
    Range("A1").Interior.Color = vbRed
    Cells.Clear
End Sub
Sub WithTest2()
    With Range("A1")
        .Value = "hello"
        With .Font
            .Bold = True
            .Size = 16
        End With
        .Interior.Color = vbRed
    End With
End Sub
