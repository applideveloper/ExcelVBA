Attribute VB_Name = "Module1"
Sub CellsChange()
    Range("A1", "B3").Value = "hello"
    Range("A5:C7").Value = "hello2"
    Range("4:4").Value = "row 4"
    Range("C:C").Value = "Column C"
    Cells.Clear
End Sub
