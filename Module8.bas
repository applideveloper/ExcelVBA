Attribute VB_Name = "Module1"
Sub VariableTest()
    Dim x As Integer
    x = 10 + 5
    x = x + 1
    ' + - / *
    ' \ mod ^
    x = 2 ^ 3
    'Range("A1").Value = x
    Debug.Print x
    'イミディエイトウィンドウ
End Sub
