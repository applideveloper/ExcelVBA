Attribute VB_Name = "Module1"
Sub VariableTest()
    'sales_0 = 200
    'sales_1 = 150
    'sales_2 = 300
    Dim sales(2) As Integer
    sales(0) = 200
    sales(1) = 150
    sales(2) = 300
    
    Debug.Print sales(1)
End Sub
Sub VariableTest2()
    Dim sales As Variant
    sales = Array(200, 150, 300)
    Debug.Print sales(2)
End Sub
