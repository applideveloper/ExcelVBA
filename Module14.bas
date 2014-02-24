Attribute VB_Name = "Module1"
Sub CallTest()

    Dim names As Variant
    names = Array("taguchi", "fkoji", "dotinstall")
    
    For Each name In names
        Debug.Print SayHi(name)
    Next name

End Sub

Function SayHi(ByVal name As String)
    SayHi = "hi!, " & name
End Function

