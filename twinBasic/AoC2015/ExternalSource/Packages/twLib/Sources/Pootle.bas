Sub TestFunctionOnLoopStatement()
    Dim myCounter As Long = 0
    Do
    	myCounter += 1
        Debug.Print myCounter,
        If myCounter < 10 Then
            Continue Do
        End If
    Loop While LoopEnd
End Sub

Public Function LoopEnd() As Boolean
Debug.Print
    Debug.Print "I get printed once, not 10 times."
    Return False
End Function

Sub TestTreap()

    Dim myT As Treap = Treap.Deb
    With myT
        .Add 10, 100
        .Add 20, 200
        .Add 30, 300
    End With
        ' Dim myItems As Variant = myT.Items
        ' Dim myKeys As Variant = myT.Keys

End Sub

Public Sub TestGetRndLong()

    Dim i As Long
    For i = 1 To 100
        Debug.Print Maths.GetRndLong
    Next
End Sub

Public Sub TestMidStatement()
    Dim myString As String = "HellowWorld"
    VBA.Mid(myString, 6, 1) = " "
    Debug.Print myString
End Sub

Public Sub TestMidToStringInArray()
	
    Dim myArray() As String
    ReDim myArray(1 To 4)
    myArray(1) = VBA.Space$(10)
    myArray(2) = "__________"
    myArray(3) = VBA.Space$(10)
    myArray(4) = VBA.Space$(10)
    Mid(myArray(2), 3, 1) = "x"
    Dim myString As String = "HellowWorld"
    Mid(myString, 6, 1) = " "
    Debug.Print "test output"
    Debug.Print myArray(2) ; "Hellow world"
    Debug.Print myString
End Sub