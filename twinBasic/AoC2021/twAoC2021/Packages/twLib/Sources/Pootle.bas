
Public Sub AssignToArrayParameter()

    Dim myArrayAsParam As Variant = Array(1, 2, 3, 4, 5, 6)
    
    Debug.Print "myArrayAsParam(2) should be 3: ", myArrayAsParam(2)
    Dim myGhost As Ghost = Ghost.Deb(myArrayAsParam)
    
    myGhost.Item(2) = 300
    Debug.Print "myArrayAsParam(2) should be 300:", myArrayAsParam(2)
    Dim mySpectre As Variant = myGhost.Spectre
    Debug.Print "myArrayAsParam(2) should be 300:", myArrayAsParam(2)
    Debug.Print "mySpectre(2) should be 300:", myArrayAsParam(2)
    
    mySpectre(2) = 3000
    
    Debug.Print "myArrayAsParam(2) should be 3000:", myArrayAsParam(2)
    Debug.Print "mySpectre(2) should be 3000:", mySpectre(2)
    
End Sub


Public Sub AssignToArrayParametev2()

    Dim myArrayAsParam(0 To 5) As Long
    myArrayAsParam(0) = 1
    myArrayAsParam(1) = 2
    myArrayAsParam(2) = 3
    myArrayAsParam(3) = 4
    myArrayAsParam(4) = 5
    myArrayAsParam(5) = 6
    
    
   
    Debug.Print "Should be 3: ", myArrayAsParam(2)
    TestAssignment myArrayAsParam
    Debug.Print "Should be 300:", myArrayAsParam(2)
End Sub

Public Sub TestAssignment(ByRef ipArray As Variant)
    Static myLocalArray As Variant
    CopyMemoryToAny myLocalArray, VarPtr(ipArray), 16
    'myLocalArray = GetArrayByRef(ipArray)
    Debug.Print "Localarray(2) should be 3", myLocalArray(2)
    myLocalArray(2) = 300
    Debug.Print "LocalArray(2) should be 300", myLocalArray(2)
    'myLocalArray = 0
End Sub

Sub TestrdMin()

    Dim mySeq As Seq = Seq.Deb.AddItems(5, 4, 6, 9, 2, 10)
    Debug.Print mySeq.ReduceIt(rdMin)
End Sub