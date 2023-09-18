Attribute VB_Name = "TestIterItems"
Option Explicit

#If twinbasic Then
    'Do nothing
#Else
'@IgnoreModule
'@TestModule
Option Private Module

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
#End If


Public Sub IterItemsTest()

#If twinbasic Then
    Debug.Print CurrentProcedureName ; vbTab, vbTab,
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
#End If

    Test01_NewIterItems
    Test02_IsHasNextTrue
    Test03_IsHasPrevFalse
    
    Test04a_MoveNextCountUp
    Test04b_MoveNextCountDown
    Test04c_MoveNextCountDownAfterCountUp
    Test04d_Move56Threetime
    Test04e_MoveCountUpResetAfterFive
    Test04g_MoveCountUpResetAfterFiveUsingSpecificArrayBounds
    Test04h_DictionaryWithStringKeys
    Test04i_CollectionOfStrings
    
    Debug.Print "Testing completed"

End Sub

'@TestMethod("Seq")
Private Sub Test01_NewIterItems()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(True, True, True)

    Dim myI As IterItems = IterItems(Array(10, 20, 30, 40, 50, 60))
    
    Dim myResult As Variant
    ReDim myResult(0 To 2)
    
    'Act:
    myResult(0) = VBA.IsObject(myI)
    myResult(1) = "IterItems" = VBA.TypeName(myI)
    myResult(2) = "IterItems" = myI.TypeName
    
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IoN")
Private Sub Test02_IsHasNextTrue()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Boolean = True
    Dim myI As IterItems = IterItems(Array(10, 20, 30, 40, 50))
    Dim myResult  As Boolean

    'Act:
    myResult = myI.HasNext
   
    'Assert.Strict:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IoN")
Private Sub Test03_IsHasPrevFalse()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Boolean = False
    Dim myI As IterItems = IterItems(Array(10, 20, 30, 40, 50, 60))
    Dim myResult  As Boolean

    'Act:
    myResult = myI.HasPrev
   
    'Assert.Strict:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IoN")
Private Sub Test04a_MoveNextCountUp()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    ''On Error GoTo TestFail

    'Arrange:
    Dim myExpectedItems As Variant = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100)
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    Dim myExpectedKeys As Variant = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    Dim myI As IterItems = IterItems(Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100))
    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myResultKeys As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    Do
        DoEvents
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
    Loop While myI.MoveNext
   
    myResultItems = myItems.ToArray
    myResultIndexes = myIndexes.ToArray
    myResultKeys = myKeys.ToArray
    'Assert.Strict:
    
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    AssertStrictSequenceEquals myExpectedIndexes, myResultIndexes, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IoN")
Private Sub Test04b_MoveNextCountDown()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedItems As Variant = Array(100, 90, 80, 70, 60, 50, 40, 30, 20, 10)
    Dim myExpectedIndexes As Variant = Array(9, 8, 7, 6, 5, 4, 3, 2, 1, 0)
    Dim myExpectedKeys As Variant = Array(9, 8, 7, 6, 5, 4, 3, 2, 1, 0)
    
    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myResultKeys As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    Dim myI As IterItems = IterItems(Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100))
    myI.MoveToEnd
    Do
        DoEvents
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
    Loop While myI.MovePrev
   
    myResultItems = myItems.ToArray
    myResultIndexes = myIndexes.ToArray
    myResultKeys = myKeys.ToArray
    'Assert.Strict:
    
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    AssertStrictSequenceEquals myExpectedIndexes, myResultIndexes, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IoN")
Private Sub Test04c_MoveNextCountDownAfterCountUp()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedItems As Variant = Array(100, 90, 80, 70, 60, 50, 40, 30, 20, 10)
    Dim myExpectedIndexes As Variant = Array(9, 8, 7, 6, 5, 4, 3, 2, 1, 0)
    Dim myExpectedKeys As Variant = Array(9, 8, 7, 6, 5, 4, 3, 2, 1, 0)
    Dim myI As IterItems = IterItems(Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100))
    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myResultKeys As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    Do
    Loop While myI.MoveNext
    
    Do
        DoEvents
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
    Loop While myI.MovePrev
   
    myResultItems = myItems.ToArray
    myResultIndexes = myIndexes.ToArray
    myResultKeys = myKeys.ToArray
    'Assert.Strict:
    
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    AssertStrictSequenceEquals myExpectedIndexes, myResultIndexes, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IoN")
Private Sub Test04d_Move56Threetime()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedItems As Variant = Array(10, 20, 30, 40, 50, 60, 50, 60, 50, 60)
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4, 5, 4, 5, 4, 5)
    Dim myExpectedKeys As Variant = Array(0, 1, 2, 3, 4, 5, 4, 5, 4, 5)
    
    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myResultKeys As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    
    
    Dim myI As IterItems = IterItems(Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100))
    Dim myIndex As Long
    For myIndex = 1 To 4
        
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        myI.MoveNext
    Next
    
    For myIndex = 1 To 3
       
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        myI.MoveNext

        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        myI.MovePrev
    Next
    
 
    myResultItems = myItems.ToArray
    myResultIndexes = myIndexes.ToArray
    myResultKeys = myKeys.ToArray
    'Assert.Strict:
    
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    AssertStrictSequenceEquals myExpectedIndexes, myResultIndexes, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IoN")
Private Sub Test04e_MoveCountUpResetAfterFive()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedItems As Variant = Array(10, 20, 30, 40, 50, 10, 20, 30, 40, 50)
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4, 0, 1, 2, 3, 4)
    Dim myExpectedKeys As Variant = Array(0, 1, 2, 3, 4, 0, 1, 2, 3, 4)
    Dim myI As IterItems = IterItems(Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100))
    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myResultKeys As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    With myI
    
    Dim myIndex As Long
    For myIndex = 1 To 5
        
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        .MoveNext
    Next
    myI.MoveToStart
    For myIndex = 1 To 5
        
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        .MoveNext
    Next
    End With
 
    myResultItems = myItems.ToArray
    myResultIndexes = myIndexes.ToArray
    myResultKeys = myKeys.ToArray
    'Assert.Strict:
    
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    AssertStrictSequenceEquals myExpectedIndexes, myResultIndexes, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IoN")
Private Sub Test04g_MoveCountUpResetAfterFiveUsingSpecificArrayBounds()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    ''On Error GoTo TestFail

    'Arrange:
    Dim myExpectedItems As Variant = Array(10, 20, 30, 40, 50, 10, 20, 30, 40, 50)
    Dim myExpectedIndexes As Variant = Array(-4, -3, -2, -1, 0, -4, -3, -2, -1, 0)
    Dim myExpectedKeys As Variant = Array(0, 1, 2, 3, 4, 0, 1, 2, 3, 4)
    
    Dim myArray As Variant = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100)
    ReDim Preserve myArray(-4 To 5)
    Dim myI As IterItems = IterItems(myArray)
    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myResultKeys As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    With myI
    
    Dim myIndex As Long
    For myIndex = 1 To 5
        
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        .MoveNext
    Next
    myI.MoveToStart
    For myIndex = 1 To 5
        
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        .MoveNext
    Next
    End With
 
    myResultItems = myItems.ToArray
    myResultIndexes = myIndexes.ToArray
    myResultKeys = myKeys.ToArray
    'Assert.Strict:
    
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    AssertStrictSequenceEquals myExpectedIndexes, myResultIndexes, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IoN")
Private Sub Test04h_DictionaryWithStringKeys()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    ''On Error GoTo TestFail

    'Arrange:
    Dim myExpectedItems As Variant = Array(10, 20, 30, 40, 50, 60, 70)
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4, 5, 6)
    Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day Today", " ")
    
    Dim myH As Hkvp = Hkvp.Deb.AddPairs(Split("Hello World Its A Nice Day Today", " "), Array(10, 20, 30, 40, 50, 60, 70))

    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myResultKeys As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    
    
    Dim myI As IterItems = IterItems(myH)
    Do
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        
    Loop While myI.MoveNext
    
 
    myResultItems = myItems.ToArray
    myResultIndexes = myIndexes.ToArray
    myResultKeys = myKeys.ToArray
    'Assert.Strict:
    
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    AssertStrictSequenceEquals myExpectedIndexes, myResultIndexes, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IoN")
Private Sub Test04i_CollectionOfStrings()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    ''On Error GoTo TestFail

    'Arrange:
    Dim myExpectedItems As Variant = Split("Hello World Its A Nice Day Today", " ")
    Dim myExpectedIndexes As Variant = Array(1, 2, 3, 4, 5, 6, 7)
    Dim myExpectedKeys As Variant = Array(0, 1, 2, 3, 4, 5, 6)
    
    Dim myC As Collection = New Collection
    With myC
     
        .Add "Hello"
        .Add "World"
        .Add "Its"
        .Add "A"
        .Add "Nice"
        .Add "Day"
        .Add "Today"
    End With

    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myResultKeys As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    
    
    Dim myI As IterItems = IterItems(myC)
    Do
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        
    Loop While myI.MoveNext
    
 
    myResultItems = myItems.ToArray
    myResultIndexes = myIndexes.ToArray
    myResultKeys = myKeys.ToArray
    'Assert.Strict:
    
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    AssertStrictSequenceEquals myExpectedIndexes, myResultIndexes, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    
    TestExit:
    Exit Sub
    
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("IoN")
Private Sub Test04i_StackOfStrings()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    ''On Error GoTo TestFail

    'Arrange:
    Dim myExpectedItems As Variant = Split("Today Day Nice A Its World Hello", " ")
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4, 5, 6)
    Dim myExpectedKeys As Variant = Array(0, 1, 2, 3, 4, 5, 6)
    
    Dim myS As Stack = New Stack
    With myS
     
        .Push "Hello"
        .Push "World"
        .Push "Its"
        .Push "A"
        .Push "Nice"
        .Push "Day"
        .Push "Today"
        
    End With

    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myResultKeys As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    
    
    Dim myI As IterItems = IterItems(myS)
    Do
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        
    Loop While myI.MoveNext
    
 
    myResultItems = myItems.ToArray
    myResultIndexes = myIndexes.ToArray
    myResultKeys = myKeys.ToArray
    'Assert.Strict:
    
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    AssertStrictSequenceEquals myExpectedIndexes, myResultIndexes, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    
    TestExit:
    Exit Sub
    
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub