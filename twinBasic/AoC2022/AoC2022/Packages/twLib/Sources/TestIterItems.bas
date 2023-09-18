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
    Test02_IsHasNextTrueHasPrevFalse
    Test03_IsHasNextFalseHasPrevTrue
    
    Test04a_MoveNextCountUp
    Test04b_MoveNextCountDown
    Test04c_MoveNextCountDownAfterCountUp
    Test04d_Move56Threetime
    Test04e_MoveCountUpResetAfterFive
    Test04f_MoveCountUpResetAfterFiveUsingSpecificArrayBounds
    Test04g_DictionaryWithStringKeys
    Test04h_CollectionOfStrings
    Test04i_StackOfStrings
    
    Test05_ArrayItemKeyIndex
    Test06_WCollectionItemKeyIndex
    Test07_SeqItemKeyIndex
    Test08_ArrayListItemKeyIndex
    Test09_HkvpItemKeyIndex
    Test10_StackItemKeyIndex
    Test11_QueueItemKeyIndex
    
    Test12_ArrayMutate
    Test13_WCollectionMutate
    Test14_SeqMutate
    Test15_ArrayListMutate
    Test16_HkvpMutate
    Test17_StackMutate
    Test18_QueueMutate
    Debug.Print "Testing completed"

End Sub


'@TestMethod("IterItems")
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

    
    Dim myI As IterItems = IterItems.Deb(Array(10, 20, 30, 40, 50, 60))
    
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


'@TestMethod("IterItems")
Private Sub Test02_IsHasNextTrueHasPrevFalse()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedNext As Boolean = True
    Dim myExpectedPrev As Boolean = False
    Dim myI As IterItems = IterItems.Deb(Array(10, 20, 30, 40, 50))
    Dim myResultNext As Boolean
    Dim myResultPrev As Boolean

    'Act:
    myResultNext = myI.HasNext
    myResultPrev = myI.HasPrev
   
    'Assert.Strict:
    AssertStrictAreEqual myExpectedNext, myResultNext, myProcedureName
    AssertStrictAreEqual myExpectedPrev, myResultPrev, myProcedureName
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test03_IsHasNextFalseHasPrevTrue()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedNext As Boolean = False
    Dim myExpectedPrev As Boolean = True
    Dim myI As IterItems = IterItems.Deb(Array(10, 20, 30, 40, 50)).MoveToEnd
    Dim myResultNext As Boolean
    Dim myResultPrev As Boolean

    'Act:
    myResultNext = myI.HasNext
    myResultPrev = myI.HasPrev
   
    'Assert.Strict:
    AssertStrictAreEqual myExpectedNext, myResultNext, myProcedureName
    AssertStrictAreEqual myExpectedPrev, myResultPrev, myProcedureName
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
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
    
    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myResultKeys As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    Dim myI As IterItems = IterItems.Deb(Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100))
    Do
        DoEvents
        myItems.Add myI(0)
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

'@TestMethod("IterItems")
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
    Dim myI As IterItems = IterItems.Deb(Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100))
    myI.MoveToEnd
    Do
        DoEvents
        myItems.Add myI(0)
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

'@TestMethod("IterItems")
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
    Dim myI As IterItems = IterItems.Deb(Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100))
    
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
        myItems.Add myI(0)
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

'@TestMethod("IterItems")
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
    
    
    Dim myI As IterItems = IterItems.Deb(Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100))
    Dim myIndex As Long
    For myIndex = 1 To 4
        
        myItems.Add myI(0)
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        myI.MoveNext
    Next
    
    For myIndex = 1 To 3
       
        myItems.Add myI(0)
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        myI.MoveNext

        myItems.Add myI(0)
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

'@TestMethod("IterItems")
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
    Dim myI As IterItems = IterItems.Deb(Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100))
    
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
        
        myItems.Add myI(0)
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        .MoveNext
    Next
    myI.MoveToStart
    For myIndex = 1 To 5
        
        myItems.Add myI(0)
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

'@TestMethod("IterItems")
Private Sub Test04f_MoveCountUpResetAfterFiveUsingSpecificArrayBounds()

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
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4, 0, 1, 2, 3, 4)
    Dim myExpectedKeys As Variant = Array(-4, -3, -2, -1, 0, -4, -3, -2, -1, 0)
    
    Dim myArray As Variant = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100)
    ReDim Preserve myArray(-4 To 5)
    Dim myI As IterItems = IterItems.Deb(myArray)
    
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
        
        myItems.Add myI(0)
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        .MoveNext
    Next
    myI.MoveToStart
    For myIndex = 1 To 5
        
        myItems.Add myI(0)
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

'@TestMethod("IterItems")
Private Sub Test04g_DictionaryWithStringKeys()

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
    
    
    Dim myI As IterItems = IterItems.Deb(myH)
    Do
        myItems.Add myI(0)
        myIndexes.Add myI.Index(0)
        myKeys.Add myI.Key(0)
        
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

'@TestMethod("IterItems")
Private Sub Test04h_CollectionOfStrings()

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
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4, 5, 6)
    Dim myExpectedKeys As Variant = Array(1, 2, 3, 4, 5, 6, 7)
    
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
    
    
    Dim myI As IterItems = IterItems.Deb(myC)
    Do
        myItems.Add myI(0)
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

'@TestMethod("IterItems")
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
    Dim myExpectedItems As Variant = Split("Hello World Its A Nice Day Today", " ")
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4, 5, 6)
    Dim myExpectedKeys As Variant = Array(0, 1, 2, 3, 4, 5, 6)
    
    Dim myS As Stack = Stack.Deb
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
    
    
    Dim myI As IterItems = IterItems.Deb(myS)
    Do
        myItems.Add myI(0)
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

'@TestMethod("IterItems")
Private Sub Test05_ArrayItemKeyIndex()
    
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
    myExpected = Array(9.0, 9, 4, 12.0, 12, 7, 6.0, 6, 1)

    Dim myArray(5 To 15) As Double
    myArray(5) = 5.0
    myArray(6) = 6.0
    myArray(7) = 7.0
    myArray(8) = 8.0
    myArray(9) = 9.0
    myArray(10) = 10.0
    myArray(11) = 11.0
    myArray(12) = 12.0
    myArray(13) = 13.0
    myArray(14) = 14.0
    myArray(15) = 15.0
    
    Dim myI As IterItems = IterItems.Deb(myArray)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test06_WCollectionItemKeyIndex()
    
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
    myExpected = Array(9.0, 5, 4, 12.0, 8, 7, 6.0, 2, 1)

    Dim mywColl As wCollection = wCollection.Deb
    With mywColl
        .Add 5.0
        .Add 6.0
        .Add 7.0
        .Add 8.0
        .Add 9.0
        .Add 10.0
        .Add 11.0
        .Add 12.0
        .Add 13.0
        .Add 14.0
        .Add 15.0
     
     End With
    
    Dim myI As IterItems = IterItems.Deb(mywColl)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test07_SeqItemKeyIndex()
    
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
    myExpected = Array(9.0, 5, 4, 12.0, 8, 7, 6.0, 2, 1)

    Dim mySeq As Seq = Seq.Deb
    With mySeq
        .Add 5.0
        .Add 6.0
        .Add 7.0
        .Add 8.0
        .Add 9.0
        .Add 10.0
        .Add 11.0
        .Add 12.0
        .Add 13.0
        .Add 14.0
        .Add 15.0
     
     End With
    
    Dim myI As IterItems = IterItems.Deb(mySeq)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test08_ArrayListItemKeyIndex()
    
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
    myExpected = Array(9.0, 4, 4, 12.0, 7, 7, 6.0, 1, 1)

    Dim myAL As ArrayList = New ArrayList
    With myAL
        .Add 5.0
        .Add 6.0
        .Add 7.0
        .Add 8.0
        .Add 9.0
        .Add 10.0
        .Add 11.0
        .Add 12.0
        .Add 13.0
        .Add 14.0
        .Add 15.0
     
     End With
    
    Dim myI As IterItems = IterItems.Deb(myAL)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test09_HkvpItemKeyIndex()
    
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
    myExpected = Array(9.0, "Nine", 4, 12.0, "Twelve", 7, 6.0, "Six", 1)

    Dim myH As Hkvp = Hkvp.Deb
    With myH
        .Add "Five", 5.0
        .Add "Six", 6.0
        .Add "Seven", 7.0
        .Add "Eight", 8.0
        .Add "Nine", 9.0
        .Add "Ten", 10.0
        .Add "Eleven", 11.0
        .Add "Twelve", 12.0
        .Add "Thirteen", 13.0
        .Add "Fourteen", 14.0
        .Add "Fifteen", 15.0
     
     End With
    
    Dim myI As IterItems = IterItems.Deb(myH)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test10_StackItemKeyIndex()
    
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
    myExpected = Array(9.0, 4, 4, 12.0, 7, 7, 6.0, 1, 1)

    Dim myStack As Stack = Stack.Deb
    With myStack
        .Push 5.0
        .Push 6.0
        .Push 7.0
        .Push 8.0
        .Push 9.0
        .Push 10.0
        .Push 11.0
        .Push 12.0
        .Push 13.0
        .Push 14.0
        .Push 15.0
     
     End With
    
    Dim myI As IterItems = IterItems.Deb(myStack)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test11_QueueItemKeyIndex()
    
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
    myExpected = Array(9.0, 4, 4, 12.0, 7, 7, 6.0, 1, 1)

    Dim myQ As Queue = Queue.Deb
    With myQ
        .Enqueue 5.0
        .Enqueue 6.0
        .Enqueue 7.0
        .Enqueue 8.0
        .Enqueue 9.0
        .Enqueue 10.0
        .Enqueue 11.0
        .Enqueue 12.0
        .Enqueue 13.0
        .Enqueue 14.0
        .Enqueue 15.0
     
     End With
    
    Dim myI As IterItems = IterItems.Deb(myQ)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test12_ArrayMutate()
    
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
    myExpected = Array(42.0, 9, 4, 43.0, 12, 7, 41.0, 6, 1)

    Dim myArray(5 To 15) As Double
    myArray(5) = 5.0
    myArray(6) = 6.0
    myArray(7) = 7.0
    myArray(8) = 8.0
    myArray(9) = 9.0
    myArray(10) = 10.0
    myArray(11) = 11.0
    myArray(12) = 12.0
    myArray(13) = 13.0
    myArray(14) = 14.0
    myArray(15) = 15.0
    
    Dim myI As IterItems = IterItems.Deb(myArray)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    

    myI.SetItem 42.0
    myI.SetItem 43.0, 3
    myI.SetItem 41.0, -3
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test13_WCollectionMutate()
    
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
    myExpected = Array(42.0, 5, 4, 43.0, 8, 7, 41.0, 2, 1)

    Dim mywColl As wCollection = wCollection.Deb
    With mywColl
        .Add 5.0
        .Add 6.0
        .Add 7.0
        .Add 8.0
        .Add 9.0
        .Add 10.0
        .Add 11.0
        .Add 12.0
        .Add 13.0
        .Add 14.0
        .Add 15.0
     
     End With
    
    Dim myI As IterItems = IterItems.Deb(mywColl)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    myI.SetItem 42.0
    myI.SetItem 43.0, 3
    myI.SetItem 41.0, -3
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test14_SeqMutate()
    
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
    myExpected = Array(42.0, 5, 4, 43.0, 8, 7, 41.0, 2, 1)

    Dim mySeq As Seq = Seq.Deb
    With mySeq
        .Add 5.0
        .Add 6.0
        .Add 7.0
        .Add 8.0
        .Add 9.0
        .Add 10.0
        .Add 11.0
        .Add 12.0
        .Add 13.0
        .Add 14.0
        .Add 15.0
     
     End With
    
    Dim myI As IterItems = IterItems.Deb(mySeq)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    myI.SetItem 42.0
    myI.SetItem 43.0, 3
    myI.SetItem 41.0, -3
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    AssertStrictAreEqual 42.0, mySeq.Item(5), myProcedureName
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test15_ArrayListMutate()
    
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
    myExpected = Array(42.0, 4, 4, 43.0, 7, 7, 41.0, 1, 1)

    Dim myAL As ArrayList = New ArrayList
    With myAL
        .Add 5.0
        .Add 6.0
        .Add 7.0
        .Add 8.0
        .Add 9.0
        .Add 10.0
        .Add 11.0
        .Add 12.0
        .Add 13.0
        .Add 14.0
        .Add 15.0
     
     End With
    
    Dim myI As IterItems = IterItems.Deb(myAL)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    myI.SetItem 42.0
    myI.SetItem 43.0, 3
    myI.SetItem 41.0, -3
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test16_HkvpMutate()
    
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
    myExpected = Array(42.0, "Nine", 4, 43.0, "Twelve", 7, 41.0, "Six", 1)

    Dim myH As Hkvp = Hkvp.Deb
    With myH
        .Add "Five", 5.0
        .Add "Six", 6.0
        .Add "Seven", 7.0
        .Add "Eight", 8.0
        .Add "Nine", 9.0
        .Add "Ten", 10.0
        .Add "Eleven", 11.0
        .Add "Twelve", 12.0
        .Add "Thirteen", 13.0
        .Add "Fourteen", 14.0
        .Add "Fifteen", 15.0
     
     End With
    
    Dim myI As IterItems = IterItems.Deb(myH)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    myI.Item(0) = 42.0
    myI.Item(3) = 43.0
    myI.Item(-3) = 41.0
    
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key(0)     ' Key value (native index)
    myResult(2) = myI.Index(0)     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test17_StackMutate()
    
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
    myExpected = Array(42.0, 4, 4, 43.0, 7, 7, 41.0, 1, 1)

    Dim myStack As Stack = Stack.Deb
    With myStack
        .Push 5.0
        .Push 6.0
        .Push 7.0
        .Push 8.0
        .Push 9.0
        .Push 10.0
        .Push 11.0
        .Push 12.0
        .Push 13.0
        .Push 14.0
        .Push 15.0
     
     End With
    
    Dim myI As IterItems = IterItems.Deb(myStack)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    myI.SetItem 42.0
    myI.SetItem 43.0, 3
    myI.SetItem 41.0, -3
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



'@TestMethod("IterItems")
Private Sub Test18_QueueMutate()
    
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
    myExpected = Array(42.0, 4, 4, 43.0, 7, 7, 41.0, 1, 1)

    Dim myQ As Queue = Queue.Deb
    With myQ
        .Enqueue 5.0
        .Enqueue 6.0
        .Enqueue 7.0
        .Enqueue 8.0
        .Enqueue 9.0
        .Enqueue 10.0
        .Enqueue 11.0
        .Enqueue 12.0
        .Enqueue 13.0
        .Enqueue 14.0
        .Enqueue 15.0
     
     End With
    
    Dim myI As IterItems = IterItems.Deb(myQ)
    
    With myI
        
        .MoveNext
        .MoveNext
        .MoveNext
        .MoveNext ' should be at item 9.0
        
    End With
    
    myI.SetItem 42.0
    myI.SetItem 43.0, 3
    myI.SetItem 41.0, -3
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    myResult(0) = myI(0)   ' Ittem value
    myResult(1) = myI.Key     ' Key value (native index)
    myResult(2) = myI.Index     ' Index (offset from firstindex)
    myResult(3) = myI(3)
    myResult(4) = myI.Key(3)
    myResult(5) = myI.Index(3)
    myResult(6) = myI(-3)
    myResult(7) = myI.Key(-3)
    myResult(8) = myI.Index(-3)
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub