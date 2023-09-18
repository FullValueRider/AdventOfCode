Attribute VB_Name = "TestIterNum"
Option Explicit
Option Private Module
'@IgnoreModule
'@TestModule
'@Folder("Tests")

#If twinbasic Then
    'Do nothing
#Else
'@TestModule
Option Private Module
'@IgnoreModule
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

Public Sub IterNumTests()
    
#If twinbasic Then
    Debug.Print CurrentProcedureName ; vbTab, vbTab,
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
#End If

    Test01_IsIterNum
    Test02_IsHasNextTrue
    Test03_IsHasPrevFalse
    
    Test04a_MoveNextCountUp
    Test04b_MoveNextCountDown
    Test04c_MoveNextCountDownAfterCountUp
    Test04d_Move56Threetime
    
   ' Test05_Count
    Debug.Print "Testing completed"

End Sub

'@TestMethod("IoN")
Private Sub Test01_IsIterNum()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, True, True)
    Dim myI As IterNum = IterNum(1, 10, 1)
    Dim myResult(0 To 2)  As Boolean

    'Act:
    myResult(0) = VBA.IsObject(myI)
    myResult(1) = "IterNum" = TypeName(myI)
    myResult(2) = "IterNum" = myI.TypeName
    'Assert.Strict:
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
    Dim myI As IterNum = IterNum.Deb(1, 10, 1)
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
    Dim myI As IterNum = IterNum.Deb(1, 10, 1)
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

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedItems As Variant = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    Dim myExpectedKeys As Variant = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    Dim myI As IterNum = IterNum.Deb(1, 10, 1)
    
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

    Loop While myI.MoveNext
   
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
    Dim myExpectedItems As Variant = Array(10, 9, 8, 7, 6, 5, 4, 3, 2, 1)
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    Dim myExpectedKeys As Variant = Array(0, -1, -2, -3, -4, -5, -6, -7, -8, -9)
    Dim myI As IterNum = IterNum.Deb(10, 1, 1)
    
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
    Dim myExpectedItems As Variant = Array(10, 9, 8, 7, 6, 5, 4, 3, 2, 1)
    Dim myExpectedIndexes As Variant = Array(9, 8, 7, 6, 5, 4, 3, 2, 1, 0)
    Dim myExpectedKeys As Variant = Array(9, 8, 7, 6, 5, 4, 3, 2, 1, 0)
    Dim myI As IterNum = IterNum.Deb(1, 10, 1)
    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myREsultKeys As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    
    Do While myI.MoveNext
    Loop
    
    Do
        DoEvents
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
    Loop While myI.MovePrev
   
    myResultItems = myItems.ToArray
    myResultIndexes = myIndexes.ToArray
    myREsultKeys = myKeys.ToArray
    'Assert.Strict:
    
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    AssertStrictSequenceEquals myExpectedIndexes, myResultIndexes, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myREsultKeys, myProcedureName
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
    Dim myExpectedItems As Variant = Array(1, 2, 3, 4, 5, 6, 5, 6, 5, 6)
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4, 5, 4, 5, 4, 5)
    Dim myExpectedKeys As Variant = Array(0, 1, 2, 3, 4, 5, 4, 5, 4, 5)
    Dim myI As IterNum = IterNum.Deb(1, 10, 1)
    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myResultKeys As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    
    With myI
    
    Dim myIndex As Long
    For myIndex = 1 To 4
        
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        .MoveNext
    Next
    
    For myIndex = 1 To 3
        
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        .MoveNext
        
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        myKeys.Add myI.Key
        .MovePrev
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
    Dim myExpectedItems As Variant = Array(1, 2, 3, 4, 5, 1, 2, 3, 4, 5)
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4, 0, 1, 2, 3, 4)
    Dim myI As IterNum = IterNum.Deb(1, 10, 1)
    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    'Act:
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
     
    With myI
    
    Dim myIndex As Long
    For myIndex = 1 To 5
        
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        .MoveNext
    Next
    myI.MoveToStart
    For myIndex = 1 To 5
        
        myItems.Add myI.Item
        myIndexes.Add myI.Index
        .MoveNext
    Next
    End With
 
    myResultItems = myItems.ToArray
    myResultIndexes = myIndexes.ToArray
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
Private Sub Test04f_MoveCountUpFractions()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail
    'Arrange:
    Dim myExpectedItems As Variant = Split("1,1.2,1.4,1.6,1.8", ",")
    Dim myExpectedIndexes As Variant = Array(0, 1, 2, 3, 4)
    Dim myExpectedKeys As Variant = Split("0,0.2,0.4,0.6,0.8", ",")
    Dim myI As IterNum = IterNum.Deb(1.0, 2.0, 0.2)
    
    Dim myResultItems As Variant
    Dim myResultIndexes As Variant
    Dim myResultKeys As Variant
    Dim myItems As Seq = Seq.Deb
    Dim myIndexes As Seq = Seq.Deb
    Dim myKeys As Seq = Seq.Deb
    
    'Act:
    Do
        
        myItems.Add CStr(myI.Item)
        myIndexes.Add myI.Index
        myKeys.Add CStr(myI.Key)
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

' '@TestMethod("IoN")
' Private Sub Test05_Count()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'     'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected As Long = 8
'     Dim myI As IterNum = IterNum.Deb(1.0, 2.4, 0.2)
'     Dim myResult As Long
 
'     'Act:
'     myResult = myI.Count
    
'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub