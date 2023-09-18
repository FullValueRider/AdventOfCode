Attribute VB_Name = "TestSeq"
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


Public Sub SeqTests()

#If twinbasic Then
    Debug.Print CurrentProcedureName ; vbTab, vbTab, vbTab,
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
#End If

    Test01_NewSeqIsObject
    
    Test02a_AddElementsByAdd
    Test02b_AddElementsByAddItems
    
    Test03a_AddElementsByDeb
    Test03b_AddStringByDeb
    Test03c_AddArrayByDeb
    Test03d_AddCollectionByDeb
    Test03e_AddDictionaryByDeb

    Test04a_AddStringByAddRange
    Test04b_AddArrayByAddRange
    Test04c_AddCollectionByAddRange
    Test04d_AddDictionaryByAddRange
    
    Test05_Clear
    
    Test06_Clone
    
    Test07a_ContainsTrue
    Test07b_ContainsFalse
    Test07c_HoldsItemTrue
    Test07d_HoldsValueFalse
    Test07e_LacksValueTrue
    Test07f_LacksValueFalse
    
    Test08a_CopyToAllArray
    Test08b_CopyToArrayStart
    Test08c_CopyToSeqStartArrayRun
    
    Test09a_GetRangeViaSlice
    Test09b_GetRangeViaGetRange
    Test09c_GetRangeViaSliceWithEndIndex
    Test09d_GetRangeViaGetRangeWithEndIndex
    
    Test10a_FirstIndexOfInWholeSeq
    Test10b_FirstIndexOfInWithStart
    Test10c_FirstIndexOfInWithStartRun
    Test10d_FirstIndexOfInWithStartEnd
 
    Test11A_InsertSingleItem
    Test11b_InsertMultipleItems
    Test11c_InsertRangeString
    Test11D_InsertRangeArray
    Test11e_InsertRangeCollection
    Test11f_InsertRangeStack
    Test11g_InsertRangeDictionary
  
    Test12a_LastIndexOf
    
    Test13a_AssignItemTwoPrimitive
    Test13b_AssignAndGetItemTwoArray
    
    Test14a_RemoveAll
    Test14b_RemoveItemsIndividually
    Test14c_RemoveBlockWithinLastIndex
    Test14d_RemoveBlockPastEndOfItems
    
    Test15a_SliceItem1ToItem3
    Test15b_SliceItem3Run4ToItem3
    Test15c_SliceItem3RunPastEnd
    
    Test16a_InsertAtItem1
    Test16b_InsertRangeFivetemsFromItem1
    
    Test17a_RemoveAtItem4
    
    Debug.Print "Testing completed"

End Sub

'@TestMethod("Seq")
Private Sub Test01_NewSeqIsObject()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(True, "Seq", "Seq")

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb
    
    Dim myResult As Variant
    ReDim myResult(0 To 2)
    
    'Act:
    myResult(0) = VBA.IsObject(mySeq)
    myResult(1) = VBA.TypeName(mySeq)
    myResult(2) = mySeq.TypeName
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test02a_AddElementsByAdd()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedCount  As Variant
    myExpectedCount = 5

    Dim myExpectedSeq As Variant
    myExpectedSeq = Array(10, 20, 30, 40, 50)
    Dim myExpectedIndexes As Variant = Array(1, 2, 3, 4, 5)
    
    Dim myResultIndexes As Variant
    ReDim myResultIndexes(0 To 4)
    Dim myResultCount As Long
    Dim myResultSeq As Variant
    
    Dim mySeq As Seq
    Set mySeq = Seq.Deb
    
    'Act:
    ' Add. as per cHashD, returns the index at which the item was added
    myResultIndexes(0) = mySeq.Add(10)
    myResultIndexes(1) = mySeq.Add(20)
    myResultIndexes(2) = mySeq.Add(30)
    myResultIndexes(3) = mySeq.Add(40)
    myResultIndexes(4) = mySeq.Add(50)
    
    myResultCount = mySeq.Count
    myResultSeq = mySeq.ToArray
    'Assert:
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedSeq, myResultSeq, myProcedureName
    AssertStrictSequenceEquals myExpectedIndexes, myResultIndexes, myProcedureName
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test02b_AddElementsByAddItems()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedCount  As Variant
    myExpectedCount = 5

    Dim myExpectedSeq As Variant
    myExpectedSeq = Array(10, 20, 30, 40, 50)
    '@Ignore IntegerDataType
    Dim mySeq As Seq = Seq.Deb
    
    mySeq.AddItems 10
    mySeq.AddItems 20
    mySeq.AddItems 30
    mySeq.AddItems 40
    mySeq.AddItems 50
    
    Dim myResultCOunt As Long
    Dim myResultSeq As Variant
    'Act:
    myResultCOunt = mySeq.Count
    myResultSeq = mySeq.ToArray
    'Assert:
    AssertStrictAreEqual myExpectedCount, myResultCOunt, myProcedureName
    AssertStrictSequenceEquals myExpectedSeq, myResultSeq, myProcedureName
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test03a_AddElementsByDeb()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 5
    Dim myExpectedItems As Variant = Array(10, 20, 30, 40, 50)

    '@Ignore IntegerDataType
    
    Dim mySeq As Seq = Seq.Deb(10, 20, 30, 40, 50)
    Dim myResultCount As Long
    Dim myResultItems As Variant
    'Act:
    myResultCount = mySeq.Count
    myResultItems = mySeq.ToArray
    
    'Assert:
    AssertStrictAreEqual myExpected, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test03b_AddStringByDeb()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array("H", "e", "l", "l", "o")

    Dim mySeq As Seq = Seq.Deb("Hello")
    Dim myResult As Variant
    
    'Act:
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test03c_AddArrayByDeb()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)

    '@Ignore IntegerDataType
    
    Dim mySeq As Seq = Seq.Deb(Array(10, 20, 30, 40, 50, 60, 70, 80, 90))
    Dim myResult As Variant
    
    'Act:
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test03d_AddCollectionByDeb()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)

    '@Ignore IntegerDataType
    Dim myC As Collection = New Collection
    With myC
    
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    
    End With
    Dim mySeq As Seq = Seq.Deb(myC)
    Dim myResult As Variant
    
    'Act:
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test03e_AddDictionaryByDeb()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    
    Dim myHkvp As Hkvp = Hkvp.Deb.AddKnownArrayPairs(Array("a", "b", "c", "d", "e", "f", "g", "h ", "i"), Array(10, 20, 30, 40, 50, 60, 70, 80, 90))
    Dim mySeq As Seq = Seq.Deb(myHkvp)
    
    Dim myResult As Variant
    
    'Act:
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Seq")
Private Sub Test04a_AddStringByAddRange()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array("H", "e", "l", "l", "o")

    '@Ignore IntegerDataType
    
    Dim mySeq As Seq = Seq.Deb.AddRange("Hello")
    Dim myResult As Variant
    
    'Act:
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test04b_AddArrayByAddRange()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)

    '@Ignore IntegerDataType
    
    Dim mySeq As Seq = Seq.Deb.AddKnownRange(Array(10, 20, 30, 40, 50, 60, 70, 80, 90))
    Dim myResult As Variant
    
    'Act:
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test04c_AddCollectionByAddRange()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)

    '@Ignore IntegerDataType
    Dim myC As Collection = New Collection
    With myC
    
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    
    End With
    Dim mySeq As Seq = Seq.Deb.AddRange(myC)
    Dim myResult As Variant
    
    'Act:
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test04d_AddDictionaryByAddRange()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)

    '@Ignore IntegerDataType
    
   
    
    Dim myHkvp As Hkvp = Hkvp.Deb.AddKnownArrayPairs(Array("a", "b", "c", "d", "e", "f", "g", "h ", "i"), Array(10, 20, 30, 40, 50, 60, 70, 80, 90))
    
    Dim mySeq As Seq = Seq.Deb(myHkvp)
    Dim myResult As Variant
    'Act:
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test05_Clear()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedBefore  As Long = 9
    Dim myExpectedAfter As Long = 0
    
    Dim myHkvp As Hkvp = Hkvp.Deb.AddPairs(Array("a", "b", "c", "d", "e", "f", "g", "h ", "i"), Array(10, 20, 30, 40, 50, 60, 70, 80, 90))
    
    Dim mySeq As Seq = Seq.Deb(myHkvp)
    
    'Act:
    Dim myResultBefore As Long = mySeq.Count
    mySeq.Clear
    Dim myResultAfter As Long = mySeq.Count

    'Assert:
    AssertStrictAreEqual myExpectedBefore, myResultBefore, myProcedureName
    AssertStrictAreEqual myExpectedAfter, myResultAfter, myProcedureName
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test06_Clone()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    
    Dim myHkvp As Hkvp = Hkvp.Deb.AddPairs(Array("a", "b", "c", "d", "e", "f", "g", "h ", "i"), Array(10, 20, 30, 40, 50, 60, 70, 80, 90))
    
    Dim mySeq As Seq = Seq.Deb(myHkvp)
    
    'Act:
    Dim myResult As Variant = mySeq.Clone.ToArray
    mySeq.Clear
    Dim myResultAfter As Long = mySeq.Count

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
   
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Seq")
Private Sub Test07a_ContainsTrue()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)
    
    Dim myResult As Boolean
    
    'Act:
    myResult = mySeq.Contains(10)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test07b_ContainsFalse()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)
    
    Dim myResult As Boolean
    
    'Act:
    myResult = mySeq.Contains(100)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Seq")
Private Sub Test07c_HoldsItemTrue()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)
    
    Dim myResult As Boolean
    
    'Act:
    myResult = mySeq.HoldsItem(10)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test07d_HoldsValueFalse()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)
    
    Dim myResult As Boolean
    
    'Act:
    myResult = mySeq.HoldsItem(100)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test07e_LacksValueTrue()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)
    
    Dim myResult As Boolean
    
    'Act:
    myResult = mySeq.LacksItem(100)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test07f_LacksValueFalse()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)
    
    Dim myResult As Boolean
    
    'Act:
    myResult = mySeq.LacksItem(10)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test08a_CopyToAllArray()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 40, 50)

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)

    Dim myResult(0 To 4) As Variant

    'Act:
    mySeq.CopyTo myResult

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
TestExit:
    Exit Sub

TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("Seq")
Private Sub Test08b_CopyToArrayStart()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(1, 2, 3, 10, 20, 30, 40, 50)

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)

    Dim myResult(0 To 7) As Long
    myResult(0) = 1
    myResult(1) = 2
    myResult(2) = 3

    'Act:
    mySeq.CopyTo myResult, 3

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
TestExit:
    Exit Sub

TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub


'@TestMethod("Seq")
Private Sub Test08c_CopyToSeqStartArrayRun()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(1, 2, 3, 30, 40, 50, 7, 8, 9, 10)

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)

    Dim myResult(0 To 9) As Long
    myResult(0) = 1
    myResult(1) = 2
    myResult(2) = 3
    myResult(3) = 4
    myResult(4) = 5
    myResult(5) = 6
    myResult(6) = 7
    myResult(7) = 8
    myResult(8) = 9
    myResult(9) = 10

    'Act:
    'Remeber, the array will be zero based but the seq wil be 1 based
    mySeq.CopyTo 3, myResult, 3, 3

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
TestExit:
    Exit Sub

TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub


'@TestMethod("Seq")
Private Sub Test09a_GetRangeViaSlice()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(20, 30, 40)

    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50, 60, 70, 80)

    Dim myResult As Variant
    

    'Act:
    myResult = mySeq.Slice(2, 3).ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
TestExit:
    Exit Sub

TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub


'@TestMethod("Seq")
Private Sub Test09b_GetRangeViaGetRange()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(20, 30, 40)

    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50, 60, 70, 80)

    Dim myResult As Variant
    

    'Act:
    myResult = mySeq.GetRange(2, 3).ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
TestExit:
    Exit Sub

TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("Seq")
Private Sub Test09c_GetRangeViaSliceWithEndIndex()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(20, 30, 40)

    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50, 60, 70, 80)

    Dim myResult As Variant
    

    'Act:
    myResult = mySeq.Slice(2, ipendindex:= 4).ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
TestExit:
    Exit Sub

TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("Seq")
Private Sub Test09d_GetRangeViaGetRangeWithEndIndex()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(20, 30, 40)

    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50, 60, 70, 80)

    Dim myResult As Variant
    

    'Act:
    myResult = mySeq.GetRange(2, ipendindex:= 4).ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
TestExit:
    Exit Sub

TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub



'@TestMethod("Seq")
Private Sub Test10a_FirstIndexOfInWholeSeq()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 3

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 30, 30, 30, 40, 50)
    
    Dim myResult As Long
    
    'Act:
    myResult = mySeq.IndexOf(30)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test10b_FirstIndexOfInWithStart()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 3

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 30, 30, 30, 40, 50)
    
    Dim myResult As Long
    
    'Act:
    myResult = mySeq.IndexOf(30, 2)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Seq")
Private Sub Test10c_FirstIndexOfInWithStartRun()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 3

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 30, 30, 30, 40, 50)
    
    Dim myResult As Long
    
    'Act:
    myResult = mySeq.IndexOf(30, 1, 5)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test10d_FirstIndexOfInWithStartEnd()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 3

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 30, 30, 30, 40, 50)
    
    Dim myResult As Long
    
    'Act:
    myResult = mySeq.IndexOf(30, 1, ipendindex:=5)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Seq")
Private Sub Test11A_InsertSingleItem()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 400, 40, 50, 60, 70)

    Dim myResult As Variant
    
    'Act:
    Dim mySeq As Seq
    Set mySeq = Seq.Deb(Array(10, 20, 30, 40, 50, 60, 70))
    mySeq.Insert 4, 400
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test11b_InsertMultipleItems()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 400, 500, 600, 700, 800, 40, 50, 60, 70)

    Dim myResult As Variant
    
    'Act:
    Dim mySeq As Seq
    Set mySeq = Seq.Deb(Array(10, 20, 30, 40, 50, 60, 70))
    mySeq.Insert 4, 400, 500, 600, 700, 800
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test11c_InsertRangeString()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, "H", "e", "l", "l", "o", 40, 50, 60, 70)

    Dim myResult As Variant
    
    'Act:
    Dim mySeq As Seq
    Set mySeq = Seq.Deb(Array(10, 20, 30, 40, 50, 60, 70))
    mySeq.InsertRange 4, "Hello"
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test11D_InsertRangeArray()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 400, 500, 600, 700, 800, 40, 50, 60, 70)

    Dim myResult As Variant
    
    'Act:
    Dim mySeq As Seq
    Set mySeq = Seq.Deb(Array(10, 20, 30, 40, 50, 60, 70))
    mySeq.InsertRange 4, Array(400, 500, 600, 700, 800)
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test11e_InsertRangeCollection()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 400, 500, 600, 700, 800, 40, 50, 60, 70)
    
    Dim myC As Collection = New Collection
    With myC
    
        .Add 400
        .Add 500
        .Add 600
        .Add 700
        .Add 800
        
    End With
    Dim myResult As Variant
    
    'Act:
    Dim mySeq As Seq
    Set mySeq = Seq.Deb(Array(10, 20, 30, 40, 50, 60, 70))
    mySeq.InsertRange 4, myC
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test11f_InsertRangeStack()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 400, 500, 600, 700, 800, 40, 50, 60, 70)
    
    Dim myS As Stack = Stack.Deb
    With myS
    
        .Push 400
        .Push 500
        .Push 600
        .Push 700
        .Push 800
        
    End With
    Dim myResult As Variant
    
    'Act:
    Dim mySeq As Seq
    Set mySeq = Seq.Deb(Array(10, 20, 30, 40, 50, 60, 70))
    mySeq.InsertRange 4, myS
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Seq")
Private Sub Test11g_InsertRangeDictionary()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 400, 500, 600, 700, 800, 40, 50, 60, 70)
    
    Dim myH As Hkvp = Hkvp.Deb.AddPairs(Array("a", "b", "c", "d", "e"), Array(400, 500, 600, 700, 800))
    
    Dim myResult As Variant
    
    'Act:
    Dim mySeq As Seq
    Set mySeq = Seq.Deb(Array(10, 20, 30, 40, 50, 60, 70))
    mySeq.InsertRange 4, myH
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub



'@TestMethod("Seq")
Private Sub Test12a_LastIndexOf()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 6

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 30, 30, 30, 40, 50)
    
    Dim myResult As Long
    
    'Act:
    myResult = mySeq.LastIndexOf(30)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


' '@TestMethod("Seq")
' Private Sub Test19_IndexOfFromItem1()

'     #If twinbasic Then
    
'         myProcedureName = CurrentProcedureName
'         myComponentName = CurrentComponentName
        
        
'     #Else
    
'         myProcedureName = ErrEx.LiveCallstack.ProcedureName
'         myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
'     #End If
    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Long
'     myExpected = 3

'     '@Ignore IntegerDataType
'     Dim mySeq As Seq
'     Set mySeq = Seq.Deb.Add(10, 20, 30, 40, 50)
    
'     Dim myResult As Long
    
'     'Act:
'     myResult = mySeq.IndexOf(30)

'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
' End Sub

'@TestMethod("Seq")
Private Sub Test13a_AssignItemTwoPrimitive()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If


   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 300, 30, 40, 50)

    '@Ignore IntegerDataType
    Dim mySeq As Seq = Seq.Deb.AddItems(10, 20, 30, 40, 50)
    
    Dim myResult As Variant
    
    'Act:
    mySeq.Item(2) = 300
    myResult = mySeq.ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test13b_AssignAndGetItemTwoArray()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Seq
    Set myExpected = Seq.Deb.AddItems(100, 200, 300, 400, 500)

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)
    
    Dim myResult As Seq
    
    'Act:
    mySeq.Item(2) = Seq.Deb.AddRange(Array(100, 200, 300, 400, 500))
    Set myResult = mySeq.Item(2)

    'Assert:
    Assert.Exact.SequenceEquals myExpected.ToArray, myResult.ToArray, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Seq")
Private Sub Test14a_RemoveAll()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 0

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)

    Dim myResult As Long
    
    'Act:
    If mySeq.Count <> 5 Then Err.Raise 17
    mySeq.RemoveAll
    myResult = mySeq.Count
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test14b_RemoveItemsIndividually()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 30, 50)

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)

    Dim myResult As Variant
    
    'Act:
    mySeq.RemoveAt 2
    mySeq.RemoveAt 3
    myResult = mySeq.ToArray
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test14c_RemoveBlockWithinLastIndex()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 50, 60, 70, 80, 90)

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50, 60, 70, 80, 90)

    Dim myResult As Variant
    
    'Act:
    mySeq.RemoveAt 2, 3
    myResult = mySeq.ToArray
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test14d_RemoveBlockPastEndOfItems()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10)

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50, 60, 70, 80, 90)

    Dim myResult As Variant
    
    'Act:
    mySeq.RemoveAt 2, 2048
   
    myResult = mySeq.ToArray
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Seq")
Private Sub Test15a_SliceItem1ToItem3()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(30, 40, 50, 60, 70, 80, 90)

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50, 60, 70, 80, 90)

    Dim myResult As Variant

    'Act:
     myResult = mySeq.Slice(3, 7).ToArray

    'Assert:
    Assert.Exact.SequenceEquals myExpected, myResult, myProcedureName
    Exit Sub
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("Seq")
Private Sub Test15b_SliceItem3Run4ToItem3()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(30, 40, 50, 60)

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50, 60, 70, 80, 90)

    Dim myResult As Variant

    'Act:
     myResult = mySeq.Slice(3, iprun:=4).ToArray

    'Assert:
    Assert.Exact.SequenceEquals myExpected, myResult, myProcedureName
    Exit Sub
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("Seq")
Private Sub Test15c_SliceItem3RunPastEnd()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(30, 40, 50, 60, 70, 80, 90, Empty, Empty, Empty, Empty, Empty)

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50, 60, 70, 80, 90)

    Dim myResult As Variant

    'Act:
     myResult = mySeq.Slice(3, iprun:=12).ToArray

    'Assert:
    Assert.Exact.SequenceEquals myExpected, myResult, myProcedureName
    Exit Sub
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub



'@TestMethod("InsertAt")
Private Sub Test16a_InsertAtItem1()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 70, 30, 40, 50)

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 50)
    
    Dim myResult As Seq
    
    'Act:
    Set myResult = mySeq.Insert(3, 70)

    'Assert:
    Assert.Exact.SequenceEquals myExpected, myResult.ToArray, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("InsertRange")
Private Sub Test16b_InsertRangeFivetemsFromItem1()
   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 15, 16, 17, 18, 19, 30, 40, 50)

    '@Ignore IntegerDataType
    Dim mySeq As Seq = Seq.Deb.AddItems(10, 20, 30, 40, 50)

    Dim myResult As Variant

    'Act:
     myResult = mySeq.InsertRange(3, Array(15, 16, 17, 18, 19)).ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub

TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub



'@TestMethod("LastIndexOf")
' Private Sub Test22_LastIndexOfWholeSeq()

'     #If twinbasic Then
    
'         myProcedureName = CurrentProcedureName
'         myComponentName = CurrentComponentName
        
        
'     #Else
    
'         myProcedureName = ErrEx.LiveCallstack.ProcedureName
'         myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
'     #End If
    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Long
'     myExpected = 7

'     '@Ignore IntegerDataType
'     Dim mySeq As Seq
'     Set mySeq = Seq.Deb.Add(10, 20, 30, 40, 40, 40, 40, 50)
    
'     Dim myResult As Long
    
'     'Act:
'     myResult = mySeq.LastIndexof(40)

'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
' End Sub

' '@TestMethod("LastIndexOf")
' Private Sub Test23_LastIndexOfStartItem4()

'     #If twinbasic Then
    
'         myProcedureName = CurrentProcedureName
'         myComponentName = CurrentComponentName
        
        
'     #Else
    
'         myProcedureName = ErrEx.LiveCallstack.ProcedureName
'         myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
'     #End If
    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Long
'     myExpected = 7

'     '@Ignore IntegerDataType
'     Dim mySeq As Seq
'     Set mySeq = Seq.Deb.Add(10, 20, 30, 40, 40, 40, 40, 50)
    
'     Dim myResult As Long
    
'     'Act:
'     myResult = mySeq.LastIndexof(40, 4)

'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
' End Sub

'     '@TestMethod("LastIndexOf")
'     Private Sub Test24_LastIndexOfStartItem1EndItem4()
'        'On Error GoTo TestFail
    
'         'Arrange:
'         Dim myExpected  As Long
'         myExpected = 4

'         '@Ignore IntegerDataType
'         Dim mySeq As Seq
'         Set mySeq = Seq.Deb.Add(10, 20, 30, 40, 40, 40, 40, 50)
    
'         Dim myResult As Long
    
'         'Act:
'         myResult = mySeq.LastIndexof(40, 1, 4)

'         'Assert:
'         AssertStrictAreEqual myExpected, myResult  , myProcedureName
    
' TestExit:
'         Exit Sub
    
' TestFail:
'         Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
'         Resume TestExit
    
'     End Sub

' @TestMethod("RemoveValue")
' Private Sub Test25_RemoveFirstOf40()

'     #If twinbasic Then
'         myProcedureName = CurrentProcedureName
'         myComponentName = CurrentComponentName
'     #Else
'         myProcedureName = ErrEx.LiveCallstack.ProcedureName
'         myComponentName = ErrEx.LiveCallstack.ModuleName
'     #End If
    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Seq
'     Set myExpected = Seq.Deb.Add(10, 20, 30, 40, 40, 40, 50)

'     '@Ignore IntegerDataType
'     Dim mySeq As Seq
'     Set mySeq = Seq.Deb.Add(10, 20, 30, 40, 40, 40, 40, 50)
    
'     Dim myResult As Seq
    
'     'Act:
'     Set myResult = mySeq.RemoveFirstOf(40)

'     'Assert:
'     Assert.Exact.SequenceEquals myExpected.ToArray, myResult.ToArray, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
' End Sub

'@TestMethod("RemoveAt")
Private Sub Test17a_RemoveAtItem4()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Seq
    Set myExpected = Seq.Deb.AddItems(10, 20, 30, 20, 40, 40, 50)

    '@Ignore IntegerDataType
    Dim mySeq As Seq
    Set mySeq = Seq.Deb.AddItems(10, 20, 30, 40, 20, 40, 40, 50)
    
    Dim myResult As Seq
    
    'Act:
    Set myResult = mySeq.RemoveAt(4)

    'Assert:
    Assert.Exact.SequenceEquals myExpected.ToArray, myResult.ToArray, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'     '@TestMethod("RemoveRange")
'     Private Sub Test27_RemoveRangeItem3Count4()
'        'On Error GoTo TestFail
    
'         'Arrange:
'         Dim myExpected  As Seq
'         Set myExpected = Seq.Deb.Add(10, 20, 30, 50)

'         '@Ignore IntegerDataType
'         Dim mySeq As Seq
'         Set mySeq = Seq.Deb.Add(10, 20, 30, 40, 40, 40, 40, 50)
    
'         Dim myResult As Seq
    
'         'Act:
'         Set myResult = mySeq.RemoveRange(3, 4)

'         'Assert:
'         Assert.Exact.SequenceEquals myExpected.ToArray, myResult.ToArray, myProcedureName
    
' TestExit:
'         Exit Sub
    
' TestFail:
'         Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
'         Resume TestExit
    
'     End Sub


' '@TestMethod("Reverse")
' Private Sub Test28_ReverseAll()

'     #If twinbasic Then
    
'         myProcedureName = CurrentProcedureName
'         myComponentName = CurrentComponentName
        
        
'     #Else
    
'         myProcedureName = ErrEx.LiveCallstack.ProcedureName
'         myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
'     #End If
    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Seq
'     Set myExpected = Seq.Deb.Add(50, 40, 40, 40, 40, 30, 20, 10)

'     '@Ignore IntegerDataType
'     Dim mySeq As Seq
'     Set mySeq = Seq.Deb.Add(10, 20, 30, 40, 40, 40, 40, 50)
    
'     Dim myResult As Seq
    
'     'Act:
'     Set myResult = mySeq.Reverse

'     'Assert:
'     Assert.Exact.SequenceEquals myExpected.ToArray, myResult.ToArray, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
' End Sub


'     '@TestMethod("Reverse")
'     Private Sub Test29_ReverseItem1Count4()
'        'On Error GoTo TestFail
    
'         'Arrange:
'         Dim myExpected  As Seq
'         Set myExpected = Seq.Deb.Add(10, 40, 40, 30, 20, 40, 40, 50)

'         '@Ignore IntegerDataType
'         Dim mySeq As Seq
'         Set mySeq = Seq.Deb.Add(10, 20, 30, 40, 40, 40, 40, 50)
    
'         Dim myResult As Seq
    
'         'Act:
'         Set myResult = mySeq.Reverse(1, 4)

'         'Assert:
'         Assert.Exact.SequenceEquals myExpected.ToArray, myResult.ToArray, myProcedureName
    
' TestExit:
'         Exit Sub
    
' TestFail:
'         Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
'         Resume TestExit
    
'     End Sub

'     '@TestMethod("Reverse")
'     Private Sub Test30_SetRangeItem1ToFouritems()
'        'On Error GoTo TestFail
    
'         'Arrange:
'         Dim myExpected  As Seq
'         Set myExpected = Seq.Deb.Add(10, 50, 50, 50, 50, 40, 40, 50)

'         '@Ignore IntegerDataType
'         Dim mySeq As Seq
'         Set mySeq = Seq.Deb.Add(10, 20, 30, 40, 40, 40, 40, 50)
    
'         Dim myResult As Seq
    
'         'Act:
'         Set myResult = mySeq.SetRange(2, Array(50, 50, 50, 50))

'         'Assert:
'         Assert.Exact.SequenceEquals myExpected.ToArray, myResult.ToArray, myPlace & CurrentProcedureName
    
' TestExit:
'         Exit Sub
    
' TestFail:
'         Debug.Print myComponentName, myProcedureName, " raised an Error: #" & Err.Number & " - " & Err.Description
'         Resume TestExit
    
'     End Sub
