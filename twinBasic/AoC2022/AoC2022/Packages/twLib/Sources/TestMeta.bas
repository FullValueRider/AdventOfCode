Attribute VB_Name = "TestMeta"
Option Explicit
Option Private Module
'@IgnoreModule
'@TestModule
'@Folder("Tests")

#If twinbasic Then
    'Do nothing
#Else



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


Public Sub MetaTests()

#If twinbasic Then
    Debug.Print CurrentProcedureName ; vbTab, vbTab, vbTab,
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
#End If

    Test01_IsMetaObject
    
    Test02a_IsQueryableEmptyRanksIsFalse
    Test02b_IsNotQueryableEmptyRanksIsTrue
    
    ' Test03a_ItemArrayOfRankGetItemTwo
    ' Test03b_ItemArrayOfRankSetItemTwoIsTrue
    ' Test04a_CountOfRanksIsThree
    Test04b_IsNotQueryableIsTrue
    
    Test05a_ItemsOfIsQueryable
    Test05b_ItemsOfIsNotQueryableIsEmpty
    
    Test06a_HasOneItemIsTrue
    Test06b_HasOneItemIsFalse
    Test06c_HasItemsIsFalse
    Test06d_HasItemsIsTrue
    Test06e_HasAnyItemIsTrue
    Test06f_HasAnyItemsIsFalse
    
    ' Test09a_HasRankTwoIsTrue
    ' Test09b_HasRankFourIsFalse
    
   ' Test10a_ForEachRank
    
    Debug.Print "Testing completed"

End Sub

'@TestMethod("Meta")
Private Sub Test01_IsMetaObject()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant = Array(True, True, True)

    Dim myMeta As Meta = Meta.Deb(Array(10, 10, 10))
    
    Dim myResult As Variant
    ReDim myResult(0 To 2)
    
    'Act:
    myResult(0) = VBA.IsObject(myMeta)
    myResult(1) = "Meta" = TypeName(myMeta)
    myResult(2) = "Meta" = myMeta.TypeName

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Meta")
Private Sub Test02a_IsQueryableEmptyRanksIsFalse()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean = False

    '@Ignore IntegerDataType
    Dim myMeta As Meta = Meta.Deb(Array())
    
    Dim myResult As Boolean = True
    
    'Act:
    myResult = myMeta.IsQueryable

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Meta")
Private Sub Test02b_IsNotQueryableEmptyRanksIsTrue()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean = True

    Dim myMeta As Meta = Meta.Deb(Array())
    
    Dim myResult As Boolean
    
    'Act:
    myResult = myMeta.IsNotQueryable

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Meta")
Private Sub Test03a_ItemArrayOfRankGetItemTwo()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Rank = Rank.Deb(2, 4)

    '@Ignore IntegerDataType
    Dim myArray(1 To 3, 2 To 4, 3 To 5) As Variant
    Dim myMeta As Meta = Meta.Deb(myArray)
    
    Dim myResult As Rank
    
    'Act:
    Set myResult = myMeta.Ranks.Item(2)

    'Assert:
    AssertStrictSequenceEquals myExpected.ToArray, myResult.ToArray, myProcedureName
    AssertStrictAreEqual myExpected.FirstIndex, myResult.FirstIndex, myProcedureName
    AssertStrictAreEqual myExpected.LastIndex, myResult.LastIndex, myProcedureName
    AssertStrictAreEqual myExpected.Count, myResult.Count, myProcedureName
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Meta")
Private Sub Test03b_ItemArrayOfRankSetItemTwoIsTrue()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Rank = Rank.Deb(6, 11)

    Dim myArray(1 To 3, 2 To 4, 3 To 5) As Variant
    Dim myMeta As Meta = Meta.Deb(myArray)
    
    Dim myResult As Rank
    
    'Act:
    Set myMeta.Ranks(2) = Rank.Deb(6, 11)
    Set myResult = myMeta.Ranks(2)

    'Assert:
    AssertStrictSequenceEquals myExpected.ToArray, myResult.ToArray, myProcedureName
    AssertStrictAreEqual myExpected.FirstIndex, myResult.FirstIndex, myProcedureName
    AssertStrictAreEqual myExpected.LastIndex, myResult.LastIndex, myProcedureName
    AssertStrictAreEqual myExpected.Count, myResult.Count, myProcedureName
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Meta")
Private Sub Test04a_CountOfRanksIsThree()
    
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

    Dim myArray(1 To 3, 2 To 4, 3 To 5) As Variant
    Dim myMeta As Meta = Meta.Deb(myArray)
    
    Dim myResult As Long
    
    'Act:
   
    myResult = myMeta.RanksCount

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Meta")
Private Sub Test04b_IsNotQueryableIsTrue()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean = True

    Dim myMeta As Meta = Meta(Array())
    
    Dim myResult As Boolean
    
    'Act:
    myResult = myMeta.IsNotQueryable

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Meta")
Private Sub Test05a_ItemsOfIsQueryable()
    
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
    myExpected = Array(Rank.Deb(1, 3)) ', Rank.Deb(2, 4), Rank.Deb(3, 5))

    Dim myArray(1 To 3) As Variant ', 2 To 4, 3 To 5) As Variant
    Dim myResult As Variant
    
    'Act:
   myResult = Meta(myArray).Ranks.ToArray

    'Assert:
    
    AssertStrictSequenceEquals myExpected(0).ToArray, myResult(0).ToArray, myProcedureName
   ' AssertStrictSequenceEquals myExpected(1).ToArray, myResult(1).ToArray, myProcedureName
   ' AssertStrictSequenceEquals myExpected(2).ToArray, myResult(2).ToArray, myProcedureName
    
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Meta")
Private Sub Test05b_ItemsOfIsNotQueryableIsEmpty()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant = Array(Rank.Deb(0, -1))
    

    Dim myResult As Variant
    
    'Act:
    myResult = Meta(Array()).Ranks.ToArray

    'Assert:
   AssertExactAreEqual myExpected(0).firstindex, myResult(0).firstindex, myProcedureName
   AssertExactAreEqual myExpected(0).lastindex, myResult(0).lastindex, myProcedureName
   AssertExactAreEqual myExpected(0).count, myResult(0).count, myProcedureName
   
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

' '@TestMethod("Meta")
' Private Sub Test05c_KeysOfIsQueryable()
    
'     #If twinbasic Then
    
'         myProcedureName = CurrentProcedureName
'         myComponentName = CurrentComponentName


'     #Else

'         myProcedureName = ErrEx.LiveCallstack.ProcedureName
'         myComponentName = ErrEx.LiveCallstack.ModuleName
        

'     #End If
    
'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Variant
'     myExpected = Array(1, 2, 3)

    
'     Dim myResult As Variant
    
'     'Act:
'    myResult = Meta.Deb(Rank(1, 2), Rank(2, 3), Rank(3, 4)).Keys

'     'Assert:
'     AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub

'@TestMethod("Meta")
' Private Sub Test05d_ItemsOfIsNotQueryableIsEmpty()
    
'     #If twinbasic Then
    
'         myProcedureName = CurrentProcedureName
'         myComponentName = CurrentComponentName


'     #Else

'         myProcedureName = ErrEx.LiveCallstack.ProcedureName
'         myComponentName = ErrEx.LiveCallstack.ModuleName
        

'     #End If
    
'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Variant
'     myExpected = Empty

    
'     Dim myResult As Variant
    
'     'Act:
'    myResult = Meta.Deb.Keys

'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
   
' TestExit:
'     Exit Sub
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub

'@TestMethod("Meta")
Private Sub Test06a_HasOneItemIsTrue()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean = True

    Dim myMeta As Meta = Meta.Deb(Array(10))
    
    Dim myResult As Boolean
   
    'Act:
   myResult = myMeta.HasOneItem

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
   
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Meta")
Private Sub Test06b_HasOneItemIsFalse()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean = False
    
    Dim myMeta As Meta = Meta.Deb(Array(10, 20))
    
    Dim myResult As Boolean
   
    'Act:
   myResult = myMeta.HasOneItem

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
   
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Meta")
Private Sub Test06c_HasItemsIsFalse()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean = False
   
    Dim myMeta As Meta = Meta.Deb(Array(10))
    
    Dim myResult As Boolean
   
    'Act:
   myResult = myMeta.HasItems

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
   
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Meta")
Private Sub Test06d_HasItemsIsTrue()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean = True
    Dim myArray(1 To 3) As Variant ', 2 To 4, 3 To 5) As Variant
    Dim myMeta As Meta = Meta.Deb(myArray)
    
    Dim myResult As Boolean
   
    'Act:
    myResult = myMeta.HasItems

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
   
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Meta")
Private Sub Test06e_HasAnyItemIsTrue()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean = True
    Dim myArray(1 To 3) As Variant
    Dim myMeta As Meta = Meta.Deb(myArray)
    
    Dim myResult As Boolean
   
    'Act:
   myResult = myMeta.HasAnyItems

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
   
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Meta")
Private Sub Test06f_HasAnyItemsIsFalse()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean = False
    
    Dim myMeta As Meta = Meta.Deb(Array())
    
    Dim myResult As Boolean
   
    'Act:
   myResult = myMeta.HasAnyItems

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
   
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



' '@TestMethod("Meta")
' Private Sub Test07b_AddRankTwoRanks()
    
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
'     myExpected = 2

'     Dim myMeta As Meta
'     Set myMeta = Meta.Deb.AddRank(1, 2).AddRank(2, 3)
    
'     Dim myResult As Long
   
'     'Act:
'    myResult = myMeta.RankCount

'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
   
' TestExit:
'     Exit Sub
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub

' '@TestMethod("Meta")
' Private Sub Test08a_RemoveRankIsTwoRanks()
    
'     #If twinbasic Then
    
'         myProcedureName = CurrentProcedureName
'         myComponentName = CurrentComponentName


'     #Else

'         myProcedureName = ErrEx.LiveCallstack.ProcedureName
'         myComponentName = ErrEx.LiveCallstack.ModuleName
        

'     #End If
    
'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Variant
'     myExpected = Array(Rank(1, 2), Rank(3, 4))

'     Dim myMeta As Meta
'     Set myMeta = Meta.Deb.AddRank(1, 2).AddRank(2, 3).AddRank(3, 4)
    
'     Dim myResult As Variant
   
'     'Act:
'    myResult = myMeta.Remove(2).Items

'     'Assert:
'     AssertStrictSequenceEquals myExpected(0).ToArray, myResult(0).ToArray, myProcedureName
'     AssertStrictSequenceEquals myExpected(1).ToArray, myResult(1).ToArray, myProcedureName
' TestExit:
'     Exit Sub
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub

'@TestMethod("Meta")
Private Sub Test09a_HasRankTwoIsTrue()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean = True
    Dim myArray(1 To 3, 2 To 4, 3 To 5) As Variant
    Dim myMeta As Meta = Meta.Deb(myArray)
    
    Dim myResult As Boolean
   
    'Act:
   myResult = myMeta.HasRank(2)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
   
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Meta")
Private Sub Test09b_HasRankFourIsFalse()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean = False
    Dim myArray(1 To 3, 2 To 4, 3 To 5) As Variant
    Dim myMeta As Meta = Meta.Deb(myArray)
    
    Dim myResult As Boolean
   
    'Act:
   myResult = myMeta.HasRank(4)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
   
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Meta")
Private Sub Test10a_ForEachRank()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant = Array(1, 9, 3)
    Dim myArray(1 To 2, 9 To 32, 3 To 4) As Variant
    Dim myMeta As Meta = Meta.Deb(myArray)
    
    Dim myResult As Variant
   
    'Act:
    ReDim myResult(0 To 2)
    Dim myIndex As Long = 0
    Dim myRank As Variant
    For Each myRank In myMeta
        myResult(myIndex) = myRank.FirstIndex
        myIndex = myIndex + 1
    Next

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
   
TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
