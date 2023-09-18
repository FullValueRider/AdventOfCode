Attribute VB_Name = "TestRank"
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


Public Sub RankTests()

#If twinbasic Then

    Debug.Print CurrentProcedureName ; vbTab, vbTab, vbTab,
    
#Else

    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
    
#End If

    Test01_IsRankObject
 
    Test02a_IsQueryableEmptyRankIsFalse
    Test02b_IsNotQueryableEmptyRankIsTrue
    Test02c_IsQueryableValidRankIsTrue
    Test02d_IsNotQueryableValidRankIsFalse
    
    Test03a_FirstIndexIsOne
   
    Test04a_LastIndexIsTwo
   
    Test05a_CountIsTwo
   
    Debug.Print "Testing completed"

End Sub



'@TestMethod("Rank")
Private Sub Test01_IsRankObject()
    
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

    '@Ignore IntegerDataType
    Dim myRank As Rank
    Set myRank = Rank.Deb
    
    Dim myResult As Variant
    ReDim myResult(0 To 2)
    
    'Act:
    myResult(0) = VBA.IsObject(myRank)
    myResult(1) = "Rank" = VBA.TypeName(myRank)
    myResult(2) = "Rank" = myRank.TypeName
    
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Rank")
Private Sub Test02a_IsQueryableEmptyRankIsFalse()

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
    Dim myRank As Rank = Rank.Deb
    
    Dim myResult As Boolean
    
    'Act:
    myResult = myRank.IsQueryable

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Rank")
Private Sub Test02b_IsNotQueryableEmptyRankIsTrue()

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
    Dim myRank As Rank = Rank.Deb
    
    Dim myResult As Boolean
    
    'Act:
    myResult = myRank.IsNotQueryable

    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Rank")
Private Sub Test02c_IsQueryableValidRankIsTrue()
    
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

    Dim myRank As Rank = Rank.Deb(1, 2)
    
    Dim myResult As Boolean
    
    'Act:
    myResult = myRank.IsQueryable

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Rank")
Private Sub Test02d_IsNotQueryableValidRankIsFalse()
    
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
    Dim myRank As Rank = Rank.Deb(1, 2)
    
    Dim myResult As Boolean
    
    'Act:
    myResult = myRank.IsNotQueryable

    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Rank")
Private Sub Test03a_FirstIndexIsOne()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long = 1

    '@Ignore IntegerDataType
    Dim myRank As Rank = Rank.Deb(1, 2)
    
    Dim myResult As Long
    
    'Act:
    myResult = myRank.FirstIndex

    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



'@TestMethod("Rank")
Private Sub Test04a_LastIndexIsTwo()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long = 2

    '@Ignore IntegerDataType
    Dim myRank As Rank = Rank.Deb(1, 2)
    
    Dim myResult As Long
    
    'Act:
    myResult = myRank.LastIndex

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



'@TestMethod("Rank")
Private Sub Test05a_CountIsTwo()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long = 2

    '@Ignore IntegerDataType
    Dim myRank As Rank = Rank.Deb(1, 2)
    
    Dim myResult As Long
    
    'Act:
    myResult = myRank.Count

    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub