Attribute VB_Name = "TestHkvp"
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

Public Sub HkvpTests()
    
#If twinbasic Then
    Debug.Print CurrentProcedureName ; vbTab, vbTab, vbTab,
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
#End If

    Test01_IsHkvp
    Test02_CountKeysItems
    Test03_Clear
    Test04_RemoveByKey
    Test05_RemoveByIndex
    Test06a_Exists
    Test06b_HoldsKey
    Test06c_LacksKey
    Test06d_HoldsItem
    Test06e_LacksItem
    
    Test07_IndexByKey
    Test08_GetItemByIndex
    Test09_LetItemByIndex
    Test10_SetItemByIndex
    Test11_GetItem
    Test12_LetItem
    Test13_SetItem
    Test14_DuplicateKeys
    Test15_UniqueKeys
    'Test16_Reverse  reversing a dictionary doewsn't make sense
    
    Debug.Print "Testing completed"

End Sub

'@TestMethod("Hkvp")
Private Sub Test01_IsHkvp()

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
    Dim myHkvp As Hkvp = Hkvp.Deb
    Dim myResult(0 To 2)  As Boolean

    'Act:
    myResult(0) = VBA.IsObject(myHkvp)
    myResult(1) = "Hkvp" = TypeName(myHkvp)
    myResult(2) = "Hkvp" = myHkvp.TypeName
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test02_CountKeysItems()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
       'On Error GoTo TestFail
    
        'Arrange:
        Dim myExpectedCount As Long = 6
        Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day", " ")
        Dim myExpectedItems As Variant = Array(10, 20, 30, 40, 50, 60)
        Dim myHkvp As Hkvp = Hkvp.Deb
        
        With myHkvp
            .Add "Hello", 10
            .Add "World", 20
            .Add "Its", 30
            .Add "A", 40
            .Add "Nice", 50
            .Add "Day", 60
        End With
        
        Dim myResultKeys  As Variant
        Dim myResultItems As Variant
        Dim myResultCount As Long
        
        'Act:
        myResultCount = myHkvp.Count
        myResultKeys = myHkvp.Keys
        myResultItems = myHkvp.Items
       
        'Assert.Strict:
        AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
        AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
        AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
        
        TestExit:
        Exit Sub
        TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
    
'@TestMethod("Hkvp")
Private Sub Test03_Clear()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
       'On Error GoTo TestFail
    
        'Arrange:
        Dim myExpected As Long = 0
        
        
        Dim myHkvp As Hkvp = Hkvp.Deb
        
        With myHkvp
            .Add "Hello", 10
            .Add "World", 20
            .Add "Its", 30
            .Add "A", 40
            .Add "Nice", 50
            .Add "Day", 60
        End With
        
        Dim myResult As Long
        
        'Act:
        myHkvp.Clear
        myResult = myHkvp.Count
        
        
        'Assert.Strict:
        AssertStrictAreEqual myExpected, myResult, myProcedureName
        
        TestExit:
        Exit Sub
        TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
    
'@TestMethod("Hkvp")
Private Sub Test04_RemoveByKey()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 4
    Dim myExpectedKeys As Variant = Split("World A Nice Day", " ")
    Dim myExpectedItems As Variant = Array(20, 40, 50, 60)
    
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResultKeys  As Variant
    Dim myResultItems As Variant
    Dim myResultCount As Long
    
    'Act:
    myHkvp.Remove "Hello"
    myHkvp.Remove "Its"
    myResultCount = myHkvp.Count
    myResultKeys = myHkvp.Keys
    myResultItems = myHkvp.Items
    
    'Assert.Strict:
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test05_RemoveByIndex()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 4
    Dim myExpectedKeys As Variant = Split("World Its Nice Day", " ")
    Dim myExpectedItems As Variant = Array(20, 30, 50, 60)
    
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResultKeys  As Variant
    Dim myResultItems As Variant
    Dim myResultCount As Long
    
    'Act:
    myHkvp.RemoveByIndex 0&
    myHkvp.RemoveByIndex 2&
    myResultCount = myHkvp.Count
    myResultKeys = myHkvp.Keys
    myResultItems = myHkvp.Items
    
    'Assert.Strict:
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test06a_Exists()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, True, False, False)
    
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myHkvp.Exists("World")
    myResult(1) = myHkvp.Exists("Its")
    myResult(2) = myHkvp.Exists("Theree")
    myResult(3) = myHkvp.Exists(" Its")
    
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test06b_HoldsKey()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, True, False, False)
    
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myHkvp.HoldsKey("World")
    myResult(1) = myHkvp.HoldsKey("Its")
    myResult(2) = myHkvp.HoldsKey("Theree")
    myResult(3) = myHkvp.HoldsKey(" Its")
    
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test06c_LacksKey()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(False, False, True, True)
    
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myHkvp.LacksKey("World")
    myResult(1) = myHkvp.LacksKey("Its")
    myResult(2) = myHkvp.LacksKey("There")
    myResult(3) = myHkvp.LacksKey(" Its")
    
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
    
    
'@TestMethod("Hkvp")
Private Sub Test06d_HoldsItem()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, True, False, False)
    
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myHkvp.HoldsItem(10)
    myResult(1) = myHkvp.HoldsItem(50)
    myResult(2) = myHkvp.HoldsItem(42)
    myResult(3) = myHkvp.HoldsItem(-1)
    
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test06e_LacksItem()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(False, False, True, True)
    
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myHkvp.LacksItem(10)
    myResult(1) = myHkvp.LacksItem(50)
    myResult(2) = myHkvp.LacksItem(42)
    myResult(3) = myHkvp.LacksItem(-1)
    
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test07_IndexByKey()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(1, 2, -1, -1)
    
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myHkvp.IndexByKey("World")
    myResult(1) = myHkvp.IndexByKey("Its")
    myResult(2) = myHkvp.IndexByKey("Theree")
    myResult(3) = myHkvp.IndexByKey(" Its")
    
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test08_GetItemByIndex()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(10, 30, 50, Null)
    
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myHkvp.ItemByIndex(0)
    myResult(1) = myHkvp.ItemByIndex(2)
    myResult(2) = myHkvp.ItemByIndex(4)
    ' currently cHashD errors when out of range
   ' myResult(3) = myHkvp.ItemByIndex(7)
    
    'Assert.Strict:
    'This format is required as the VBA spec states that Null is not equal to Null
    ' so we cannot use sequence comparing
    AssertStrictAreEqual myExpected(0), myResult(0), myProcedureName
    AssertStrictAreEqual myExpected(1), myResult(1), myProcedureName
    AssertStrictAreEqual myExpected(2), myResult(2), myProcedureName
   ' AssertStrictAreEqual IsNull(myExpected(3)), IsNull(myResult(3)), myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test09_LetItemByIndex()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 6
    Dim myExpectedItems As Variant = Array(10, 20, 30, 42, 50, 60)
    Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day", " ")
    
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    'Act:
    myHkvp.ItemByIndex(3) = 42
    Dim myResultCount As Long = myHkvp.Count
    Dim myResultKeys As Variant = myHkvp.Keys
    Dim myResultItems As Variant = myHkvp.Items
    
    'Assert.Strict:
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test10_SetItemByIndex()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 6
    Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day", " ")
    
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myCollection As Collection = New Collection
    
    With myCollection
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
    End With
    
    'Act:
    Set myHkvp.ItemByIndex(3) = myCollection
    Dim myResultCount As Long = myHkvp.Count
    Dim myResultKeys As Variant = myHkvp.Keys
    
    
    'Assert.Strict:
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    'ToDO: revise class so that Item is not needed
    AssertStrictAreEqual 40, myHkvp.Item("A")(4), myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub Test11_GetItem()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
       'On Error GoTo TestFail
    
        'Arrange:
        Dim myExpected As Long = 50
       
        Dim myHkvp As Hkvp = Hkvp.Deb
        
        With myHkvp
            .Add "Hello", 10
            .Add "World", 20
            .Add "Its", 30
            .Add "A", 40
            .Add "Nice", 50
            .Add "Day", 60
        End With
        
        
        'Act:
        Dim myResult As Long = myHkvp.Item("Nice")
       
        'Assert.Strict:
        'This format is required as the VBA spec states that Null is not equal to Null
        ' so we cannot use sequence comparing
        AssertStrictAreEqual myExpected, myResult, myProcedureName
       
        
        TestExit:
        Exit Sub
        TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

'@TestMethod("Hkvp")
Private Sub Test12_LetItem()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 6
    Dim myExpectedItems As Variant = Array(10, 20, 30, 42, 50, 60)
    Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day", " ")
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    
    'Act:
    myHkvp.Item("A") = 42
    Dim myResultCount As Long = myHkvp.Count
    Dim myResultKeys As Variant = myHkvp.Keys
    Dim myResultItems As Variant = myHkvp.Items
    'Assert.Strict:
    'This format is required as the VBA spec states that Null is not equal to Null
    ' so we cannot use sequence comparing
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test13_SetItem()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 6
    Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day", " ")
    Dim myHkvp As Hkvp = Hkvp.Deb
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myCollection As Collection = New Collection
    
    With myCollection
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
    End With
    'Act:
    Set myHkvp.Item("A") = myCollection
    Dim myResultCount As Long = myHkvp.Count
    Dim myResultKeys As Variant = myHkvp.Keys
   
    'Assert.Strict:
    'This format is required as the VBA spec states that Null is not equal to Null
    ' so we cannot use sequence comparing
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    AssertStrictAreEqual 40, myHkvp.Item("A")(4), myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test14_DuplicateKeys()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 8
    Dim myExpectedItems As Variant = Array(10, 20, 30, 100, 40, 50, 60, 1000)
    Dim myExpectedKeys As Variant = Split("Hello World Its Hello A Nice Day Hello", " ")
   
    ' This line fails
    Dim myHkvp As Hkvp = Hkvp.Deb
    myHkvp.ReInit(ensureuniquekeys:=False)
    
    ' These two lines work OK
    'Dim myHkvp As Hkvp = Hkvp.Deb
    'myHkvp.ReInit(ensureuniquekeys:=False)
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "Hello", 100
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
        .Add "Hello", 1000
    End With
    
    'Act:
    Dim myResultCount As Long = myHkvp.Count
    Dim myResultKeys As Variant = myHkvp.Keys
    Dim myResultItems As Variant = myHkvp.Items
   
    'Assert.Strict:
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Hkvp")
Private Sub Test15_UniqueKeys()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    ' with unique keys ensured adding a duplicate key causes an error
    ' which we deliberatelately ignore
    ' but the duplicate keys do not get added
    ' ***On error resume next must be enabled for this test to pass***
   On Error Resume Next

    'Arrange:
    Dim myExpectedCount As Long = 6
    Dim myExpectedItems As Variant = Array(10, 20, 30, 40, 50, 60)
    Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day", " ")
    Dim myHkvp As Hkvp = Hkvp.Deb
   
    myHkvp.ReInit EnsureUniqueKeys:=True
    
    With myHkvp
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "Hello", 100
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
        .Add "Hello", 1000
    End With
   'On Error GoTo TestFail
    'Act:
    Dim myResultCount As Long = myHkvp.Count
    Dim myResultKeys As Variant = myHkvp.Keys
    Dim myResultItems As Variant = myHkvp.Items
   
    'Assert.Strict:
    'This format is required as the VBA spec states that Null is not equal to Null
    ' so we cannot use sequence comparing
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'** Reversing a dictionary doesn't amke sense
' '@TestMethod("Hkvp")
' Private Sub Test16_Reverse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

  
'    'On Error Resume Next

'     'Arrange:
'     Dim myExpectedCount As Long = 6
'     Dim myExpectedItems As Variant = Array(60, 50, 40, 30, 20, 10)
'     Dim myExpectedKeys As Variant = Split("Day Nice A Its World Hello", " ")
'     Dim myHkvp As Hkvp = Hkvp.Deb
   
  
    
'     With myHkvp
'         .Add "Hello", 10
'         .Add "World", 20
'         .Add "Its", 30
'         .Add "A", 40
'         .Add "Nice", 50
'         .Add "Day", 60
        
'     End With
'    'On Error GoTo TestFail
'     'Act:
'     Dim myReversed As Hkvp = myHkvp.Reverse
'     Dim myResultCount As Long = myReversed.Count
'     Dim myResultKeys As Variant = myReversed.Keys
'     Dim myResultItems As Variant = myReversed.Items
   
'     'Assert.Strict:
'     'This format is required as the VBA spec states that Null is not equal to Null
'     ' so we cannot use sequence comparing
'     AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
'     AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
'     AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub