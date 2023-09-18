Attribute VB_Name = "TestKvp"
Option Explicit
'@IgnoreModule
'@TestModule
'@Folder("VBASupport")


Private Assert                                  As Rubberduck.AssertClass
'Private Fakes                                   As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    'Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

''@TestInitialize
'Private Sub TestInitialize()
''this method runs before every test in the module.
'End Sub
'
'
''@TestCleanup
'Private Sub TestCleanup()
''this method runs after every test in the module.
'End Sub


'@TestMethod("Kvp")
Private Sub IsObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    
    'Act:
    Set myKvp = New KvpOD
    
    'Assert:
    Assert.AreEqual "Kvp", TypeName(myKvp)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub IsEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    
    'Act:
    Set myKvp = New KvpOD
    
    'Assert:
    Assert.AreEqual True, myKvp.IsEmpty

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub IsNotEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    
    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByIndex "Hello World"
    'Assert:
    Assert.AreEqual True, myKvp.IsNotEmpty

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Count()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    
    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByIndex "Hello World"
    'Assert:
    Assert.AreEqual 1&, myKvp.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Add_ByIndex()
    On Error GoTo TestFail

    'Arrange:
    Dim myKvp                              As KvpOD

    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByIndex "Hello World"
    'Assert:

    Assert.AreEqual "Hello World", myKvp.Item(0&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Add_ByIndexFromArray()
    On Error GoTo TestFail

    'Arrange:
    Dim myKvp                                      As KvpOD
    Dim myArray(3)                                 As Variant
    Dim myResult_Array()                           As Variant
    'Act:
    myArray(0) = "Hello"
    myArray(1) = True
    myArray(2) = 42
    myArray(3) = 3.142
    
    Set myKvp = New KvpOD
    myKvp.AddByIndexFromSeq myArray
    myResult_Array = myKvp.GetValues
    'Assert:

    Assert.SequenceEquals myArray, myResult_Array

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Add_ByKey()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    
    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByKey Key:=22&, Value:="Hello World"
    'Assert:
    'Debug.Print myKvp.Item(CLng(1))
    Assert.AreEqual "Hello World", myKvp.Item(22&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Add_ByKey_string_keys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    
    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByKey Key:="one", Value:="Hello World one"
    myKvp.AddByKey Key:="two", Value:="hellow WOrld two"
    'Assert:
    'Debug.Print myKvp.Item(CLng(1))
    Assert.AreEqual "Hello World one", myKvp.Item("one")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub HoldsKey()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    
    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByKey Key:=22&, Value:="Hello World"
    'Assert:
    'Debug.Print myKvp.Item(CLng(1))
    Assert.AreEqual True, myKvp.HoldsKey(22&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub HoldsValue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    
    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByKey Key:=22&, Value:="Hello World"
    'Assert:
    'Debug.Print myKvp.Item(CLng(1))
    Assert.AreEqual True, myKvp.HoldsValue("Hello World")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub LacksValue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    
    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByKey Key:=22&, Value:="Hello World"
    'Assert:
    'Debug.Print myKvp.Item(CLng(1))
    Assert.AreEqual True, myKvp.LacksValue(22&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub LacksKey()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    
    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByKey Key:=22, Value:="Hello World"
    'Assert:
    'Debug.Print myKvp.Item(CLng(1))
    Assert.AreEqual True, myKvp.LacksKey(80)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Remove()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    
    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByKey Key:=22&, Value:="Hello World"
    'Assert:
    'Debug.Print myKvp.Item(CLng(1))
    If myKvp.LacksKey(22&) Then GoTo TestFail
    myKvp.Remove 22&
    Assert.AreEqual True, myKvp.LacksKey(22&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub RemoveAll()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    
    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=25&, Value:="Hello World 3"
    'Assert:
    'Debug.Print myKvp.Item(CLng(1))
    
    myKvp.RemoveAll
    Assert.AreEqual True, myKvp.IsEmpty

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub MakeFromString()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                              As KvpOD
    Dim myKvp2                             As KvpOD
    Dim myResult1()                    As Variant
    Dim myResult2()                    As Variant
    'Act:
    Set myKvp1 = New KvpOD
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    
    Set myKvp2 = New KvpOD
    myResult1 = myKvp1.GetValues
    
    myKvp2.AddByIndexAsSubStr "Hello World 1,Hello World 2,Hello World 3"
    myResult2 = myKvp2.GetValues
    'Assert:
    
    Assert.SequenceEquals myResult1, myResult2

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub GetKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                               As KvpOD
    Dim myKvp_keys(2)                       As Long
    Dim myResult_keys                       As Variant

    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=25&, Value:="Hello World 3"
    
    myKvp_keys(0) = 22&
    myKvp_keys(1) = 23&
    myKvp_keys(2) = 25&
    
    myResult_keys = myKvp.GetKeys
    'Assert:
    Assert.SequenceEquals myKvp_keys, myResult_keys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub GetValues()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As KvpOD
    ' Dynamicops integers rather than long
    Dim myItems(2)                          As String
    Dim myKvp_items()                       As Variant
    'Act:
    Set myKvp = New KvpOD
    myKvp.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp.AddByKey Key:=25&, Value:="Hello World 3"
    
    myItems(0) = "Hello World 1"
    myItems(1) = "Hello World 2"
    myItems(2) = "Hello World 3"
    myKvp_items = myKvp.GetValues
    'Assert:
    Assert.SequenceEquals myItems, myKvp_items

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Cohorts_all_A_with_B_not_input_A()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As KvpOD
    Dim myKvp2                                 As KvpOD
    Dim myResult_keys(7)                       As Long
    Dim myResult                               As KvpOD
    
    
    Dim myCohortKeys()                          As Variant
    'Act:
    Set myKvp1 = New KvpOD
    myKvp1.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=3&, Value:="Hello World 3a"
    myKvp1.AddByKey Key:=4&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=5&, Value:="Hello World 5"
    myKvp1.AddByKey Key:=6&, Value:="Hello World 6"
    
    Set myKvp2 = New KvpOD
    myKvp2.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp2.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp2.AddByKey Key:=3&, Value:="Hello World 3b"
    myKvp2.AddByKey Key:=6&, Value:="Hello World 6"
    myKvp2.AddByKey Key:=7&, Value:="Hello World 7"
    myKvp2.AddByKey Key:=8&, Value:="Hello World 8"
    
    myResult_keys(0) = 1&
    myResult_keys(1) = 2&
    myResult_keys(2) = 3&
    myResult_keys(3) = 4&
    myResult_keys(4) = 5&
    myResult_keys(5) = 6&
    myResult_keys(6) = 7&
    myResult_keys(7) = 8&
    
    'Debug.Print myKvp1(1&)
    'Debug.Print myKvp1(2&)
    ' Debug.Print myKvp1(3&)
    ' Debug.Print myKvp1(4&)
    ' Debug.Print myKvp1(5&)
    ' Debug.Print myKvp1(6&)

    Set myResult = myKvp1.Cohorts(myKvp2)
    myCohortKeys = myResult.Item(1&).GetKeys
    '    myCohortKeys = myResult(1&).GetKeys
    '    myCohortKeys = myResult(2&).GetKeys
    '    myCohortKeys = myResult(3&).GetKeys
    '    myCohortKeys = myResult(4&).GetKeys
    '    myCohortKeys = myResult(5&).GetKeys
    'Assert:
    Assert.SequenceEquals myResult_keys, myCohortKeys

    Set myKvp1 = Nothing
    Set myKvp2 = Nothing
    Set myResult.Item(0) = Nothing
    Set myResult.Item(1) = Nothing
    Set myResult.Item(2) = Nothing
    Set myResult.Item(3) = Nothing
    Set myResult.Item(4) = Nothing
    Set myResult.Item(5) = Nothing
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Cohorts_B_input_A_with_different_value()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As KvpOD
    Dim myKvp2                                 As KvpOD
    Dim myResult_keys(0)                       As Long
    Dim myResult                             As KvpOD
    
    Dim myCohortKeys()                        As Variant
    'Act:
    Set myKvp1 = New KvpOD
    myKvp1.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=3&, Value:="Hello World 3a"
    myKvp1.AddByKey Key:=4&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=5&, Value:="Hello World 5"
    myKvp1.AddByKey Key:=6&, Value:="Hello World 6"
    
    Set myKvp2 = New KvpOD
    myKvp2.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp2.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp2.AddByKey Key:=3&, Value:="Hello World 3b"
    myKvp2.AddByKey Key:=6&, Value:="Hello World 6"
    myKvp2.AddByKey Key:=7&, Value:="Hello World 7"
    myKvp2.AddByKey Key:=8&, Value:="Hello World 8"
    
    myResult_keys(0) = 3&
    
    Set myResult = myKvp1.Cohorts(myKvp2)
    myCohortKeys = myResult.Item(2&).GetKeys
    'Assert:
    Assert.SequenceEquals myResult_keys, myCohortKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Cohorts_input_A_only_and_input_B_only()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As KvpOD
    Dim myKvp2                                 As KvpOD
    Dim myResult_keys(3)                       As Long
    Dim myResult                             As KvpOD
    
    Dim myCohortKeys()                        As Variant
    'Act:
    Set myKvp1 = New KvpOD
    myKvp1.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=3&, Value:="Hello World 3a"
    myKvp1.AddByKey Key:=4&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=5&, Value:="Hello World 5"
    myKvp1.AddByKey Key:=6&, Value:="Hello World 6"
    
    Set myKvp2 = New KvpOD
    myKvp2.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp2.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp2.AddByKey Key:=3&, Value:="Hello World 3b"
    myKvp2.AddByKey Key:=6&, Value:="Hello World 6"
    myKvp2.AddByKey Key:=7&, Value:="Hello World 7"
    myKvp2.AddByKey Key:=8&, Value:="Hello World 8"
    
    
    myResult_keys(0) = 4&
    myResult_keys(1) = 5&
    myResult_keys(2) = 7&
    myResult_keys(3) = 8&
    
    Set myResult = myKvp1.Cohorts(myKvp2)
    myCohortKeys = myResult.Item(3&).GetKeys
    
    'Assert:
    Assert.SequenceEquals myResult_keys, myCohortKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Cohorts_input_both_A_and_B()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As KvpOD
    Dim myKvp2                                 As KvpOD
    Dim myResult_keys(3)                       As Long
    Dim myResult                             As KvpOD
    
    Dim myCohortKeys()                        As Variant
    'Act:
    Set myKvp1 = New KvpOD
    myKvp1.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=3&, Value:="Hello World 3a"
    myKvp1.AddByKey Key:=4&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=5&, Value:="Hello World 5"
    myKvp1.AddByKey Key:=6&, Value:="Hello World 6"
    
    Set myKvp2 = New KvpOD
    myKvp2.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp2.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp2.AddByKey Key:=3&, Value:="Hello World 3b"
    myKvp2.AddByKey Key:=6&, Value:="Hello World 6"
    myKvp2.AddByKey Key:=7&, Value:="Hello World 7"
    myKvp2.AddByKey Key:=8&, Value:="Hello World 8"
    
    myResult_keys(0) = 1&
    myResult_keys(1) = 2&
    myResult_keys(2) = 3&
    myResult_keys(3) = 6&
    
    Set myResult = myKvp1.Cohorts(myKvp2)
    myCohortKeys = myResult.Item(4&).GetKeys
    
    'Assert:
    Assert.SequenceEquals myResult_keys, myCohortKeys


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Cohorts_input_A_only()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As KvpOD
    Dim myKvp2                                 As KvpOD
    Dim myResult_keys(1)                  As Long
    Dim myResult                             As KvpOD
    
    Dim myCohortKeys()                        As Variant
    'Act:
    Set myKvp1 = New KvpOD
    myKvp1.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=3&, Value:="Hello World 3a"
    myKvp1.AddByKey Key:=4&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=5&, Value:="Hello World 5"
    myKvp1.AddByKey Key:=6&, Value:="Hello World 6"
    
    Set myKvp2 = New KvpOD
    myKvp2.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp2.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp2.AddByKey Key:=3&, Value:="Hello World 3b"
    myKvp2.AddByKey Key:=6&, Value:="Hello World 6"
    myKvp2.AddByKey Key:=7&, Value:="Hello World 7"
    myKvp2.AddByKey Key:=8&, Value:="Hello World 8"
    
    myResult_keys(0) = 4&
    myResult_keys(1) = 5&
  
    Set myResult = myKvp1.Cohorts(myKvp2)
    myCohortKeys = myResult.Item(5&).GetKeys
    
    'Assert:
    Assert.SequenceEquals myResult_keys, myCohortKeys

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub Cohorts_input_B_only()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As KvpOD
    Dim myKvp2                                 As KvpOD
    Dim myResult_keys(1)                       As Long
    Dim myResult                             As KvpOD
    
    Dim myCohortKeys()                        As Variant
    'Act:
    Set myKvp1 = New KvpOD
    myKvp1.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=3&, Value:="Hello World 3a"
    myKvp1.AddByKey Key:=4&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=5&, Value:="Hello World 5"
    myKvp1.AddByKey Key:=6&, Value:="Hello World 6"
    
    Set myKvp2 = New KvpOD
    myKvp2.AddByKey Key:=1&, Value:="Hello World 1"
    myKvp2.AddByKey Key:=2&, Value:="Hello World 2"
    myKvp2.AddByKey Key:=3&, Value:="Hello World 3b"
    myKvp2.AddByKey Key:=6&, Value:="Hello World 6"
    myKvp2.AddByKey Key:=7&, Value:="Hello World 7"
    myKvp2.AddByKey Key:=8&, Value:="Hello World 8"
    
    myResult_keys(0) = 7&
    myResult_keys(1) = 8&
  
    Set myResult = myKvp1.Cohorts(myKvp2)
    myCohortKeys = myResult.Item(6&).GetKeys
    
    'Assert:
    Assert.SequenceEquals myResult_keys, myCohortKeys


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")

Private Sub Mirror()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                     As KvpOD
    Dim myKvp2                                     As KvpOD
    'Act:
    Set myKvp1 = New KvpOD
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    Set myKvp2 = myKvp1.Mirror
    
    'Assert:
    Assert.SequenceEquals myKvp1.GetKeys, myKvp2.Item(1&).GetValues

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub ItemsAreUnique()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As KvpOD
    
    'Act:
    Set myKvp1 = New KvpOD
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    
    
    'Assert:
    Assert.AreEqual True, myKvp1.ValuesAreUnique

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub PullFirst()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As KvpOD
    
    'Act:
    Set myKvp1 = New KvpOD
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    Dim myResult As String
    myResult = myKvp1.PullFirst
    
    
    'Assert:
    Assert.IsTrue myKvp1.LacksKey(22^)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Kvp")
Private Sub PullLast()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As KvpOD
    
    'Act:
    Set myKvp1 = New KvpOD
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    
    Dim myResult As String
    myResult = myKvp1.PullLast
    
    
    'Assert:
    Assert.IsTrue myKvp1.LacksKey(27^)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Kvp")
Private Sub PullAny()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp1                                 As KvpOD
    
    'Act:
    Set myKvp1 = New KvpOD
    myKvp1.AddByKey Key:=22&, Value:="Hello World 1"
    myKvp1.AddByKey Key:=23&, Value:="Hello World 2"
    myKvp1.AddByKey Key:=25&, Value:="Hello World 3"
    myKvp1.AddByKey Key:=26&, Value:="Hello World 4"
    myKvp1.AddByKey Key:=27&, Value:="Hello World 5"
    Debug.Print myKvp1.GetKeysAsString
    Dim myResult As String
    myResult = myKvp1.Pull(25&)
    Debug.Print myKvp1.GetKeysAsString
    
    'Assert:
    Assert.IsTrue myKvp1.LacksKey(25&)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

