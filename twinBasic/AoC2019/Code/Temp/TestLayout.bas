Attribute VB_Name = "TestLayout"
Option Explicit
Option Private Module
'@TestModule
'@Folder("Layout")

Private Assert                                  As Rubberduck.AssertClass
'Private Fakes                                   As Rubberduck.FakesProvider

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    'Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
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


'@TestMethod("Layout")
Public Sub ConvertFormatFieldXX0ToVbNullString_01()

    Dim myExpected                      As String
    Dim myTest                          As String
    Dim myResult                        As String

    On Error GoTo TestFail
    
    myExpected = "Hello    Hello"
    myTest = "Hello {nl0} {tb0} {nt0} Hello"
    myResult = ConvertFormatFieldXX0ToVbNullString(myTest)
    Assert.AreEqual myExpected, myResult
    ''TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Layout")
Public Sub ConvertFormatFieldXXToXX1_01()

    Dim myExpected                      As String
    Dim myTest                          As String
    Dim myResult                        As String

    On Error GoTo TestFail
    
    'Arrange
    myExpected = "Hello {nl1} {tb1} {nt1} Hello"
    myTest = "Hello {nl} {tb} {nt} Hello"
    
    'Act
    myResult = ConvertFormatFieldXXToXX1(myTest)
    
    'Assert
    Assert.AreEqual myExpected, myResult
    
    'TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Layout")
Public Sub GetFieldCount_01_nl3_equals_3()

    Dim myExpected                     As Long
    Dim myResult                         As Long

    On Error GoTo TestFail
    
    'Arrange
    myExpected = 3
    
    'Act
    myResult = GetFieldRepeatCount("Hello {nl3} Hello", "{nl")
    
    'Assert
    Assert.AreEqual myExpected, myResult
    
    'TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Layout")
Public Sub GetFormattingReplaceString_01_nl3()

    Dim myExpected                     As String
    Dim myTest                         As String

    On Error GoTo TestFail
    
    myExpected = String$(3, vbCrLf)              'vbCrLf & vbCrLf & vbCrLf
    myTest = GetFormattingReplaceString("{nl", 3) ' nl is the code for vbcrlf
    
    Debug.Print Len(myExpected), Len(myTest), Len(vbCrLf & vbCrLf & vbCrLf), Len(String$(3, vbCrLf))
    Assert.AreEqual myExpected, myTest
    'TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Layout")
Public Sub GetFormattingReplaceString_02_tb4()

    Dim myExpected                     As String
    Dim myTest                         As String

    On Error GoTo TestFail
    
    myExpected = vbTab & vbTab & vbTab & vbTab
    myTest = GetFormattingReplaceString("{tb", 4)
    Debug.Print "02", Len(myExpected), Len(myTest), Len(vbTab & vbTab & vbTab)
    Assert.AreEqual myExpected, myTest
    'TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Layout")
Public Sub GetFormattingReplaceString_03_nt2()

    Dim myExpected                     As String
    Dim myTest                         As String

    On Error GoTo TestFail
    
    'Arrange
    myExpected = String$(2, vbCrLf) & vbTab      'vbCrLf & vbCrLf & vbTab
    
    'Act
    myTest = GetFormattingReplaceString("{nt", 2)
    Debug.Print "03", Len(myExpected), Len(myTest), Len(vbCrLf & vbCrLf & vbTab), Len(String$(2, vbCrLf) & vbTab)
    'Assert
    Assert.AreEqual myExpected, myTest
    
    'TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Layout")
Public Sub ConvertVariableFieldsItemingRepresentations_01_FourArgs()

    Dim myArray                                 As Variant
    Dim myExpected                              As String
    Dim myTest                                  As String
    Dim myResult                                As String

    On Error GoTo TestFail

    'Arrange:
    myArray = Array(1, "Hello World", True, 5.134)
    myExpected = "Integer 1: 1, Text: Hello World, Boolean: True, Double: 5.134"
    myTest = "Integer 1: {0}, Text: {1}, Boolean: {2}, Double: {3}"
    
    'Act
    myResult = ConvertVariableFieldsItemingRepresentations(myTest, myArray)

    'Assert:
    Assert.AreEqual myExpected, myResult

    'TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Layout")
Public Sub ConvertVariableFieldsItemingRepresentations_01_FiveArgs_withobject()

    Dim myArray                                 As Variant
    Dim myExpected                              As String
    Dim myResult                                As String
    Dim myTest                                  As String
    On Error GoTo TestFail

    'Arrange:
    myArray = Array(1, "Hello World", True, 5.134, New KvpOD)
    myExpected = "Integer 1: 1, Text: Hello World, Boolean: True, Double: 5.134, Object: {Can't stringify Type: Kvp}"
    myTest = "Integer 1: {0}, Text: {1}, Boolean: {2}, Double: {3}, Object: {4}"
    
    'Act
    myResult = ConvertVariableFieldsItemingRepresentations(myTest, myArray)

    'Assert:
    Assert.AreEqual myExpected, myResult

    'TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

