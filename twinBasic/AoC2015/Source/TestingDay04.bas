Attribute VB_Name = "TestingDay04"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
'Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    'Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub


'@TestMethod("Day04")
Private Sub Test01_abcdef609043()                        'TODO Rename test

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    
    'Act:
    myResult = VBA.Left$(Day04.StringToMD5Hex("abcdef609043"), 5) = "00000"
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Day04")
Private Sub Test01_pqrstuv1048970()                        'TODO Rename test

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    
    'Act:
    myResult = VBA.Left$(Day04.StringToMD5Hex("pqrstuv1048970"), 5) = "00000"
    
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
