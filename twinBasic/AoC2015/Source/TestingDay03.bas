Attribute VB_Name = "TestingDay03"
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


'@TestMethod("Day03")
Private Sub Test01_VisitsIs2()                        'TODO Rename test

    On Error GoTo TestFail
    
    'Arrange:
    Dim myInstructions As Kvp: Set myInstructions = New Kvp
    myInstructions.AddByIndexAsCharacters ">"
    
    Dim myJourney As Journey
    Set myJourney = Journey.Make(myInstructions)
    
    Dim myExpected As Long
    myExpected = 2
    
    Dim myResult As Long
    'Act:

    myResult = myJourney.GetVisits.Count
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Day03")
Private Sub Test02_VisitsIs4()                        'TODO Rename test

    On Error GoTo TestFail
    
    'Arrange:
    Dim myInstructions As Kvp: Set myInstructions = New Kvp
    myInstructions.AddByIndexAsCharacters "^>v<"
    
    Dim myJourney As Journey
    Set myJourney = Journey.Make(myInstructions)
    
    Dim myExpected As Long
    myExpected = 4
    
    Dim myResult As Long
    'Act:

    myResult = myJourney.GetVisits.Count
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Day03")
Private Sub Test03_VisitsIs2()                        'TODO Rename test

    On Error GoTo TestFail
    
    'Arrange:
    Dim myInstructions As Kvp: Set myInstructions = New Kvp
    myInstructions.AddByIndexAsCharacters "^v^v^v^v^v"
    
    Dim myJourney As Journey
    Set myJourney = Journey.Make(myInstructions)
    
    Dim myExpected As Long
    myExpected = 2
    
    Dim myResult As Long
    'Act:

    myResult = myJourney.GetVisits.Count
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
