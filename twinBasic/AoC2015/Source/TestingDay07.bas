Attribute VB_Name = "TestingDay07"
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


'@TestMethod("Day07")
Private Sub Test01_TrialData()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(123&, 456&, 72&, 507&, 492&, 114&, 65412, 65079)
    Dim myCircuit As Kvp
    Set myCircuit = New Kvp
    myCircuit.AddByIndexFromArray Split("123 -> x,456 -> y,x AND y -> d,x OR y -> e,x LSHIFT 2 -> f,y RSHIFT 2 -> g,NOT x -> h,NOT y -> i", ",")
    
    Dim myKvpByWire As Kvp
    Set myKvpByWire = Day07.GetOutputWireVsLogicArray(myCircuit)
    Dim myOperatedCircuit As Kvp
    Set myOperatedCircuit = Day07.OperateCircuit(myKvpByWire)
    Dim myResult As Variant
    
    'Act:
    myResult = myOperatedCircuit.GetValues
    
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

