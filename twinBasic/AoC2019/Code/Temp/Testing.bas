Attribute VB_Name = "Testing"
Option Explicit
Option Private Module
'@ignoreModule
'@TestModule
'@Folder("Tests")

Private Assert As Object
'Private Fakes As Object
Private myComp                      As IntComputer
Private myProgram                   As KvpOD
Private myInput                     As KvpOD


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

'@TestInitialize
Private Sub TestInitialize()
    
    Set myComp = New IntComputer
    Set myProgram = New KvpOD
    
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set myComp = Nothing
    Set myProgram = Nothing
    Set myInput = Nothing
End Sub

''@TestMethod("IntComputer")
'Public Sub TestGetAmpOutputSingleMode_Phases43210()
'
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected As Long
'    myExpected = 43210
'    Dim myResult As Long
'
'    'Act:
'    myResult = _
'        GetAmplifierOutputInSinglePassMode _
'        ( _
'            MakeProgram(Array(3, 15, 3, 16, 1002, 16, 10, 16, 1, 16, 15, 15, 4, 15, 99, 0, 0)), _
'            "43210" _
'        ).GetLast.Value
'
'    'Assert:
'    Assert.AreEqual myExpected, myResult
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'
'End Sub
''@TestMethod("IntComputer")
'Public Sub TestGetAmpOutputSingleMode_Phases01234()
'
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected As Long
'    myExpected = 54321
'    Dim myResult As Long
'
'    'Act:
'    myResult = _
'        GetAmplifierOutputInSinglePassMode _
'        ( _
'            MakeProgram(Array(3, 23, 3, 24, 1002, 24, 10, 24, 1002, 23, -1, 23, 101, 5, 23, 23, 1, 24, 23, 23, 4, 23, 99, 0, 0)), _
'            "01234" _
'        ).GetLast.Value
'
'    Assert.AreEqual myExpected, myResult
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub
''@TestMethod("IntComputer")
'Public Sub TestGetAmpOutputSingleMode_Phases10432()
'
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected As Long
'    myExpected = 65210
'    Dim myResult As Long
'
'    'Act:
'    myResult = _
'        GetAmplifierOutputInSinglePassMode _
'        ( _
'            MakeProgram(Array(3, 31, 3, 32, 1002, 32, 10, 32, 1001, 31, -2, 31, 1007, 31, 0, 33, 1002, 33, 7, 33, 1, 33, 31, 31, 1, 32, 31, 31, 4, 31, 99, 0, 0, 0)), _
'            "10432" _
'        ).GetLast.Value
'    'Assert:
'    Assert.AreEqual myExpected, myResult
'
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub
''@TestMethod("IntComputer")
'Public Sub TestGetAmoutputRepeatMode_98765()
'
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected As Long
'    myExpected = 139629729
'    Dim myResult As Long
'
'    'Act:
'    myResult = _
'        GetAmplifierOutputInLoopMode _
'        ( _
'            "5,6,7,8,9", _
'            AmplifierConfiguration.looped, _
'            MakeProgram(Array(3, 26, 1001, 26, -4, 26, 3, 27, 1002, 27, 2, 27, 1, 27, 26, 27, 4, 27, 1001, 28, -1, 28, 1005, 28, 6, 99, 0, 0, 5)) _
'        )
'
'    'Assert:
'    Assert.AreEqual myExpected, myResult
'
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub
'
'
'
''@TestMethod("IntComputer")
'Public Sub TestGetAmoutputRepeatMode_97856()
'
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected As Long
'    myExpected = 18216
'    Dim myResult As Long
'
'    'Act:
'    myResult = _
'        GetAmplifierOutputInLoopMode _
'        ( _
'            "5,6,7,8,9", _
'            AmplifierConfiguration.looped, _
'            MakeProgram(Array(3, 52, 1001, 52, -5, 52, 3, 53, 1, 52, 56, 54, 1007, 54, 5, 55, 1005, 55, 26, 1001, 54, -5, 54, 1105, 1, 12, 1, 53, 54, 53, 1008, 54, 0, 55, 1001, 55, 1, 55, 2, 53, 55, 53, 4, 53, 1001, 56, -1, 56, 1005, 56, 6, 99, 0, 0, 0, 0, 10)) _
'        )
'
'    'Assert:
'    Assert.AreEqual myExpected, myResult
'
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'
'End Sub

    
'@TestMethod("IntComputer")
Private Sub Day02_Input_1_9_10_3_2_3_11_0_99_30_40_50_gives_3500_9_10_70_2_3_11_0_99_30_40_50()                      'TODO Rename test
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(1, 9, 10, 3, 2, 3, 11, 0, 99, 30, 40, 50))
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run
    'Assert:
    Assert.AreEqual "3500,9,10,70,2,3,11,0,99,30,40,50", myComp.Program.GetValuesAsString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

    
'@TestMethod("IntComputer")
Private Sub Day02_Input_1_0_0_0_99_gives_2_0_0_0_99()                      'TODO Rename test
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(1, 0, 0, 0, 99))
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run
    'Assert:
    Assert.AreEqual "2,0,0,0,99", myComp.Program.GetValuesAsString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IntComputer")
Private Sub Day02_Input_2_3_0_3_99_gives_2_3_0_6_99()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(2, 3, 0, 3, 99))
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run
    'Assert:
    Assert.AreEqual "2,3,0,6,99", myComp.Program.GetValuesAsString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IntComputer")
Private Sub Day02_Input_2_4_4_5_99_0_gives_2_4_4_5_99_9801()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(2, 4, 4, 5, 99, 0))
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run
    'Assert:
    Assert.AreEqual "2,4,4,5,99,9801", myComp.Program.GetValuesAsString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

  

'@TestMethod("IntComputer")
Private Sub Day02_Input_1_1_1_4_99_5_6_0_99_gives_30_1_1_4_2_5_6_0_99()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(1, 1, 1, 4, 99, 5, 6, 0, 99))
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run
    'Assert:
    Assert.AreEqual "30,1,1,4,2,5,6,0,99", myComp.Program.GetValuesAsString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub




'@TestMethod("IntComputer")
Private Sub Day05_OpEquals_PositionMode_8_equals_8_is_1()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 9, 8, 9, 10, 9, 4, 9, 99, -1, 8))
    Set myInput = MakeKvp(8)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 1^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("IntComputer")
Private Sub Day05_OpEquals_PositionMode_1_equals_8_is_0()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 9, 8, 9, 10, 9, 4, 9, 99, -1, 8))
    Set myInput = MakeKvp(1)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 0^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IntComputer")
Private Sub Day05_OpEquals_PositionMode_1001_equals_8_is_0()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 9, 8, 9, 10, 9, 4, 9, 99, -1, 8))
    Set myInput = MakeKvp(1001)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 0^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IntComputer")
Private Sub Day05_OpLessThan_PositionMode_0_lessthan_8_is_1()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 9, 7, 9, 10, 9, 4, 9, 99, -1, 8))
    Set myInput = MakeKvp(0)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 1^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IntComputer")
Private Sub Day05_OpLessThan_PositionMode_8_lessthan_8_is_0()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 9, 7, 9, 10, 9, 4, 9, 99, -1, 8))
    Set myInput = MakeKvp(8)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 0^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



''@TestMethod("IntComputer")
'Private Sub Day05_OpLessThan_PositionMode_8_lessthan_8_is_0()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Set myProgram = MakeLongLongVsLongLongKvp(Array(3, 9, 7, 9, 10, 9, 4, 9, 99, -1, 8))
'    Set myInput = MakeKvp(8)
'    Set myComp.Program = myProgram
'
'    'Act:
'    myComp.Run myInput
'    'Assert:
'    Assert.AreEqual 0^, myComp.GetOutput.GetFirst.Value
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub

'=====================


'@TestMethod("IntComputer")
Private Sub Day05_OpEquals_ImmediateMode_8_equals_8_is_1()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 3, 1108, -1, 8, 3, 4, 3, 99))
    Set myInput = MakeKvp(8)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 1^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("IntComputer")
Private Sub Day05_OpEquals_ImmediateMode_1_equals_8_is_0()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 3, 1108, -1, 8, 3, 4, 3, 99))
    Set myInput = MakeKvp(1)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 0^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IntComputer")
Private Sub Day05_OpEquals_ImmediateMode_1001_equals_8_is_0()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 3, 1108, -1, 8, 3, 4, 3, 99))
    Set myInput = MakeKvp(1001)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 0^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IntComputer")
Private Sub Day05_OpLessThan_ImmediateMode_0_lessthan_8_is_1()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 3, 1107, -1, 8, 3, 4, 3, 99))
    Set myInput = MakeKvp(0)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 1^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IntComputer")
Private Sub Day05_OpLessThan_ImmediateMode_8_lessthan_8_is_0()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 3, 1107, -1, 8, 3, 4, 3, 99))
    Set myInput = MakeKvp(8)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 0^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==========================

'@TestMethod("IntComputer")
Private Sub Day05_Jump_PositionMode_0_gives_0()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 12, 6, 12, 15, 1, 13, 14, 13, 4, 13, 99, -1, 0, 1, 9))
    Set myInput = MakeKvp(0)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 0^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("IntComputer")
Private Sub Day05_Jump_PositionMode_not_0_gives_1()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 12, 6, 12, 15, 1, 13, 14, 13, 4, 13, 99, -1, 0, 1, 9))
    Set myInput = MakeKvp(8)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 1^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'====================


'@TestMethod("IntComputer")
Private Sub Day05_Jump_ImmediateMode_0_gives_0()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 3, 1105, -1, 9, 1101, 0, 0, 12, 4, 12, 99, 1))
    Set myInput = MakeKvp(0)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 0^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("IntComputer")
Private Sub Day05_Jump_Immediate_not_0_gives_1()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 3, 1105, -1, 9, 1101, 0, 0, 12, 4, 12, 99, 1))
    Set myInput = MakeKvp(8)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 1^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'=======================
'@TestMethod("IntComputer")
Private Sub Day05_Jump_Immediate_8_gives_1000()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 21, 1008, 21, 8, 20, 1005, 20, 22, 107, 8, 21, 20, 1006, 20, 31, 1106, 0, 36, 98, 0, 0, 1002, 21, 125, 20, 4, 20, 1105, 1, 46, 104, 999, 1105, 1, 46, 1101, 1000, 1, 20, 4, 20, 1105, 1, 46, 98, 99))
    Set myInput = MakeKvp(8)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 1000^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("IntComputer")
Private Sub Day05_Jump_Immediate_lessthan_8_gives_999()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 21, 1008, 21, 8, 20, 1005, 20, 22, 107, 8, 21, 20, 1006, 20, 31, 1106, 0, 36, 98, 0, 0, 1002, 21, 125, 20, 4, 20, 1105, 1, 46, 104, 999, 1105, 1, 46, 1101, 1000, 1, 20, 4, 20, 1105, 1, 46, 98, 99))
    Set myInput = MakeKvp(0)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 999^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("IntComputer")
Private Sub Day05_Jump_Immediate_morethan_8_gives_1001()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(3, 21, 1008, 21, 8, 20, 1005, 20, 22, 107, 8, 21, 20, 1006, 20, 31, 1106, 0, 36, 98, 0, 0, 1002, 21, 125, 20, 4, 20, 1105, 1, 46, 104, 999, 1105, 1, 46, 1101, 1000, 1, 20, 4, 20, 1105, 1, 46, 98, 99))
    Set myInput = MakeKvp(99)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 1001^, myComp.GetOutput.GetFirst.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'========================


'@TestMethod("IntComputer")
Private Sub Day05_Relative_output_is_copy_of_program()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(109, 1, 204, -1, 1001, 100, 1, 100, 1008, 100, 16, 101, 1006, 101, 0, 99))
    Set myComp.Program = myProgram
    myComp.OutputMode = ContinueOnOutput
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual "109,1,204,-1,1001,100,1,100,1008,100,16,101,1006,101,0,99", myComp.Program.GetValuesAsString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("IntComputer")
Private Sub Day05_JRelative_output_is_16_digit_number()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(1102, 34915192, 34915192, 7, 4, 7, 99, 0))
    Set myInput = MakeKvp(0)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual CLng(16), CLng(Len(CStr(myComp.GetOutput.GetFirst.Value.Value)))


TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("IntComputer")
Private Sub Day05_Relative_output_is_1125899906842624()
    On Error GoTo TestFail

    'Arrange:
    Set myProgram = MakeProgram(Array(104, 1125899906842624^, 99))
    Set myInput = MakeKvp(99)
    Set myComp.Program = myProgram
    
    'Act:
    myComp.Run myInput
    'Assert:
    Assert.AreEqual 1125899906842624^, myComp.GetOutput.GetFirst.Value
'    Dim myVar As Variant
'    myVar = Array(1, "Hello", 5.4)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'    Debug.Print "Day 9 Updates"
'    'Test 8: Copy of itself
'    Set myProgram = MakeLongLongVsLongLongKvp(Array(109, 1, 204, -1, 1001, 100, 1, 100, 1008, 100, 16, 101, 1006, 101, 0, 99))
'    Set myComp.Program = myProgram
'    myComp.Run
'    Debug.Print "Test 8", IIf(InStr(myComp.GetOutput.GetValuesAsString, myProgram.GetValuesAsString) > 0, "Pass", "Fail")
'
'    'Test 9
'    Set myComp = New IntComputer
'    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(1102, 34915192, 34915192, 7, 4, 7, 99, 0))
'    myComp.Run
'    Debug.Print "Test 9", 16 - Len(CStr(myComp.GetOutput.GetLast.Value))
'
'    'Set myComp = New IntComputer
'    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(104, 112589990684262^, 99))
'    myComp.Run
'    Debug.Print "Test 10", 112589990684262^ - CLngLng(myComp.GetOutput.GetLast.Value)
'
'End Sub







