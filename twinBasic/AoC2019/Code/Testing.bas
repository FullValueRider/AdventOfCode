Attribute VB_Name = "Testing"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule
Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


Public Sub TestGetAmpOutputSingleMode()

    Debug.Print "Amplifier Tests"
    Debug.Print "Single Pass mode", CLng("43210") - CLng(GetAmpOutput(MakeLongLongVsLongLongKvp(Array(3, 15, 3, 16, 1002, 16, 10, 16, 1, 16, 15, 15, 4, 15, 99, 0, 0)), "43210", AmplifierModeEnum.SinglePass).GetLast)
    Debug.Print "Single Pass mode", CLng("54321") - CLng(GetAmpOutput(MakeLongLongVsLongLongKvp(Array(3, 23, 3, 24, 1002, 24, 10, 24, 1002, 23, -1, 23, 101, 5, 23, 23, 1, 24, 23, 23, 4, 23, 99, 0, 0)), "01234", AmplifierModeEnum.SinglePass).GetLast)
    Debug.Print "Single Pass mode", CLng("65210") - CLng(GetAmpOutput(MakeLongLongVsLongLongKvp(Array(3, 31, 3, 32, 1002, 32, 10, 32, 1001, 31, -2, 31, 1007, 31, 0, 33, 1002, 33, 7, 33, 1, 33, 31, 31, 1, 32, 31, 31, 4, 31, 99, 0, 0, 0)), "10432", AmplifierModeEnum.SinglePass).GetLast)
    
End Sub

Public Sub TestGetAmoutputRepeatMode()
    
    Debug.Print "Amplifier Tests"
    Dim myProgram As Kvp
    Set myProgram = MakeLongLongVsLongLongKvp(Array(3, 26, 1001, 26, -4, 26, 3, 27, 1002, 27, 2, 27, 1, 27, 26, 27, 4, 27, 1001, 28, -1, 28, 1005, 28, 6, 99, 0, 0, 5))
    
    'ByVal ipAmplifierPhaseOptions As String, ByVal ipRunMode As AmplifierModeEnum, ByVal ipProgram As Variant
    Debug.Print "Loop mode", CLngLng("139629729") - CLng(GetMaxAmplifierOutput("5,6,7,8,9", AmplifierModeEnum.MultiPass, myProgram))
    
    Set myProgram = MakeLongLongVsLongLongKvp(Array(3, 52, 1001, 52, -5, 52, 3, 53, 1, 52, 56, 54, 1007, 54, 5, 55, 1005, 55, 26, 1001, 54, -5, 54, 1105, 1, 12, 1, 53, 54, 53, 1008, 54, 0, 55, 1001, 55, 1, 55, 2, 53, 55, 53, 4, 53, 1001, 56, -1, 56, 1005, 56, 6, 99, 0, 0, 0, 0, 10))
    Debug.Print "Loop Mode", CLng("18216") - CLng(GetMaxAmplifierOutput("5,6,7,8,9", AmplifierModeEnum.MultiPass, myProgram))

End Sub


'@Ignore ProcedureNotUsed
Private Sub TestCOmputerDay5Updates()
 
    Debug.Print "Computer Tests:"
    Debug.Print
    Debug.Print "Day2 Tests"
    
    Dim myComp As IntComputer:    Set myComp = New IntComputer
    Dim myProgram As Kvp: Set myProgram = New Kvp
    
    
    Set myProgram = MakeLongLongVsLongLongKvp(Array(1, 9, 10, 3, 2, 3, 11, 0, 99, 30, 40, 50))
    Set myComp.Program = myProgram
    myComp.Run
    Debug.Print "Day 2 Test 1: ", IIf(InStr(myComp.Program.GetValuesAsString, "3500,9,10,70,2,3,11,0,99,30,40,50") = 1, 0, 1)
    
   
    Set myProgram = MakeLongLongVsLongLongKvp(Array(1, 0, 0, 0, 99))
    Set myComp.Program = myProgram
    myComp.Run
    Debug.Print "Day 2 Test 2: ", IIf(InStr(myComp.Program.GetValuesAsString, "2,0,0,0,99") = 1, 0, 1)

    
    Set myProgram = MakeLongLongVsLongLongKvp(Array(2, 3, 0, 3, 99))
    Set myComp.Program = myProgram
    myComp.Run
    Debug.Print "Day 2 Test 3: ", IIf(InStr(myComp.Program.GetValuesAsString, "2,3,0,6,99") = 1, 0, 1)

  
    Set myProgram = MakeLongLongVsLongLongKvp(Array(2, 4, 4, 5, 99, 0))
    Set myComp.Program = myProgram
    myComp.Run
    Debug.Print "Day 2 Test 4: ", IIf(InStr(myComp.Program.GetValuesAsString, "2,4,4,5,99,9801") = 1, 0, 1)

    Set myProgram = MakeLongLongVsLongLongKvp(Array(1, 1, 1, 4, 99, 5, 6, 0, 99))
    Set myComp.Program = myProgram
    myComp.Run
    Debug.Print "Day 2 Test 5: ", IIf(InStr(myComp.Program.GetValuesAsString, "30,1,1,4,2,5,6,0,99") = 1, 0, 1)
    Debug.Print

    Debug.Print "Day 5 update"
'    Dim myComp As IntComputer: Set myComp = New IntComputer
'    Dim myProgram As Kvp
    Dim myInput As Kvp
     
    ' Test 1:Using position modeEqual to 8: input 8
    Set myProgram = MakeLongLongVsLongLongKvp(Array(3, 9, 8, 9, 10, 9, 4, 9, 99, -1, 8))
    Set myInput = MakeKvp(8)
    Set myComp.Program = myProgram
    myComp.Run myInput
    Debug.Print "Test 1A", 1 - CLng(myComp.GetOutput.GetLast)
       
    'Test 1:equal to 8:Using position mode input less than
    Set myProgram = MakeLongLongVsLongLongKvp(Array(3, 9, 8, 9, 10, 9, 4, 9, 99, -1, 8))
    Set myInput = MakeKvp(-1)
    Set myComp.Program = myProgram
    myComp.Run myInput
    Debug.Print "Test 1B", 0 - CLng(myComp.GetOutput.GetLast)
    
    'Test 1:equal to 8:Using position mode input greater than
    Set myProgram = MakeLongLongVsLongLongKvp(Array(3, 9, 8, 9, 10, 9, 4, 9, 99, -1, 8))
    Set myInput = MakeKvp(1001)
    Set myComp.Program = myProgram
    myComp.Run myInput
    Debug.Print "Test 1C", 0 - CLng(myComp.GetOutput.GetLast)
    Debug.Print
    
    ' Test 2:Using position mode:less than 8 input less than
    Set myProgram = MakeLongLongVsLongLongKvp(Array(3, 9, 7, 9, 10, 9, 4, 9, 99, -1, 8))
    Set myInput = MakeKvp(0)
    Set myComp.Program = myProgram
    myComp.Run myInput
    Debug.Print "Test 2A", 1 - CLng(myComp.GetOutput.GetLast)
    
    ' Test 2:Using position mode:less than 8: input 8
    Set myProgram = MakeLongLongVsLongLongKvp(Array(3, 9, 7, 9, 10, 9, 4, 9, 99, -1, 8))
    Set myInput = MakeKvp(8)
    Set myComp.Program = myProgram
    myComp.Run myInput
    Debug.Print "Test 2B", 0 - CLng(myComp.GetOutput.GetLast)
    Debug.Print
    
   ' Test 3: Using immediate mode:Equal to 8: input 8
    Set myProgram = MakeLongLongVsLongLongKvp(Array(3, 3, 1108, -1, 8, 3, 4, 3, 99))
    Set myInput = MakeKvp(8)
    Set myComp.Program = myProgram
    myComp.Run myInput
    Debug.Print "Test 3A", 1 - CLng(myComp.GetOutput.GetLast)
       
    'Test 3: Using immediate mode:equal to 8: input less than
    Set myProgram = MakeLongLongVsLongLongKvp(Array(3, 3, 1108, -1, 8, 3, 4, 3, 99))
    Set myInput = MakeKvp(-1)
    Set myComp.Program = myProgram
    myComp.Run myInput
    Debug.Print "Test 3B", 0 - CLng(myComp.GetOutput.GetLast)
    
    'Test 3: Using immediate mode:equal to 8: input greater than
    Set myProgram = MakeLongLongVsLongLongKvp(Array(3, 3, 1108, -1, 8, 3, 4, 3, 99))
    Set myInput = MakeKvp(1001)
    Set myComp.Program = myProgram
    myComp.Run myInput
    Debug.Print "Test 3C", 0 - CLng(myComp.GetOutput.GetLast)
    Debug.Print
    
    'Test4: Using immediateMode: less than 8: input less than
    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(3, 3, 1107, -1, 8, 3, 4, 3, 99))
    myComp.Run MakeKvp(0)
    Debug.Print "Test 4A", 1 - CLng(myComp.GetOutput.GetLast)
    
    'Test4: Using immediateMode: less than 8: input 8
    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(3, 3, 1107, -1, 8, 3, 4, 3, 99))
    myComp.Run MakeKvp(8)
    Debug.Print "Test 4B", 0 - CLng(myComp.GetOutput.GetLast)
    
    'Test5: Using position mode: jump input 0
    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(3, 12, 6, 12, 15, 1, 13, 14, 13, 4, 13, 99, -1, 0, 1, 9))
    myComp.Run MakeKvp(0)
    Debug.Print "Test 5A", 0 - CLng(myComp.GetOutput.GetLast)
    
    'Test5: Using position mode: jump input not 0
    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(3, 12, 6, 12, 15, 1, 13, 14, 13, 4, 13, 99, -1, 0, 1, 9))
    myComp.Run MakeKvp(9)
    Debug.Print "Test 5B", 1 - CLng(myComp.GetOutput.GetLast)
    Debug.Print
    
    'Test 6 Using Immediate mode: Input 0
    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(3, 3, 1105, -1, 9, 1101, 0, 0, 12, 4, 12, 99, 1))
    myComp.Run MakeKvp(0)
    Debug.Print "Test 6A", 0 - CLng(myComp.GetOutput.GetLast)
    
    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(3, 3, 1105, -1, 9, 1101, 0, 0, 12, 4, 12, 99, 1))
    myComp.Run MakeKvp(9)
    Debug.Print "Test 6B", 1 - CLng(myComp.GetOutput.GetLast)
    Debug.Print
    
    'Test for 8: input <8
    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(3, 21, 1008, 21, 8, 20, 1005, 20, 22, 107, 8, 21, 20, 1006, 20, 31, 1106, 0, 36, 98, 0, 0, 1002, 21, 125, 20, 4, 20, 1105, 1, 46, 104, 999, 1105, 1, 46, 1101, 1000, 1, 20, 4, 20, 1105, 1, 46, 98, 99))
    myComp.Run MakeKvp(0)
    Debug.Print "Test 7A", 999 - CLng(myComp.GetOutput.GetLast)
    
    'Test for 8: input =8
    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(3, 21, 1008, 21, 8, 20, 1005, 20, 22, 107, 8, 21, 20, 1006, 20, 31, 1106, 0, 36, 98, 0, 0, 1002, 21, 125, 20, 4, 20, 1105, 1, 46, 104, 999, 1105, 1, 46, 1101, 1000, 1, 20, 4, 20, 1105, 1, 46, 98, 99))
    myComp.Run MakeKvp(8)
    Debug.Print "Test 7B", 1000 - CLng(myComp.GetOutput.GetLast)
    
    'Test7: Test for 8: input>8
    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(3, 21, 1008, 21, 8, 20, 1005, 20, 22, 107, 8, 21, 20, 1006, 20, 31, 1106, 0, 36, 98, 0, 0, 1002, 21, 125, 20, 4, 20, 1105, 1, 46, 104, 999, 1105, 1, 46, 1101, 1000, 1, 20, 4, 20, 1105, 1, 46, 98, 99))
    myComp.Run MakeKvp(1080)
    Debug.Print "Test 7C", 1001 - CLng(myComp.GetOutput.GetLast)
    Debug.Print
    
    Debug.Print "Day 9 Updates"
    'Test 8: Copy of itself
    Set myProgram = MakeLongLongVsLongLongKvp(Array(109, 1, 204, -1, 1001, 100, 1, 100, 1008, 100, 16, 101, 1006, 101, 0, 99))
    Set myComp.Program = myProgram
    myComp.Run
    Debug.Print "Test 8", IIf(InStr(myComp.GetOutput.GetValuesAsString, myProgram.GetValuesAsString) > 0, 0, 1)
    
    'Test 9
    Set myComp = New IntComputer
    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(1102, 34915192, 34915192, 7, 4, 7, 99, 0))
    myComp.Run
    Debug.Print "Test 9", 16 - Len(CStr(myComp.GetOutput.GetLast))
    
    'Set myComp = New IntComputer
    Set myComp.Program = MakeLongLongVsLongLongKvp(Array(104, 112589990684262^, 99))
    myComp.Run
    Debug.Print "Test 10", 112589990684262^ - CLngLng(myComp.GetOutput.GetLast)
        
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod1()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

