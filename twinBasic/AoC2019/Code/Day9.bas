Attribute VB_Name = "Day9"
Option Explicit

Private Sub TestComputer()

Dim myComp As IntComputer
 
    Debug.Print "Computer Tests"
    Set myComp = New IntComputer
    Set myComp.Program = MakeKvp(109, 1, 204, -1, 1001, 100, 1, 100, 1008, 100, 16, 101, 1006, 101, 0, 99)
    Debug.Print "Program", myComp.Program.GetValuesAsString
    Dim myOutput As Kvp: Set myOutput = myComp.Run(Nothing)
    Debug.Print "Test 1", myOutput.GetValuesAsString
    
    Set myComp.Program = MakeKvp(1102, 34915192, 34915192, 7, 4, 7, 99, 0)
    Debug.Print "Program", myComp.Program.GetValuesAsString
    Set myOutput = myComp.Run(Nothing)
    Debug.Print "Test 2", 16 - Len(CStr(CLngLng(myOutput.GetLast)))
    
    Set myComp.Program = MakeKvp(104, 1.12589990684262E+15, 99)
    Debug.Print "Program", myComp.Program.GetValuesAsString
    Set myOutput = myComp.Run(Nothing)
    Debug.Print "Test 3", 1125899906842624^ - CLngLng(myOutput.GetLast)
End Sub

Private Sub Day9Part1()

    Dim myComp As IntComputer: Set myComp = New IntComputer
    Set myComp.Program = GetDay9Program
    
    Dim myInput As Kvp: Set myInput = New Kvp
    myInput.AddByIndex 1^
    Dim myOutput As Kvp
    Set myOutput = myComp.Run(myInput)
    Debug.Print CLngLng(myOutput.GetLast)
    
End Sub

Private Sub Day9Part2()

    Dim myComp As IntComputer: Set myComp = New IntComputer
    Set myComp.Program = GetDay9Program
    
    Dim myInput As Kvp: Set myInput = New Kvp
    myInput.AddByIndex 2^
    Dim myOutput As Kvp
    Set myOutput = myComp.Run(myInput)
    Debug.Print CLngLng(myOutput.GetLast)
    
End Sub

Private Function GetDay9Program() As Kvp

    Dim myFso  As Scripting.FileSystemObject: Set myFso = New Scripting.FileSystemObject
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile("C:\Projects\Code\AdventOfCode\Day9\Day9Input.txt", ForReading)
    Dim myProgram As String:  myProgram = myfile.ReadAll
    myfile.Close
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.AddByIndexFromArray Split(myProgram, ",")
    Set GetDay9Program = myKvp
    
End Function

