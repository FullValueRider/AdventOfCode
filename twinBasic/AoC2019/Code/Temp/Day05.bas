Attribute VB_Name = "Day05"
Option Explicit

'@Folder("AdventOfCode")

Const DAY05_INPUT_PATH_AND_NAME         As String = "C:\Users\slayc\source\repos\VBA\AdventOfCode\2019\Day05Input.txt"

Public Sub Day5Part1()

    Dim myResult As String
    myResult = Day05Common(1&)
    
    Debug.Print "The Day 5 Part 1 answer should be a sequence of 0 followed by 7988899", myResult
    
    
End Sub

Public Sub Day5Part2()

    Dim myResult As String
    myResult = Day05Common(5&)
    
    Debug.Print "The Day 5 Part 2 answer should be 13758663", myResult
    
End Sub
Public Function Day05Common(ByVal ipDiagnosticCode As Long) As String

    Dim myProgram As KvpOD
    Set myProgram = MakeProgram(Split(Common.GetFileAsString(DAY05_INPUT_PATH_AND_NAME), ","))
    
    Dim myInput As KvpOD: Set myInput = New KvpOD
    myInput.AddByIndex ipDiagnosticCode
    
    Dim myComp As IntComputer: Set myComp = New IntComputer
    Set myComp.Program = myProgram
    myComp.OutputMode = OutputResponse.ContinueOnOutput
    myComp.Run myInput
    
    Day05Common = myComp.GetOutput.GetValuesAsString
    
End Function


'Private Function GetDay05Program() As KvpOD
'
'    Dim myFso  As Scripting.FileSystemObject: Set myFso = New Scripting.FileSystemObject
'    Dim myfile As TextStream
'    Set myfile = myFso.OpenTextFile("C:\Users\slayc\source\repos\VBA\AdventOfCode\2019\Day05Input.txt", ForReading)
'    Dim myProgram As String:  myProgram = myfile.ReadAll
'    myfile.Close
'
'    Set GetDay05Program = MakeProgram(Split(myProgram, ","))
'
'End Function
