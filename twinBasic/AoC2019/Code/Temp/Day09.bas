Attribute VB_Name = "Day09"
Option Explicit

'@Ignore ProcedureNotUsed

'@Ignore ProcedureNotUsed
Private Sub Day9Part1()

    Dim myComp As IntComputer: Set myComp = New IntComputer
    Set myComp.Program = GetDay9Program
    
    Dim myInput As KvpOD: Set myInput = New KvpOD
    myInput.AddByIndex 1&
    myComp.Run myInput
    Debug.Print "The answer for Day9 Part9 should be 3100786347", CLngLng(myComp.GetOutput.GetLast.Value)
    
End Sub

'@Ignore ProcedureNotUsed
Private Sub Day9Part2()

    Dim myComp As IntComputer: Set myComp = New IntComputer
    Set myComp.Program = GetDay9Program
    
    Dim myInput As KvpOD: Set myInput = New KvpOD
    myInput.AddByIndex 2&
    myComp.Run myInput
    Debug.Print "The answer for Day9 Part9 should be 87023", CLngLng(myComp.GetOutput.GetLast.Value)
    
End Sub

Private Function GetDay9Program() As KvpOD
    
    Dim myFso  As Scripting.FileSystemObject: Set myFso = New Scripting.FileSystemObject
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile("C:\Users\slayc\source\repos\VBA\AdventOfCode\2019\Day09Input.txt", ForReading)
    Dim myProgram As String:  myProgram = myfile.ReadAll
    myfile.Close
    Dim myKvp As KvpOD: Set myKvp = New KvpOD
    myKvp.AddByIndexFromArray Split(myProgram, ",")
    Set GetDay9Program = myKvp
    
End Function

