Attribute VB_Name = "Day09"
Option Explicit

'@Ignore ProcedureNotUsed

'@Ignore ProcedureNotUsed
Private Sub Day9Part1()

    Dim myComp As IntComputer: Set myComp = New IntComputer
    Set myComp.Program = GetDay9Program
    
    Dim myInput As Kvp: Set myInput = New Kvp
    myInput.AddByIndex 1&
    myComp.Run myInput
    Debug.Print "The answer for Day9 Part9 should be 3100786347", CLngLng(myComp.GetOutput.GetLast)
    
End Sub

'@Ignore ProcedureNotUsed
Private Sub Day9Part2()

    Dim myComp As IntComputer: Set myComp = New IntComputer
    Set myComp.Program = GetDay9Program
    
    Dim myInput As Kvp: Set myInput = New Kvp
    myInput.AddByIndex 2&
    myComp.Run myInput
    Debug.Print "The answer for Day9 Part9 should be 87023", CLngLng(myComp.GetOutput.GetLast)
    
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

