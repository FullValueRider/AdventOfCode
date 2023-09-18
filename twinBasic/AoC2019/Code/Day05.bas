Attribute VB_Name = "Day05"
Option Explicit

'@Folder("AdventOfCode")

'@Ignore ConstantNotUsed

Public Sub Day5Part1()

    Dim myProgram As Kvp
    Set myProgram = GetDay05Program
    
    Dim myInput As Kvp: Set myInput = New Kvp
    myInput.AddByIndex 1&
    
    Dim myComp As IntComputer: Set myComp = New IntComputer
    Set myComp.Program = myProgram
    myComp.OutputMode = HaltOnOutput
    Debug.Print myInput.GetValuesAsString
    myComp.Run myInput
    
    Do Until myComp.RunHasCompleted
    
        Debug.Print myComp.GetOutput.GetValuesAsString
        myComp.Run
        
    Loop
    
    Debug.Print "The Day 5 Part 1 output should be 7988899", myComp.GetOutput.GetFirst
    
    
End Sub








'End Sub






Private Function GetDay05Program() As Kvp

    Dim myFso  As Scripting.FileSystemObject: Set myFso = New Scripting.FileSystemObject
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile("C:\Projects\Code\AdventOfCode\Day05\Day05Program.txt", ForReading)
    Dim myProgram As String:  myProgram = myfile.ReadAll
    myfile.Close
    
    Set GetDay05Program = MakeLongLongVsLongLongKvp(Split(myProgram, ","))
    
End Function
