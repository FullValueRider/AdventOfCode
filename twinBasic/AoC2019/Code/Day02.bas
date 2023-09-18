Attribute VB_Name = "Day02"
Option Explicit

Public Sub TestCOmputer()

Const NOUN_INDEX                            As LongLong = 1&
Const VERB_INDEX                            As LongLong = 2&

    Dim myProgram As Kvp
    Set myProgram = GetDay02Program
    
    Dim myComp As IntComputer
    Set myComp = New IntComputer
    
    Dim myNoun As LongLong
    For myNoun = 0 To 99

        Debug.Print myNoun
        
        Dim myVerb As LongLong
        For myVerb = 0 To 99
            
            Set myComp.Program = myProgram.Clone
           
            myComp.Program.SetItem NOUN_INDEX, myNoun
            myComp.Program.SetItem VERB_INDEX, myVerb
            myComp.Run
            
            If myComp.Program.GetFirst = 19690720 Then
    
                Debug.Print "The Day 2 Part 2 answer should be 6635", 100 * myComp.Program.Item(NOUN_INDEX) + myComp.Program.Item(VERB_INDEX)
                Exit Sub
                
            End If
            
        Next

    Next

End Sub

Public Sub Day2Part1()

    Dim myComp As IntComputer
    Set myComp = New IntComputer
    
    Dim myProgram As Kvp
    Set myProgram = GetDay02Program
    
    myProgram.SetItem CLngLng(1), CLngLng(12)
    myProgram.SetItem CLngLng(2), CLngLng(2)

    Set myComp.Program = myProgram
    myComp.Run
    
    Debug.Print "The Day2 Part 2 answer should be 4138687", myComp.Program.GetFirst
    
    
End Sub






Public Function GetDay02Program() As Kvp

    Dim myFso  As Scripting.FileSystemObject: Set myFso = New Scripting.FileSystemObject
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile("C:\Projects\Code\AdventOfCode\Day02\Day02Program.txt", ForReading)
    
    Dim myProgram As String
    myProgram = myfile.ReadAll
    myfile.Close
    
    Dim myArray As Variant
    myArray = Split(myProgram, ",")
    
    Set GetDay02Program = MakeLongLongVsLongLongKvp(myArray)
    
End Function
