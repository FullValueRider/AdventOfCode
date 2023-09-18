Attribute VB_Name = "Day02"
Option Explicit

Const DAY01_INPUT_PATH_AND_NAME             As String = "C:\Users\slayc\source\repos\VBA\AdventOfCode\2019\Day02Input.txt"
Const NOUN_INDEX                            As Currency = 1&
Const VERB_INDEX                            As Currency = 2&

Public Sub Part1()

    Dim myComp As IntComputer
    Set myComp = New IntComputer
    
    Dim myProgram As KvpOD
    Set myProgram = MakeProgram(Split(Common.GetFileAsString(DAY01_INPUT_PATH_AND_NAME), ","))
    
    myProgram.Item(NOUN_INDEX) = 12
    myProgram.Item(VERB_INDEX) = 2

    Set myComp.Program = myProgram
    myComp.Run
    
    Debug.Print Fmt("The answer for Day02 Part 1 should be 4138687: {0}", myComp.Program.GetFirst.Value)
    
End Sub

Public Sub Part2()

    Dim myProgram As KvpOD
    Set myProgram = MakeProgram(Split(Common.GetFileAsString(DAY01_INPUT_PATH_AND_NAME), ","))
    
    Dim myComp As IntComputer
    Set myComp = New IntComputer
    
    Dim myNoun As Currency
    For myNoun = 0 To 99
        DoEvents
        
        Dim myVerb As Currency
        For myVerb = 0 To 99
            DoEvents
            Set myComp.Program = myProgram.Clone
           
            myComp.Program.Item(NOUN_INDEX) = myNoun
            myComp.Program.Item(VERB_INDEX) = myVerb
            myComp.Run
            
            If myComp.Program.GetFirst.Value = 19690720 Then
    
                Debug.Print Fmt("The answer for Day02 Part 2  should be 6635", 100 * myComp.Program.Item(NOUN_INDEX) + myComp.Program.Item(VERB_INDEX))
                Exit Sub
                
            End If
            
        Next

    Next

End Sub
