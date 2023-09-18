VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@PredeclaredId
Private Type State

    RawData                       As String
    
End Type

Private s                       As State


Public Sub Part1()

    s.RawData = Common.GetRawDataByFile("C:\Users\slayc\source\repos\VBA\AdventOfCode\2015\Day01RawData.txt")
    
    Dim myUp As Long
    myUp = Len(s.RawData) - Len(Replace(s.RawData, "(", vbNullString))
    
    Dim myDown As Long
    myDown = Len(s.RawData) - myUp
    
    Dim myFloor As Long
    myFloor = myUp - myDown
    
    Debug.Print "The answer to Day 1 Part 1 is 74. Found is ", myFloor
    
End Sub

Public Sub Part2()

Const UP                            As String = "("

    Dim myInstructions As Kvp: Set myInstructions = New Kvp
    myInstructions.SetFirstIndexAsLong 1
    myInstructions.AddByIndexAsCharacters s.RawData
    Dim myInstruction As Variant
    Dim myFloor As Long
    myFloor = 0
    For Each myInstruction In myInstructions
    
        myFloor = myFloor + IIf(myInstruction.Value = UP, 1, -1)
        If myFloor = -1 Then
    
            Debug.Print "The answer to day 1 part 2 should be 1795 ", myInstruction.Key
            Exit For
            
        End If
        
    Next
    
    

End Sub
