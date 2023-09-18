Attribute VB_Name = "Day11"
Option Explicit

Public Enum PanelColour

    Black = 0&
    White = 1&
    
End Enum


Public Type State

    Robot                   As PaintingRobot

End Type


Public Sub Day11Part1()
    
    
    Dim myRobot As PaintingRobot: Set myRobot = PaintingRobot.Debutante
    Set myRobot.Program = GetDay11Program
    
    Dim myInput As Kvp
    Set myInput = MakeKvp(0)
    
    myRobot.Run myInput
    
    Debug.Print "The Day11 Part1 answer is 1185", myRobot.Track.Count
        
    Debug.Print "Finished"
    
    
    
End Sub

Public Sub Day11Part2()
    Dim myRobot As PaintingRobot: Set myRobot = PaintingRobot.Debutante
    Set myRobot.Program = GetDay11Program
    
    Dim myInput As Kvp
    Set myInput = MakeKvp(1)
    
    myRobot.Run myInput
    
    Debug.Print "The Day11 Part2 answer is in the displayed Excel worksheet and should be BFEAGHAF"
    
End Sub

Private Function GetDay11Program() As Kvp

    Dim myFso  As Scripting.FileSystemObject: Set myFso = New Scripting.FileSystemObject
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile("C:\Projects\Code\AdventOfCode\Day11\RobotProgram.txt", ForReading)
    
    Dim myProgram As String
    myProgram = myfile.ReadAll
    myfile.Close
    
    Dim myItem As Variant
    Dim myKvp As Kvp: Set myKvp = New Kvp
    Dim myCode As Variant
    myCode = Split(myProgram, ",")
    Dim myIndex As LongLong
    myIndex = 0^
    For Each myItem In myCode
        
        myKvp.AddByKey myIndex, CLngLng(myItem)
        myIndex = myIndex + 1^
    
    Next
    
    Set GetDay11Program = myKvp
    
End Function

