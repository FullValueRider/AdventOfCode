Attribute VB_Name = "MainModule"
Option Explicit

Public Const AoC2021Data As String = "C:\Users\slayc\source\repos\AoC2021\RawData\"

' This project type is set to 'Standard EXE' in the Settings file, so you need a Main() subroutine to run when the EXE is started.
Public Sub Main()
    
    Dim myDuration As Variant
    myDuration = Timer
    
    Day01.Execute
    Day02.Execute
    Day03.Execute
    Day04.Execute
    Day05.Execute
    Day06.Execute
    Day07.Execute
    Day08.Execute
    Day09.Execute
    Day10.Execute
    Day11.Execute
   ' Day12.Execute
    Day13.Execute
    Day14.Execute
    myDuration = Timer - myDuration
    Debug.Print "Executed in " & myDuration & " Seconds"
    
End Sub
