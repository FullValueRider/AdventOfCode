Attribute VB_Name = "AllDays"
Option Explicit

Public Const AoC2015Data As String = "C:\Users\slayc\source\repos\AoC2015\RawData\"
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
    
    
    myDuration = Timer - myDuration
    Debug.Print "Executed in " & myDuration & " Seconds"
    
End Sub
