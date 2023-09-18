Attribute VB_Name = "AllDays"
Option Explicit

Public Const AoCRawData As String = "C:\Users\slayc\source\repos\AdventOfCode\RawData\"
Public Const Year       As String = "\2015"

#Const UseTwinbasic = False

Public Sub Main()

    Dim myDuration As Variant
    myDuration = Timer

'    Day01.Execute
'    Day02.Execute
'    Day03.Execute
'    Day04.Execute
'    Day05.Execute
'    Day06.Execute
'    Day07.Execute
'    Day08.Execute
'    Day09.Execute
'    Day10.Execute
'    Day11.Execute
'    dAY12.Execute 'part 2 not working
'    Day13.Execute
'    Day14.Execute
'     Day15.Execute
    myDuration = Timer - myDuration
    Debug.Print "Executed in " & myDuration & " Seconds"
    
End Sub



