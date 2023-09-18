Attribute VB_Name = "Day01"
Option Explicit
'@Folder("AdventOfCode")

Const DAY01_INPUT_PATH_AND_NAME As String = "C:\Users\slayc\source\repos\VBA\AdventOfCode\2019\Day01Input.txt"
Public Sub Part1()

    Dim myComponents As KvpOD
    Set myComponents = Common.GetFileByLines(DAY01_INPUT_PATH_AND_NAME)
    
    Dim myFuel As Long
    Dim myComponent As KVPair
    For Each myComponent In myComponents
        DoEvents
        myFuel = myFuel + Int(myComponent.Value / 3) - 2
    
    Next

    Debug.Print Fmt("The answer to Day 01 Part 1 should be 3364035: {0} ", myFuel)
End Sub


Public Sub Part2()
    
    Dim myComponents As KvpOD
    Set myComponents = Common.GetFileByLines(DAY01_INPUT_PATH_AND_NAME)
    
    Dim myComponent As KVPair
    For Each myComponent In myComponents
        DoEvents
        Dim myFuel As Long
        myFuel = myFuel + ComponentFuel(myComponent.Value)
    
    Next
    
    Debug.Print Fmt("The answer to Day 01 Part 2 should be 5043167: {0}", myFuel)
    
End Sub


Private Function ComponentFuel(ByVal ipComponent As Long) As Long

    Dim myFuel As Long
    myFuel = Int(CDbl(ipComponent) / CDbl(3)) - 2
    
    If myFuel <= 0 Then
    
        ComponentFuel = 0
        
    Else
        
        ComponentFuel = myFuel + ComponentFuel(myFuel)
        
    End If
        
End Function
