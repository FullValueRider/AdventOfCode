Attribute VB_Name = "Day01"
Option Explicit
'@Folder("AdventOfCode")

Public Sub Day1Part1()

    Dim myComponents As Kvp
    Set myComponents = GetDay01ComponentMasses
    
    Dim myFuel As Long
    Dim myComponent As Variant
    For Each myComponent In myComponents
    
        myFuel = myFuel + (myComponent \ 3) - 2
    
    Next

    Debug.Print "Part 1 fuel requirement should be 3364035 ", myFuel
End Sub


Public Sub Day1Part2()
    
    Dim myComponents As Kvp
    Set myComponents = GetDay01ComponentMasses
    
    Dim myComponent As Variant
    Dim myFuel As Long
    For Each myComponent In myComponents
    
        myFuel = myFuel + ComponentFuel(CLng((myComponent) \ 3) - 2)
    
    Next
    
    Debug.Print "Part 2 fuel requirement should be 5043167 ", myFuel
    
End Sub


Public Function ComponentFuel(ByVal ipComponent As Long) As Long

    If ipComponent <= 0 Then
    
        ComponentFuel = 0
        Exit Function
        
    Else
        
        ComponentFuel = ipComponent + ComponentFuel((ipComponent \ 3) - 2)
        
    End If
        
End Function


Public Function GetDay01ComponentMasses() As Kvp

    Dim myFso  As Scripting.FileSystemObject
    Set myFso = New Scripting.FileSystemObject
    
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile("C:\Projects\Code\AdventOfCode\Day01\ComponentMasses.txt", Scripting.IOMode.ForReading)
        
    Dim myMasses As Kvp: Set myMasses = New Kvp
    Dim myIndex As Long
    myIndex = 0
    Do
    
        Dim myMass As String
        myMass = myfile.ReadLine()
        myMasses.AddByKey myIndex, CLng(myMass)
        myIndex = myIndex + 1
        
    Loop Until myfile.AtEndOfStream
    
    myfile.Close
    Set GetDay01ComponentMasses = myMasses
    
End Function









