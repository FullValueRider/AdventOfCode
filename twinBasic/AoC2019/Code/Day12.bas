Attribute VB_Name = "Day12"
Option Explicit



Public Function GetDay12Input() As Kvp

    Dim myFso  As Scripting.FileSystemObject: Set myFso = New Scripting.FileSystemObject
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile("C:\Projects\Code\AdventOfCode\Day12\Day12Coordinates.txt", Scripting.IOMode.ForReading)
        
    Dim myMap  As Kvp: Set myMap = New Kvp
    
    Do
    
        myMap.AddByIndex myfile.ReadLine
        
    Loop Until myfile.AtEndOfStream
        
    myfile.Close
    Set GetDay12Input = myMap
    
End Function
