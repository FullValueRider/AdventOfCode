Attribute VB_Name = "Day12"
Option Explicit



Public Sub Day12Part1()

    Dim myMoons As KvpOD
    Set myMoons = GetDay12Moons(GetDay12Input)
    'Set myMoons = GetDay12Moons(TestCase)
    '@Ignore VariableNotUsed
    ShowMoons myMoons
    '@Ignore VariableNotUsed
    Dim myTime As Long
    For myTime = 1 To 1000
    
        UpdateGravityForAllMoons myMoons
        UpdateVelocityForAllMoons myMoons
       ' ShowMoons myMoons
    Next
    
    
    Debug.Print "The day12 Part1 answer is ", SystemEnergy(myMoons)
    
End Sub

'Public Sub Day12Part2()
'
'    Dim myMoons As Kvp
'    'Set myMoons = GetDay12Moons(GetDay12Input)
'    Set myMoons = GetDay12Moons(TestCase)
'    Dim myTime As Long
'    Dim myStates As KvpOD: Set myStates = New KvpOD
'    Dim myOverallTime As Long
'
'    Do
'
'        DoEvents
'        Debug.Print myOverallTime
'        For myTime = myOverallTime To myOverallTime + 10000
'
'            UpdateGravityForAllMoons myMoons
'            UpdateVelocityForAllMoons myMoons
'            Dim myState As String
'            myState = GetSystemState(myMoons)
'            If myStates.LacksKey(myState) Then
'
'                myStates.AddByKey myState, myTime
'
'            Else
'
'                Exit Do
'
'            End If
'
'        Next
'
'        myOverallTime = myOverallTime + 1000
'
'    Loop
'
'
'    Debug.Print myStates.GetFirst.ValueKey
'    Debug.Print myState
'    Debug.Print "The answer to day 12 Part 12 is state " & myState & " repeats after Time " & myTime + 2
'End Sub


Public Sub Day12Part2()

Dim myMoons As KvpOD

    Const Moon1                             As Long = 0&
    Const Moon2                             As Long = 1&
    Const Moon3                             As Long = 2&
    Const Moon4                             As Long = 3&
    
    Dim myFoundMoon(Moon1 To Moon4)         As Boolean
    
    'Set myMoons = GetDay12Moons(GetDay12Input)
    Set myMoons = GetDay12Moons(TestCase)
    Dim myTime As Long
    Dim myMoonStates As KvpOD: Set myMoonStates = New KvpOD
    Dim myMoon As Long
    For myMoon = Moon1 To Moon4
    
        myMoonStates.AddByKey myMoon, New KvpOD
        myFoundMoon(myMoon) = False
        
    Next
        
    Do
        
        '@Ignore FunctionReturnValueDiscarded
        DoEvents
        UpdateGravityForAllMoons myMoons
        UpdateVelocityForAllMoons myMoons
        
        For myMoon = Moon1 To Moon4
        
            If Not myFoundMoon(myMoon) Then
            
                Dim myState As String
                myState = GetMoonState(myMoons.Item(myMoon))
                If myMoonStates.Item(myMoon).LacksKey(myState) Then
                
                    myMoonStates.Item(myMoon).AddByKey myState, myTime
                    
                Else
                
                    myFoundMoon(myMoon) = True
                    Debug.Print "Period for moon " & myMoon + 1 & " is " & myTime
                    If myFoundMoon(Moon1) And myFoundMoon(Moon2) And myFoundMoon(Moon3) And myFoundMoon(Moon4) Then Exit Do
                    
                End If
                
            End If
        
        Next
        
        myTime = myTime + 1
        
    Loop

End Sub


Public Sub ShowMoons(ByRef ipMoons As KvpOD)
    Dim myMoon As Moon
    For Each myMoon In ipMoons
    'Debug.Print "x=" & VBA.Format$(VBA.Format$(2, "##"), "@@@")
    Debug.Print _
        "pos=<X=" & VBA.Format$(VBA.Format$(myMoon.XPos), "@@@") _
        & ", y=" & VBA.Format$(VBA.Format$(myMoon.YPos), "@@@") _
        & ", z=" & VBA.Format$(VBA.Format$(myMoon.ZPos), "@@@") _
        & ">, Vel=<x=" & VBA.Format$(VBA.Format$(myMoon.XVel), "@@@") _
        & ", y=" & VBA.Format$(VBA.Format$(myMoon.YVel), "@@@") _
        & ", z=" & VBA.Format$(VBA.Format$(myMoon.ZVel), "@@@") _
        & ">"
    
    Next
    
    Debug.Print: Debug.Print
    
End Sub

Public Function GetMoonState(ByRef ipMoon As Moon) As String

    GetMoonState = _
        CStr(ipMoon.XPos) & "," _
        & CStr(ipMoon.YPos) & "," _
        & CStr(ipMoon.ZPos) & "," _
        & CStr(ipMoon.XVel) & "," _
        & CStr(ipMoon.YVel) & "," _
        & CStr(ipMoon.ZVel) & ","

End Function


'@Ignore FunctionReturnValueNotUsed
Public Function SystemEnergy(ByRef ipMoons As KvpOD) As Long

    Dim myEnergy As Long
    Dim myMoon As Moon
    For Each myMoon In ipMoons
    
        myEnergy = myEnergy + myMoon.TotalEnergy
        
    Next
    
    SystemEnergy = myEnergy
End Function


Public Sub UpdateVelocityForAllMoons(ByRef ipMoons As KvpOD)

    Dim myMoon As Moon
    For Each myMoon In ipMoons
    
        myMoon.ApplyVelocity
    
    Next
    
End Sub


Public Sub UpdateGravityForAllMoons(ByRef ipMoons As KvpOD)

    Dim myIndex As Long
    Dim myStart As Long
    myStart = 0
    Do While myStart < ipMoons.Count
        
        Dim myFirstMoon As Moon
        Set myFirstMoon = ipMoons.Item(myStart)
        myStart = myStart + 1
        For myIndex = myStart To ipMoons.Count - 1
        
            myFirstMoon.ApplyGravity ipMoons.Item(myIndex)
            
        Next
        
    Loop
    
End Sub


Public Function GetDay12Input() As KvpOD

    Dim myFso  As Scripting.FileSystemObject: Set myFso = New Scripting.FileSystemObject
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile("C:\Projects\Code\AdventOfCode\Day12\Day12Coordinates.txt", Scripting.IOMode.ForReading)
        
    Dim myMap  As KvpOD: Set myMap = New KvpOD
    
    Do
    
        myMap.AddByIndex myfile.ReadLine
        
    Loop Until myfile.AtEndOfStream
        
    myfile.Close
    Set GetDay12Input = myMap
    Debug.Print myMap.GetValuesAsString
End Function


Public Function GetDay12Moons(ByVal ipInputStrings As KvpOD) As KvpOD

    Dim myItem As Variant
    Dim myMoons As KvpOD: Set myMoons = New KvpOD
    For Each myItem In ipInputStrings
    
        myMoons.AddByIndex GetMoonFromString(myItem)
        
    Next
    
    Set GetDay12Moons = myMoons
    
End Function


Public Function GetMoonFromString(ByVal ipString As String) As Moon
        
     Dim myAxes As Variant
     myAxes = Split(CleanMoonString(ipString), ",")
     
     Dim myAxis As Variant
     '@Ignore VariableNotAssigned
     Dim myMoon As KvpOD: Set myMoon = New KvpOD
     For Each myAxis In myAxes
     
         Dim myItems As Variant
         myItems = Split(myAxis, "=")
         '@Ignore UnassignedVariableUsage
         myMoon.AddByKey LCase$(Trim$(myItems(0))), CLng(myItems(1))
     
     Next
     
    '@Ignore UnassignedVariableUsage
    Set GetMoonFromString = Moon.Debutante(myMoon.Item("x"), myMoon.Item("y"), myMoon.Item("z"))
End Function

Public Function CleanMoonString(ByVal ipString As String) As String

    Dim myArrayStr As String
    '@Ignore AssignmentNotUsed
    myArrayStr = ipString
    myArrayStr = Replace(myArrayStr, "<", vbNullString)
    myArrayStr = Replace(myArrayStr, ">", vbNullString)
    CleanMoonString = myArrayStr
    
End Function


Public Function TestCase() As KvpOD

    Dim myStringKvp As KvpOD: Set myStringKvp = New KvpOD
    myStringKvp.AddByIndex "<x=-1, y=0, z=2>"
    myStringKvp.AddByIndex "<x=2, y=-10, z=-7>"
    myStringKvp.AddByIndex "<x=4, y=-8, z=8>"
    myStringKvp.AddByIndex "<x=3, y=5, z=-1>"
    Set TestCase = myStringKvp
    
End Function


