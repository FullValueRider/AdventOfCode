Attribute VB_Name = "Day07"
Option Explicit

Const DAY07_INPUT_PATH_AND_NAME                     As String = "C:\Users\slayc\source\repos\VBA\AdventOfCode\2019\Day07Input.txt"
Const DAY07_NUMBER_OF_AMPLIFIERS                    As Long = 5

Private Type State

    Amplifiers                                      As AmplifierBank
    PhaseSettings                                   As Collection
    Program                                         As KvpOD
    Configuration                                   As Configuration
    
End Type

Private s                                   As State


Public Sub Day7Part1()
    
    s.Configuration = Configuration.SinglePass
    Set s.PhaseSettings = GetListOfPhaseSettings("0,1,2,3,4")
    
    Debug.Print Fmt("The day 7 Part 1 answer should be 255590: {0}", GetMaximumAmplifierOutput)
    
End Sub


Public Sub Day7Part2()

    s.Configuration = Configuration.Looped
    Set s.PhaseSettings = GetListOfPhaseSettings("5,6,7,8,9")
    
    Debug.Print Fmt("The day 7 Part 2 answer should be 58285150: {0}", GetMaximumAmplifierOutput)
    
End Sub


Public Function GetMaximumAmplifierOutput() As Long

    Set s.Amplifiers = AmplifierBank(DAY07_NUMBER_OF_AMPLIFIERS)
    Set s.Program = MakeProgram(Split(Common.GetFileAsString(DAY07_INPUT_PATH_AND_NAME), ","))
    
    Dim myPhaseSettings As Variant
    Dim myAmplifiersOutput As Long
    Dim myMaxAmplifiersOutput As Long
    For Each myPhaseSettings In s.PhaseSettings
    
        myAmplifiersOutput = GetAmplifierOutput(myPhaseSettings)
        If myMaxAmplifiersOutput < myAmplifiersOutput Then myMaxAmplifiersOutput = myAmplifiersOutput
    
    Next
    
    GetMaximumAmplifierOutput = myMaxAmplifiersOutput
    
End Function


Public Function GetAmplifierOutput(ByVal ipPhaseSettings As String) As Long

    With s.Amplifiers
        
        .PhaseSettings = ipPhaseSettings
        .Configuration = s.Configuration
        Set .Program = s.Program
        .Run
        GetAmplifierOutput = .Output
        
    End With
    
End Function


Public Function GetListOfPhaseSettings(ByVal ipPhaseOptions As String) As Collection
    
    Dim myPhaseOptions As Variant: myPhaseOptions = Split(ipPhaseOptions, ",")
    Dim mycoll As Collection: Set mycoll = New Collection
    Dim i As Variant
    For Each i In myPhaseOptions
        
        Dim j As Variant
        For Each j In myPhaseOptions
        
            If InStr(i, j) = 0 Then
            
                Dim K As Variant
                For Each K In myPhaseOptions
                
                    If InStr(i & j, K) = 0 Then
                    
                        Dim l As Variant
                        For Each l In myPhaseOptions
                        
                            If InStr(i & j & K, l) = 0 Then
                            
                                Dim m As Variant
                                For Each m In myPhaseOptions
                                
                                    If InStr(i & j & K & l, m) = 0 Then
                                    
                                        mycoll.Add i & j & K & l & m
                                    
                                    End If
                                
                                Next
                                
                            End If
                            
                        Next
                        
                    End If
                    
                Next
                
            End If
            
        Next
        
    Next
        
    Set GetListOfPhaseSettings = mycoll
    
End Function
