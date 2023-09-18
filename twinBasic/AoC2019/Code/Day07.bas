Attribute VB_Name = "Day07"
Option Explicit

Public Enum AmplifierModeEnum

    SinglePass = 1
    MultiPass = 2
    
End Enum

Public Enum OutputModeEnum

    HaltOnOutput = 1
    ContinueOnOutput = 2

End Enum


Private Type State

    Amps                                    As Collection

End Type

Private s                                   As State

Public Sub Day7Part1()

    Dim myProgram As Kvp: Set myProgram = GetDay7Program
    Dim myThrusterSetting As String:
    
    myThrusterSetting = GetMaxAmplifierOutput("0,1,2,3,4", AmplifierModeEnum.SinglePass, myProgram)
    Debug.Print "The day 7 Part 1 answer should be 255590", myThrusterSetting
    
End Sub


Public Sub Day7Part2()

    Dim myProgram As Kvp
    Set myProgram = GetDay7Program
    
    Dim myThrusterSetting As String
    myThrusterSetting = GetMaxAmplifierOutput("5,6,7,8,9", AmplifierModeEnum.MultiPass, myProgram)
    
    Debug.Print "The day 7 part 2 answer should be 58285150", myThrusterSetting
    
End Sub


Public Function GetMaxAmplifierOutput(ByVal ipAmplifierPhaseOptions As String, ByVal ipRunMode As AmplifierModeEnum, ByVal ipProgram As Kvp) As String

    Dim myListOfAmplifierPhaseSettings As Collection
    Set myListOfAmplifierPhaseSettings = GetListOfAmplifierPhaseSettings(ipAmplifierPhaseOptions)
    
    Dim myOutput As Kvp: Set myOutput = New Kvp
    Dim myMaxOutput As Long:    myMaxOutput = 0
    Dim myPhaseSetting As Variant
    For Each myPhaseSetting In myListOfAmplifierPhaseSettings
        
        Set myOutput = GetAmpOutput(ipProgram, myPhaseSetting, ipRunMode)
        
        If myOutput.GetLast > myMaxOutput Then
        
            myMaxOutput = myOutput.GetLast
            
        End If
        
    Next
    
    GetMaxAmplifierOutput = myMaxOutput
    
End Function

Public Function GetAmpOutput(ByVal ipProgram As Kvp, ByVal ipPhaseSequence As String, ByVal ipRunMode As AmplifierModeEnum) As Kvp

    Dim myOutputMode As OutputModeEnum
    myOutputMode = GetOutputMode(ipRunMode)
    SetupAmplifiers ipProgram, myOutputMode, ipPhaseSequence
        
    Dim myAmpInput As Kvp
    Set myAmpInput = MakeLongLongVsLongLongKvp(Array(0))
        
    Do
    
        Dim myAmp As Amplifier
        For Each myAmp In s.Amps
            ' Catch first amplifier completing
            If myAmp.RunHasFinished Then Exit Do
            myAmp.Run myAmpInput
            'catch last amplifier completing
            If myAmp.RunHasFinished Then Exit Do
            Set myAmpInput = myAmp.GetAmpOutput.Clone
            
        Next
                        
    Loop While Not ipRunMode = AmplifierModeEnum.SinglePass
    
    Set GetAmpOutput = myAmpInput
    
End Function


Public Sub AddAmps(ByVal ipAmpCount As Long)

    If s.Amps Is Nothing Then
    
        Set s.Amps = New Collection
        
    End If
    
    '@Ignore VariableNotUsed
    Dim myIndex As Long
    For myIndex = 1 To ipAmpCount
    
        s.Amps.Add New Amplifier
    
    Next
    
End Sub

Private Sub SetupAmplifiers(ByVal ipProgram As Kvp, ByVal ipOutputMode As OutputModeEnum, ByVal ipPhaseSequence As String)
    
    If s.Amps Is Nothing Then
    
        AddAmps 5
    
    End If
    
    Dim myAmp As Amplifier
    Dim myAmpPhase As Long
    myAmpPhase = 1
    For Each myAmp In s.Amps
    
        With myAmp
            
            Set .Program = ipProgram
            .OutputMode = ipOutputMode
            .Phase = Mid$(ipPhaseSequence, myAmpPhase, 1)
            myAmpPhase = myAmpPhase + 1
            
        End With
        
    Next
    
End Sub


Private Function GetOutputMode(ByVal ipRunMode As AmplifierModeEnum) As OutputModeEnum

    If ipRunMode = AmplifierModeEnum.SinglePass Then
        
        GetOutputMode = OutputModeEnum.ContinueOnOutput
        
    ElseIf ipRunMode = AmplifierModeEnum.MultiPass Then
        
        GetOutputMode = OutputModeEnum.HaltOnOutput
        
    Else
    
        Debug.Print "Unknown run mode"
        End
        
    End If
    
End Function

Public Function MakeKvp(ParamArray ipArray() As Variant) As Kvp

    Dim myKvp As Kvp: Set myKvp = New Kvp
    
    Dim myIndex As LongLong
    Dim myItem As Variant
    For Each myItem In ipArray
    
        myKvp.AddByKey myIndex, CLngLng(myItem)
        myIndex = myIndex + 1
        
    Next
        
        
    Set MakeKvp = myKvp
    
End Function







'Private Sub ClearAmplifiers()
'
'    Set s.Amps = Nothing
'
'End Sub


Public Function GetListOfAmplifierPhaseSettings(ByVal ipPhaseOptions As String) As Collection

    Dim myPhaseOptions As Variant: myPhaseOptions = Split(ipPhaseOptions, ",")
    Dim mycoll As Collection: Set mycoll = New Collection
    Dim i As Variant
    For Each i In myPhaseOptions
        
        Dim j As Variant
        For Each j In myPhaseOptions
        
            If InStr(i, j) = 0 Then
            
                Dim k As Variant
                For Each k In myPhaseOptions
                
                    If InStr(i & j, k) = 0 Then
                    
                        Dim l As Variant
                        For Each l In myPhaseOptions
                        
                            If InStr(i & j & k, l) = 0 Then
                            
                                Dim m As Variant
                                For Each m In myPhaseOptions
                                
                                    If InStr(i & j & k & l, m) = 0 Then
                                    
                                        mycoll.Add i & j & k & l & m
                                    
                                    End If
                                
                                Next
                                
                            End If
                            
                        Next
                        
                    End If
                    
                Next
                
            End If
            
        Next
        
    Next
        
    Set GetListOfAmplifierPhaseSettings = mycoll
    
End Function


Public Function GetDay7Program() As Kvp

    Dim myFso  As Scripting.FileSystemObject: Set myFso = New Scripting.FileSystemObject
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile("C:\Projects\Code\AdventOfCode\Day7\Day7Program.txt", ForReading)
    Dim myProgram As String:  myProgram = myfile.ReadAll
    myfile.Close
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.AddByIndexFromArray Split(myProgram, ",")
    Set GetDay7Program = myKvp
    
End Function





