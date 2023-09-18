Attribute VB_Name = "Day7"
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
    Dim myThrusterSetting As String: myThrusterSetting = GetMaxAmplifierOutput("0,1,2,3,4", AmplifierModeEnum.SinglePass, myProgram)
    Debug.Print myThrusterSetting
    
End Sub


Public Sub Day7Part2()

    Dim myProgram As Kvp: Set myProgram = GetDay7Program
    Dim myThrusterSetting As String: myThrusterSetting = GetMaxAmplifierOutput("5,6,7,8,9", AmplifierModeEnum.MultiPass, myProgram)
    Debug.Print myThrusterSetting
    
End Sub


Public Function GetMaxAmplifierOutput(ByVal ipAmplifierPhaseOptions As String, ByVal ipRunMode As AmplifierModeEnum, ByVal ipProgram As Variant) As String

    Dim myListOfAmplifierPhaseSettings As Collection: Set myListOfAmplifierPhaseSettings = GetListOfAmplifierPhaseSettings(ipAmplifierPhaseOptions)
    
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

Public Function GetAmpOutput(ByVal ipProgram As Variant, ByVal ipPhaseSequence As String, ByVal ipRunMode As AmplifierModeEnum) As Kvp

    Dim myOutputMode As OutputModeEnum: myOutputMode = GetOutputMode(ipRunMode)
    SetupAmplifiers ipProgram, myOutputMode, ipPhaseSequence
        
    Dim myAmpOutput As Kvp:
    Set myAmpOutput = New Kvp
    myAmpOutput.AddByIndex "0"
    Dim myAmp As Amplifier
    
    Do
    
        For Each myAmp In s.Amps
        
            
            Set myAmpOutput = myAmp.Run(myAmpOutput).Clone
            Debug.Print myAmpOutput.GetValuesAsString
            Debug.Print myAmp.RunHasFinished
            If (ipRunMode = AmplifierModeEnum.MultiPass) And myAmp.RunHasFinished Then Exit For
            
        Next

        Debug.Print myAmpOutput.GetLast
        If ipRunMode = AmplifierModeEnum.SinglePass Then
    
            Set GetAmpOutput = myAmpOutput
            ClearAmplifiers
            Exit Function
        
        End If
        
    Loop
    
    Set GetAmpOutput = myAmpOutput
    ClearAmplifiers
    
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

Public Sub SetupAmplifiers(ByVal ipProgram As Kvp, ByVal ipOutputMode As OutputModeEnum, ByVal ipPhaseSequence As String)
    
    If s.Amps Is Nothing Then
    
        AddAmps 5
    
    End If
    
    Dim myAmp As Amplifier
    Dim myAmpPhase As Long: myAmpPhase = 1
    For Each myAmp In s.Amps
    
        With myAmp
            
            Set .Program = ipProgram.Clone
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
    
    myKvp.AddByIndexFromArray ipArray
    Set MakeKvp = myKvp
    
End Function
Private Sub TestComputer()

Dim myComp As IntComputer
 
    Debug.Print "Computer Tests"
    Set myComp = New IntComputer
    Set myComp.Program = MakeKvp(3, 9, 8, 9, 10, 9, 4, 9, 99, -1, 8)
    
    Debug.Print "Test 1", 1 - CLng(myComp.Run(MakeKvp(8)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 9, 8, 9, 10, 9, 4, 9, 99, -1, 8)
    Debug.Print "Test 2", 0 - CLng(myComp.Run(MakeKvp(-1)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 9, 7, 9, 10, 9, 4, 9, 99, -1, 8)
    Debug.Print "Test 3", 1 - CLng(myComp.Run(MakeKvp(0)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 9, 7, 9, 10, 9, 4, 9, 99, -1, 8)
    Debug.Print "Test 4", 0 - CLng(myComp.Run(MakeKvp(8)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 3, 1108, -1, 8, 3, 4, 3, 99)
    Debug.Print "Test 5", 0 - CLng(myComp.Run(MakeKvp(7)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 3, 1108, -1, 8, 3, 4, 3, 99)
    Debug.Print "Test 6", 1 - CLng(myComp.Run(MakeKvp(8)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 3, 1107, -1, 8, 3, 4, 3, 99)
    Debug.Print "Test 7", 1 - CLng(myComp.Run(MakeKvp(-1)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 3, 1107, -1, 8, 3, 4, 3, 99)
    Debug.Print "Test 8", 0 - CLng(myComp.Run(MakeKvp(9)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 12, 6, 12, 15, 1, 13, 14, 13, 4, 13, 99, -1, 0, 1, 9)
    Debug.Print "Test 9", 0 - CLng(myComp.Run(MakeKvp(0)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 12, 6, 12, 15, 1, 13, 14, 13, 4, 13, 99, -1, 0, 1, 9)
    Debug.Print "Test 10", 1 - CLng(myComp.Run(MakeKvp(9)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 3, 1105, -1, 9, 1101, 0, 0, 12, 4, 12, 99, 1)
    Debug.Print "Test 11", 0 - CLng(myComp.Run(MakeKvp(0)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 3, 1105, -1, 9, 1101, 0, 0, 12, 4, 12, 99, 1)
    Debug.Print "Test 12", 1 - CLng(myComp.Run(MakeKvp(9)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 21, 1008, 21, 8, 20, 1005, 20, 22, 107, 8, 21, 20, 1006, 20, 31, 1106, 0, 36, 98, 0, 0, 1002, 21, 125, 20, 4, 20, 1105, 1, 46, 104, 999, 1105, 1, 46, 1101, 1000, 1, 20, 4, 20, 1105, 1, 46, 98, 99)
    Debug.Print "Test 13", 999 - CLng(myComp.Run(MakeKvp(0)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 21, 1008, 21, 8, 20, 1005, 20, 22, 107, 8, 21, 20, 1006, 20, 31, 1106, 0, 36, 98, 0, 0, 1002, 21, 125, 20, 4, 20, 1105, 1, 46, 104, 999, 1105, 1, 46, 1101, 1000, 1, 20, 4, 20, 1105, 1, 46, 98, 99)
    Debug.Print "Test 14", 1000 - CLng(myComp.Run(MakeKvp(8)).GetLast)
    
    Set myComp.Program = MakeKvp(3, 21, 1008, 21, 8, 20, 1005, 20, 22, 107, 8, 21, 20, 1006, 20, 31, 1106, 0, 36, 98, 0, 0, 1002, 21, 125, 20, 4, 20, 1105, 1, 46, 104, 999, 1105, 1, 46, 1101, 1000, 1, 20, 4, 20, 1105, 1, 46, 98, 99)
    Debug.Print "Test 15", 1001 - CLng(myComp.Run(MakeKvp(1080)).GetLast)
        
End Sub


Public Sub TestGetAmpOutputSingleMode()

    Debug.Print "Amplifier Tests"
    Debug.Print "Single Pass mode", CLng("43210") - CLng(GetAmpOutput(MakeKvp(3, 15, 3, 16, 1002, 16, 10, 16, 1, 16, 15, 15, 4, 15, 99, 0, 0), "43210", AmplifierModeEnum.SinglePass).GetLast)
    Debug.Print "Single Pass mode", CLng("54321") - CLng(GetAmpOutput(MakeKvp(3, 23, 3, 24, 1002, 24, 10, 24, 1002, 23, -1, 23, 101, 5, 23, 23, 1, 24, 23, 23, 4, 23, 99, 0, 0), "01234", AmplifierModeEnum.SinglePass).GetLast)
    Debug.Print "Single Pass mode", CLng("65210") - CLng(GetAmpOutput(MakeKvp(3, 31, 3, 32, 1002, 32, 10, 32, 1001, 31, -2, 31, 1007, 31, 0, 33, 1002, 33, 7, 33, 1, 33, 31, 31, 1, 32, 31, 31, 4, 31, 99, 0, 0, 0), "10432", AmplifierModeEnum.SinglePass).GetLast)
    
End Sub


Public Sub TestGetAmoutputRepeatMode()
    
    Debug.Print "Amplifier Tests"
    Debug.Print "Loop mode", CLng("139629729") - CLng(GetAmpOutput(MakeKvp(3, 26, 1001, 26, -4, 26, 3, 27, 1002, 27, 2, 27, 1, 27, 26, 27, 4, 27, 1001, 28, -1, 28, 1005, 28, 6, 99, 0, 0, 5), "98765", AmplifierModeEnum.MultiPass).GetLast)
    Debug.Print "Loop Mode", CLng("18216") - CLng(GetAmpOutput(MakeKvp(3, 52, 1001, 52, -5, 52, 3, 53, 1, 52, 56, 54, 1007, 54, 5, 55, 1005, 55, 26, 1001, 54, -5, 54, 1105, 1, 12, 1, 53, 54, 53, 1008, 54, 0, 55, 1001, 55, 1, 55, 2, 53, 55, 53, 4, 53, 1001, 56, -1, 56, 1005, 56, 6, 99, 0, 0, 0, 0, 10), "97856", AmplifierModeEnum.MultiPass).GetLast)

End Sub


Private Sub ClearAmplifiers()

    Set s.Amps = Nothing

End Sub


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
    Set myfile = myFso.OpenTextFile("C:\Projects\Code\AdventOfCode\Day7\Program.txt", ForReading)
    Dim myProgram As String:  myProgram = myfile.ReadAll
    myfile.Close
    Dim myKvp As Kvp: Set myKvp = New Kvp
    myKvp.AddByIndexFromArray Split(myProgram, ",")
    Set GetDay7Program = myKvp
    
End Function





