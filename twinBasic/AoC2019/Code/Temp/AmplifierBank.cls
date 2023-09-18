VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AmplifierBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'@PredeclaredId

Public Enum Configuration

    SinglePass = 1
    Looped = 2

End Enum


Private Type Properties

    Configuration                           As Configuration
    Output                                  As Long
    Program                                 As KvpOD
    PhaseSettings                           As String
    
End Type

Private p                                   As Properties


Private Type State

    Amplifiers                              As Collection
    
End Type

Private s                                   As State


'@DefaultMember
Public Function Debut(ByVal ipNoOfAmplifiers As Long) As AmplifierBank
Attribute Debut.VB_UserMemId = 0

    With New AmplifierBank
    
        Set Debut = Self(ipNoOfAmplifiers)
    
    End With
    
End Function


Public Function Self(ByVal ipNoOfAmplifiers As Long) As AmplifierBank


    Set s.Amplifiers = New Collection
    
    '@Ignore VariableNotUsed
    Dim myItem As Long
    For myItem = 1 To ipNoOfAmplifiers
    
        s.Amplifiers.Add New Amplifier
        
    Next
    
    Set Self = Me
        
End Function


Public Property Get Output() As Long

     Output = p.Output
     
End Property


Public Property Let PhaseSettings(ByVal Value As String)

    p.PhaseSettings = Value
    
End Property


Public Property Let Configuration(ByVal Value As Configuration)

    p.Configuration = Value
    
End Property


Public Property Set Program(ByVal Value As KvpOD)
    
    Set p.Program = Value
    
End Property


Private Function GetComputerOutputModeForAmplifierConfiguration(ByVal ipConfiguration As Configuration) As OutputResponse

    If ipConfiguration = Configuration.SinglePass Then
        
        GetComputerOutputModeForAmplifierConfiguration = OutputResponse.ContinueOnOutput
        
    Else
        
        GetComputerOutputModeForAmplifierConfiguration = OutputResponse.HaltOnOutput
        
        
    End If
    
End Function


Public Sub Run()

    Dim myItem As Long
    For myItem = 1 To s.Amplifiers.Count
    
        With s.Amplifiers.Item(myItem)
        
            Set .Program = p.Program
            .OutputMode = GetComputerOutputModeForAmplifierConfiguration(p.Configuration)
           .PhaseSetting = CLng(Mid$(p.PhaseSettings, myItem, 1))
        
        End With
        
    Next

    Dim myInput As KvpOD
    Set myInput = MakeKvp(0)
    
    Do
        For myItem = 1 To s.Amplifiers.Count
        
            Dim myAmplifier As Amplifier
            Set myAmplifier = s.Amplifiers.Item(myItem)
            myAmplifier.Run myInput
            If p.Configuration = Looped And myAmplifier.RunHasCompleted Then Exit Do
            Set myInput = myAmplifier.GetAmpOutput
            
        Next
        
    Loop While p.Configuration = Looped
    
    p.Output = myInput.GetLast.Value

End Sub
