VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Amplifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private Type Properties

    PhaseSetting                        As Long

End Type

Private p                               As Properties


Private Type State
    
    Comp                                As IntComputer
    FirstRun                            As Boolean
    
End Type

Private s                               As State

Private Sub Class_Initialize()

    Set s.Comp = New IntComputer
    
End Sub

Public Sub Run(ByVal ipInput As KvpOD)
    
    Dim myInput As KvpOD: Set myInput = New KvpOD
    
    If s.FirstRun Then
    
        myInput.AddByIndex p.PhaseSetting
        s.FirstRun = False
        
    End If
    
    myInput.AddByIndex ipInput.GetLast.Value
    s.Comp.Run myInput
    
End Sub


Public Function RunHasCompleted() As Boolean

    RunHasCompleted = s.Comp.RunHasCompleted
    
End Function


Public Property Let PhaseSetting(ByVal ipPhaseSetting As String)

    s.FirstRun = True
    p.PhaseSetting = CLng(ipPhaseSetting)
    
End Property

Public Property Get PhaseSetting() As String

    PhaseSetting = CStr(p.PhaseSetting)
    
End Property


Public Function GetAmpOutput() As KvpOD

    Set GetAmpOutput = s.Comp.GetOutput
    
End Function


Public Property Set Program(ByVal ipProgram As KvpOD)

    Set s.Comp = New IntComputer
    Set s.Comp.Program = ipProgram
    
End Property

Public Property Get Program() As KvpOD

    Set Program = s.Comp.Program
    
End Property



Public Property Let OutputMode(ByVal ipValue As OutputResponse)

    s.Comp.OutputMode = ipValue
    
End Property

Public Property Get OutputMode() As OutputResponse

    OutputMode = s.Comp.OutputMode
    
End Property

