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

Public Enum RunMode

    SinglePass = 1
    MultiPass = 2

End Enum

Private Type Properties

    Phase                               As String

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

Public Sub Run(ByVal ipInput As Kvp)
    
    Dim myInput As Kvp: Set myInput = New Kvp
    
    If s.FirstRun Then
    
        Set myInput = MakeLongLongVsLongLongKvp(Array(CLngLng(p.Phase), CLngLng(ipInput.GetLast)))
        s.FirstRun = False
        
    Else
    
        Set myInput = MakeLongLongVsLongLongKvp(Array(CLngLng(ipInput.GetLast)))
    
    End If
    
    s.Comp.Run myInput
    
End Sub


Public Function RunHasFinished() As Boolean

    RunHasFinished = s.Comp.RunHasCompleted
    
End Function


Public Property Let Phase(ByVal ipPhase As String)

    s.FirstRun = True
    p.Phase = ipPhase
    
End Property

Public Property Get Phase() As String

    Phase = p.Phase
    
End Property



Public Function GetAmpOutput() As Kvp

    Set GetAmpOutput = s.Comp.GetOutput
    
End Function


Public Property Set Program(ByVal ipProgram As Kvp)

    Set s.Comp = New IntComputer
    Set s.Comp.Program = ipProgram
    
End Property

Public Property Get Program() As Kvp

    Set Program = s.Comp.Program
    
End Property



Public Property Let OutputMode(ByVal ipValue As OutputModeEnum)

    s.Comp.OutputMode = ipValue
    
End Property

Public Property Get OutputMode() As OutputModeEnum

    OutputMode = s.Comp.OutputMode
    
End Property

