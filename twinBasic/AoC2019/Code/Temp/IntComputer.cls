VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IntComputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("AdventOfCode")

Public Enum OutputResponse

    HaltOnOutput = 1
    ContinueOnOutput = 2

End Enum

Private Type Properties

    ActionOnOutput                                As OutputResponse
    RunHasCompleted                             As Boolean
    '@Ignore KeywordsUsedAsMember
    Input                                       As KvpOD
    Output                                      As KvpOD

End Type

Private p                                       As Properties


Private Type State
    
    NewOutputReady                              As Boolean
    Processor                                   As IntComputerProgramCounterFrame
    
End Type

Private s                                       As State

Public Property Get Program() As KvpOD
    
    Set Program = s.Processor.Code
    
End Property

Public Property Set Program(ByVal ipProgram As KvpOD)

    Set s.Processor = New IntComputerProgramCounterFrame
    Set s.Processor.Code = ipProgram
    
End Property

Public Property Let OutputMode(ByVal ipValue As OutputResponse)

    p.ActionOnOutput = ipValue

End Property

Public Property Get OutputMode() As OutputResponse

    OutputMode = p.ActionOnOutput

End Property

Public Function RunHasCompleted() As Boolean

    RunHasCompleted = p.RunHasCompleted
    
End Function

Public Function GetOutput() As KvpOD

    Set GetOutput = p.Output
    
End Function

Public Sub Run(Optional ByVal ipInput As KvpOD = Nothing)
    
    p.RunHasCompleted = False
    s.NewOutputReady = False
    SetupIO ipInput
    
    Do
        
        Select Case s.Processor.Instruction
        
            Case IsAdd:                     VmAdd
            Case IsMultiply:                VmMultiply
            Case IsInput:                   VmInput
            Case IsOutput:                  VmOutput
            Case IsJumpIfTrue:              VmJumpIfTrue
            Case IsJumpIfFalse:             VmJumpIfFalse
            Case IsLessThan:                VmLessThan
            Case IsEquals:                  VmEquals
            Case IsAdjustRelativeBase:      VmAdjustRelativeBase
            Case IsRunHasCompleted:
            
                p.RunHasCompleted = True
                Exit Sub
                
            Case Else
            
                Debug.Print "Ooops! Something went wrong Op code is " & s.Processor.Instruction
                End
                
        End Select
        
        If p.ActionOnOutput = OutputResponse.HaltOnOutput And s.NewOutputReady Then
        
            s.NewOutputReady = False
            Exit Do
            
        End If
    
    Loop
        
End Sub

Public Sub SetupIO(ByVal ipInput As KvpOD)

    Set p.Output = New KvpOD
    Set p.Input = New KvpOD
    
    If Not ipInput Is Nothing Then

        Set p.Input = ipInput.Clone
        
    End If
    
End Sub

Private Sub VmAdd()

    With s.Processor
    
        .Memory3 = .Memory1 + .Memory2
        .ProgramCounter = .ProgramCounter + 4
        
    End With
    
End Sub

Private Sub VmMultiply()

    With s.Processor
    
        .Memory3 = .Memory1 * .Memory2
        .ProgramCounter = .ProgramCounter + 4
        
    End With
    
End Sub

Private Sub VmOutput()
    
    With s.Processor
    
        p.Output.AddByIndex .Memory1
        .ProgramCounter = .ProgramCounter + 2
        
    End With
    
    s.NewOutputReady = True
    
End Sub

Private Sub VmInput()
    
    If p.Input.Count = 0 Then
                
        Debug.Print "IntComputer:vmInput: No input available "
        Err.Clear
        Err.Raise vbObjectError + 17, TypeName(Me) & ":VmInput", "The input queue is empty"
        
    End If
    
    With s.Processor
    
        .Memory1 = p.Input.GetFirst.Value.Value
        p.Input.RemoveFirst
        .ProgramCounter = .ProgramCounter + 2
        
    End With
    
End Sub

Private Sub VmJumpIfTrue()

    With s.Processor
    
        If .Memory1 <> 0 Then
        
            .ProgramCounter = .Memory2
            
        Else
        
            .ProgramCounter = .ProgramCounter + 3
        
        End If
        
    End With

End Sub

Private Sub VmJumpIfFalse()

    With s.Processor
    
        If .Memory1 = 0 Then
        
            .ProgramCounter = .Memory2
            
        Else
        
            .ProgramCounter = .ProgramCounter + 3
        
        End If
        
    End With
    
End Sub

Private Sub VmLessThan()
    
    With s.Processor
   
        .Memory3 = IIf(.Memory1 < .Memory2, 1, 0)
        .ProgramCounter = .ProgramCounter + 4
        
    End With
    
End Sub

Private Sub VmEquals()
    
    With s.Processor
   
        .Memory3 = IIf(.Memory1 = .Memory2, 1, 0)
        .ProgramCounter = .ProgramCounter + 4
        
    End With

End Sub

Private Sub VmAdjustRelativeBase()
    
    With s.Processor
    
        .RelativeBase = .Memory1 + .RelativeBase
        .ProgramCounter = .ProgramCounter + 2
        
    End With
    
End Sub

