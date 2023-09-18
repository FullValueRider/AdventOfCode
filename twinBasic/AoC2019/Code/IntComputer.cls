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

Private Enum AccessTypeEnum

    Indirect = 0
    Direct = 1
    Relative = 2
    
End Enum


Private Enum OffsetEnum

    p1 = 1
    P2 = 2
    P3 = 3
    
End Enum


Private Type Properties

    HaltIfOutput                                As Boolean
    RunHasCompleted                             As Boolean
    Code                                        As Kvp
    '@Ignore KeywordsUsedAsMember
    Input                                       As Kvp
    Output                                      As Kvp

End Type

Private p                                       As Properties


Private Type State

    Operand                                     As LongLong
    AccessMode(1 To 3)                          As LongLong
    ProgramCounter                              As LongLong
    NewOutputReady                              As Boolean
    LocationP1                                  As LongLong
    LocationP2                                  As LongLong
    LocationP3                                  As LongLong
    RelativeBase                                As LongLong
    
End Type

Private s                                       As State


Public Property Set Program(ByRef ipProgram As Kvp)
    
    If p.Code Is Nothing Then Set p.Code = New Kvp
    Set p.Code = ipProgram.Clone
    s.ProgramCounter = 0
    s.RelativeBase = 0
    'Debug.Print p.Code.GetValuesAsString
    
End Property


Public Property Get Program() As Kvp
    
    Set Program = p.Code
    
    
End Property


Public Property Let OutputMode(ByVal ipValue As OutputModeEnum)

    p.HaltIfOutput = ipValue

End Property

Public Property Get OutputMode() As OutputModeEnum

    OutputMode = p.HaltIfOutput

End Property


Public Function RunHasCompleted() As Boolean

    RunHasCompleted = p.RunHasCompleted
    
End Function


Public Function GetOutput() As Kvp

    Set GetOutput = p.Output
    
End Function

Public Sub Run(Optional ByVal ipInput As Kvp = Nothing)
    
    p.RunHasCompleted = False
    s.NewOutputReady = False
    
    'Debug.Print "COmputer recieved " & ipInput.Count & " input items"
    SetupIO ipInput
    
    Do
    
        vmDecodeInstruction
        
        Select Case s.Operand
        
            Case 1:     opAdd
                
            Case 2:     opMultiply
                
            Case 3:     opInput
                
            Case 4:     opOutput
            
            Case 5:     opJumpIfTrue
            
            Case 6:     opJumpIfFalse
            
            Case 7:     opLessThan
            
            Case 8:     opEquals
            
            Case 9:     opAdjustRelativeBase
            
            Case 99:    p.RunHasCompleted = True
                        Exit Do
                        
            '@Ignore UnreachableCase
            Case Else
            
                Debug.Print "Ooops! Something went wrong Op code is " & p.Code.Item(CLngLng(s.ProgramCounter))
                End
                
        End Select
        
        If p.HaltIfOutput And s.NewOutputReady Then
        
            s.NewOutputReady = False
            Exit Do
            
        End If
    
    Loop
        
End Sub
Public Sub SetupIO(ByVal ipInput As Kvp)

    Set p.Output = New Kvp
    Set p.Input = New Kvp
    
    If Not ipInput Is Nothing Then

        Set p.Input = ipInput
        'Debug.Print "Input is ", p.Input.GetValuesAsString

    End If
    
    
    
End Sub


Private Sub vmDecodeInstruction()

    Dim myFullInstruction As String
    myFullInstruction = vmGetInstructionAsString
    s.Operand = CLngLng(Mid$(myFullInstruction, 4))
    s.AccessMode(P3) = CLngLng(Mid$(myFullInstruction, 1, 1))
    s.AccessMode(P2) = CLngLng(Mid$(myFullInstruction, 2, 1))
    s.AccessMode(p1) = CLngLng(Mid$(myFullInstruction, 3, 1))
    s.LocationP1 = s.ProgramCounter + OffsetEnum.p1
    s.LocationP2 = s.ProgramCounter + OffsetEnum.P2
    s.LocationP3 = s.ProgramCounter + OffsetEnum.P3
    'Debug.Print p.Code.GetItem(255&)
    'Debug.Print p.Code.GetItem(6&)
    
End Sub

Private Function vmGetInstructionAsString() As String

Const INSTRUCTION_LENGTH                    As Long = 5

    Dim myPartialInstruction As String
    'Debug.Print p.Code.GetKeysAsString
    myPartialInstruction = CStr(p.Code.Item(CLngLng(s.ProgramCounter)))
    vmGetInstructionAsString = String$(INSTRUCTION_LENGTH - Len(myPartialInstruction), "0") & myPartialInstruction
    
End Function

Private Sub vmStore(ByVal ipLocation As LongLong, ByVal ipAccessMode As LongLong, ByVal ipValue As LongLong)

    Select Case ipAccessMode
    
        Case Indirect 'Position AccessMode
            
            ' myObj is a fudge whilst I sort out why I can't assign to a Kvp
            ' once it has been passed to another method
            
            Dim myIndirectLocation As LongLong: myIndirectLocation = CLngLng(p.Code.GetItem(CLngLng(ipLocation)))
            If p.Code.LacksKey(myIndirectLocation) Then vmExtendMemory myIndirectLocation
            p.Code.SetItem CLngLng(myIndirectLocation), CLngLng(ipValue)
        
        Case Direct 'Immediate AccessMode
        
            Debug.Print "vmStore: Illegal immediate AccessMode"
            End
            
        Case Relative
        
            Dim myRelativeLocation As LongLong: myRelativeLocation = CLngLng(p.Code.GetItem(CLngLng(ipLocation)) + s.RelativeBase)
            If p.Code.LacksKey(myRelativeLocation) Then vmExtendMemory myRelativeLocation
            p.Code.SetItem CLngLng(myRelativeLocation), CLngLng(ipValue)
            
        Case Else
        
            Debug.Print "vmStore: Unknown AccessMode " & ipAccessMode
            End
            
    End Select
    
End Sub


Private Function vmLoad(ByVal ipLocation As LongLong, ByVal ipAccessMode As LongLong) As LongLong
    
    Select Case ipAccessMode
    
        Case Indirect 'Position AccessMode
        
            Dim myIndirectLocation As LongLong: myIndirectLocation = CLngLng(p.Code.GetItem(CLngLng(ipLocation)))
            If p.Code.LacksKey(myIndirectLocation) Then vmExtendMemory myIndirectLocation
            'Debug.Print p.Code.Count
            vmLoad = CLngLng(p.Code.GetItem(CLngLng(myIndirectLocation)))
        
        Case Direct 'Immediate AccessMode
        
            If p.Code.LacksKey(ipLocation) Then vmExtendMemory ipLocation
            vmLoad = CLngLng(p.Code.GetItem(CLngLng(ipLocation)))
            
        Case Relative
        
            Dim myRelativeLocation As LongLong: myRelativeLocation = CLngLng(p.Code.GetItem(CLngLng(ipLocation)) + s.RelativeBase)
            If p.Code.LacksKey(myRelativeLocation) Then vmExtendMemory myRelativeLocation
            vmLoad = CLngLng(p.Code.GetItem(CLngLng(myRelativeLocation)))
            
        '@Ignore UnreachableCase
        Case Else
        
            Debug.Print "vmLoad: Unknown AccessMode " & ipAccessMode
            End
            
    End Select
    
End Function

Private Sub opAdd()

    Dim myParam1 As LongLong: myParam1 = vmLoad(s.LocationP1, s.AccessMode(p1))
    Dim myParam2 As LongLong: myParam2 = vmLoad(s.LocationP2, s.AccessMode(P2))
    Dim myresult As LongLong: myresult = myParam1 + myParam2
    'Debug.Print s.LocationP3, CLngLng(s.AccessMode(P3)), myresult
    vmStore s.LocationP3, CLngLng(s.AccessMode(P3)), myresult
    vmMoveProgramCounter 4&
    
End Sub


Private Sub opMultiply()

    Dim myParam1 As LongLong: myParam1 = vmLoad(s.LocationP1, s.AccessMode(p1))
    Dim myParam2 As LongLong: myParam2 = vmLoad(s.LocationP2, s.AccessMode(P2))
    Dim myresult As LongLong: myresult = myParam1 * myParam2
    vmStore s.LocationP3, s.AccessMode(P3), myresult
    vmMoveProgramCounter 4&

End Sub


Private Sub opOutput()
    
    p.Output.AddByIndex vmLoad(s.LocationP1, s.AccessMode(p1))
    vmMoveProgramCounter 2&
    s.NewOutputReady = True
    'Debug.Print "One itemstored for output", p.Output.GetValuesAsString
End Sub


Private Sub opInput()
    
    If p.Input.Count = 0 Then
                
        Debug.Print "opInput: No input available "
        End
    End If
    'Debug.Print s.AccessMode(p1)
    'Debug.Print p.Input.GetFirstKey
    Dim myCount As Long: myCount = p.Input.Count
    vmStore s.LocationP1, s.AccessMode(p1), CLngLng(p.Input.PullFirst)
    If p.Input.Count >= myCount Then
    
        Debug.Print "item not removed from input"
        End
    End If
    vmMoveProgramCounter 2&
    
End Sub


Private Sub opJumpIfTrue()

    If vmLoad(s.LocationP1, s.AccessMode(p1)) <> 0 Then
    
        vmMoveProgramCounter Absolute:=vmLoad(s.LocationP2, s.AccessMode(P2))
        
    Else
    
        vmMoveProgramCounter 3&
    
    End If

End Sub


Private Sub opJumpIfFalse()

    If vmLoad(s.LocationP1, s.AccessMode(p1)) = 0 Then
    
        vmMoveProgramCounter Absolute:=vmLoad(s.LocationP2, s.AccessMode(P2))
        
    Else
    
        vmMoveProgramCounter 3&
    
    End If
    
End Sub


Private Sub opLessThan()

    Dim myParam1 As LongLong:   myParam1 = vmLoad(s.LocationP1, s.AccessMode(p1))
    Dim myParam2 As LongLong:   myParam2 = vmLoad(s.LocationP2, s.AccessMode(P2))
    Dim myValue As LongLong:    myValue = IIf(myParam1 < myParam2, 1&, 0&)
    vmStore s.LocationP3, s.AccessMode(P3), myValue
    vmMoveProgramCounter 4&
    
End Sub


Private Sub opEquals()

    Dim myParam1 As LongLong:   myParam1 = vmLoad(s.LocationP1, s.AccessMode(p1))
    Dim myParam2 As LongLong:   myParam2 = vmLoad(s.LocationP2, s.AccessMode(P2))
    Dim myValue As LongLong:    myValue = IIf(myParam1 = myParam2, 1&, 0&)
    vmStore s.LocationP3, s.AccessMode(P3), myValue
    vmMoveProgramCounter 4&

End Sub

Private Sub opAdjustRelativeBase()
    
    s.RelativeBase = vmLoad(s.LocationP1, s.AccessMode(p1)) + s.RelativeBase
    vmMoveProgramCounter 2&
    
End Sub

Private Sub vmMoveProgramCounter(Optional ByVal ipCount As LongLong = 0, Optional ByVal Absolute As LongLong = -1)

    If Absolute > -1 Then
    
        If ipCount <> 0& Then
        
            Debug.Print "vmMoveProgramCounter: Illegal move to absolute location"
        
        End If
        
        s.ProgramCounter = Absolute
        
    Else
    
        s.ProgramCounter = s.ProgramCounter + ipCount
        
    End If
    
End Sub

Private Sub vmExtendMemory(ByVal ipMaxLocation As LongLong)


    If p.Code.HoldsKey(ipMaxLocation) Then
        
        Debug.Print "vmExtendMemory: Cannot extend within the existing memory", ipMaxLocation
        
    Else
    
        Debug.Print "vmExtendMemory: Extending from " & CStr(p.Code.Count) & " to " & CStr(ipMaxLocation)
        
    End If
    '@Ignore VariableNotUsed
    Dim myIndex As LongLong
'    Dim myarray As Variant
'    myarray = p.Code.GetKeys
'
    For myIndex = p.Code.Count To ipMaxLocation
    
        p.Code.AddByKey myIndex, 0^
    
    Next

End Sub
