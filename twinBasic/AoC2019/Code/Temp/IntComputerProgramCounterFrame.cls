VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntComputerProgramCounterFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const OpCode_LENGTH                            As Long = 2
Const ACCESS_MODE_LENGTH                        As Long = 1

Private Enum AccessModeEnum
    
    Invalid = -1
    Indirect = 0
    Direct = 1
    Relative = 2
    
End Enum

Public Enum OpCode

    IsNotKnown = 0
    IsAdd = 1
    IsMultiply
    IsInput
    IsOutput
    IsJumpIfTrue
    IsJumpIfFalse
    IsLessThan
    IsEquals
    IsAdjustRelativeBase
    IsRunHasCompleted = 99
    
End Enum

Public Enum FullInstructionString

    OpCodeIndex = 4
    Param1Index = 3
    Param2Index = 2
    Param3Index = 1
    
End Enum

Private Enum ProgramCounterOffset

    ProgramCounterPlus1 = 1
    ProgramCounterPlus2 = 2
    ProgramCounterPlus3 = 3
    
End Enum

Private Type Instruction

    OpCode                                     As OpCode
    AccessMode(1 To 3)                          As AccessModeEnum
    
    
End Type

Private Type Properties

    Code                                        As KvpOD
    Instruction                                 As Instruction
    ProgramCounter                              As Currency
    RelativeBase                                As Currency
    
End Type

Private p                                       As Properties
 
 
'Private Sub Class_Initialize()
'
'
'
'End Sub


Public Property Get Code() As KvpOD
    Set Code = p.Code
End Property


Public Property Set Code(ByVal ipValue As KvpOD)
    p.ProgramCounter = 0
    p.RelativeBase = 0
    Set p.Code = ipValue.Clone
End Property


Public Property Get ProgramCounter() As Currency
    ProgramCounter = p.ProgramCounter
End Property
    
Public Property Let ProgramCounter(ByVal ipValue As Currency)
    p.ProgramCounter = ipValue
End Property


Public Property Get Instruction() As Currency
    vmDecodeInstruction
    Instruction = p.Instruction.OpCode
End Property


Public Property Get Memory1() As Currency
    Memory1 = p.Code.Item(vmDecodeMemoryAccess(ProgramCounterPlus1))
End Property

Public Property Let Memory1(ByVal ipValue As Currency)
    p.Code.Item(vmDecodeMemoryAccess(ProgramCounterPlus1)) = ipValue
End Property


Public Property Get Memory2() As Currency
    Memory2 = p.Code.Item(vmDecodeMemoryAccess(ProgramCounterPlus2))
End Property

Public Property Let Memory2(ByVal ipValue As Currency)
    p.Code.Item(vmDecodeMemoryAccess(ProgramCounterPlus2)) = ipValue
End Property


Public Property Get Memory3() As Currency
    Memory3 = p.Code.Item(vmDecodeMemoryAccess(ProgramCounterPlus3))
End Property

Public Property Let Memory3(ByVal ipValue As Currency)
    p.Code.Item(vmDecodeMemoryAccess(ProgramCounterPlus3)) = ipValue
End Property


Public Property Get RelativeBase() As Currency
    RelativeBase = p.RelativeBase
End Property

Public Property Let RelativeBase(ByVal ipValue As Currency)
    p.RelativeBase = ipValue
End Property


Private Sub vmDecodeInstruction()

    Dim myFullInstruction As String
    myFullInstruction = VBA.Format$(p.Code.Item(p.ProgramCounter), "00000")
    
    With p.Instruction
    
        .OpCode = CLng(Mid$(myFullInstruction, OpCodeIndex, OpCode_LENGTH))
        
        If .OpCode = OpCode.IsRunHasCompleted Then
        
            .AccessMode(ProgramCounterPlus1) = Invalid
            .AccessMode(ProgramCounterPlus2) = Invalid
            .AccessMode(ProgramCounterPlus3) = Invalid
            Exit Sub
            
        End If
        
        .AccessMode(ProgramCounterPlus1) = CLng(Mid$(myFullInstruction, Param1Index, ACCESS_MODE_LENGTH))
        .AccessMode(ProgramCounterPlus2) = CLng(Mid$(myFullInstruction, Param2Index, ACCESS_MODE_LENGTH))
        .AccessMode(ProgramCounterPlus3) = CLng(Mid$(myFullInstruction, Param3Index, ACCESS_MODE_LENGTH))
        
    End With
      
End Sub


Private Function vmDecodeMemoryAccess(ByVal ipParameter As ProgramCounterOffset) As Currency

    Dim myParamAddress As Currency
    myParamAddress = ipParameter + p.ProgramCounter
    If p.Code.LacksKey(myParamAddress) Then vmExtendMemory myParamAddress
    
    Dim myLookup As Currency
    Select Case p.Instruction.AccessMode(ipParameter)
    
        Case Indirect:  myLookup = p.Code.Item(myParamAddress)
        Case Direct:    myLookup = myParamAddress
        Case Relative:  myLookup = p.Code.Item(myParamAddress) + p.RelativeBase
            
    End Select
    
    If p.Code.LacksKey(myLookup) Then vmExtendMemory myLookup
    vmDecodeMemoryAccess = myLookup
    
End Function


Private Sub vmExtendMemory(ByVal ipMaxAddress As Currency)

    If p.Code.HoldsKey(ipMaxAddress) Then
        
        Debug.Print "vmExtendMemory: Cannot extend within the existing memory", ipMaxAddress
        
    Else
    
        Debug.Print "vmExtendMemory: Extending from " & CStr(p.Code.Count) & " to " & CStr(ipMaxAddress)
        
    End If
    
    Dim myIndex As Currency
    For myIndex = p.Code.Count To ipMaxAddress
    
        p.Code.AddByKey myIndex, 0^
    
    Next

End Sub

