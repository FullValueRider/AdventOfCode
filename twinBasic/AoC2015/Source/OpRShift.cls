VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpRShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

' RShift
'@Ignore ProcedureNotUsed
Public Function Exec(ByVal ipLogicPair As KVPair, ByRef iopWires As Kvp) As Boolean

    If ipLogicPair.Value.GetItem(1&) <> "rshift" Then
    
        Err.Raise _
            5 + vbObjectError, _
            "Operation mismatch", _
            Layout.Fmt("Expecting 'rshift': Found '{0}'", ipLogicPair.Value.GetItem(1&))
    
    End If
    
    'ipLogicpair is Output Wire:String vs Logic:String
    Exec = False
    
    '@Ignore VariableNotAssigned
    Dim myValue1 As Long
    '@Ignore UnassignedVariableUsage
    If Not Day07.TryGetValueFromInputOrWire(ipLogicPair.Value.GetItem(0&), iopWires, myValue1) Then Exit Function
    
    Dim myShift As Long
    myShift = CLng(ipLogicPair.Value.GetItem(2&))
    
    '@Ignore UnassignedVariableUsage
    'Debug.Print ipLogicPair.Value.GetItem(1&), "rshift", myValue1, myShift, shr(myValue1, myShift), ipLogicPair.Key
    iopWires.AddByKey ipLogicPair.Key, shr(myValue1, myShift)
    
    Exec = True

End Function


Public Function shr(ByVal ipValue As Long, ByVal ipShift As Long) As Long
    'Dim i As Byte
    shr = ipValue
    If ipShift > 0 Then
        shr = Int(shr / (2 ^ ipShift))
    End If
End Function
