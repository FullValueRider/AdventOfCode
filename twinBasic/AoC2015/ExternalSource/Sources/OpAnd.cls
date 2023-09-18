VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpAnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

' And
'@Ignore ProcedureNotUsed
Public Function Exec(ByVal ipLogicPair As KVPair, ByRef iopWires As Kvp) As Boolean

    If ipLogicPair.Value.GetItem(1&) <> "and" Then
    
        Err.Raise _
            5 + vbObjectError, _
            "Operation mismatch", _
            Layout.Fmt("Expecting 'and': Found '{0}'", ipLogicPair.Value.GetItem(1&))
    
    End If

    'ipLogicpair is Output Wire:String vs Logic:Variant containing array
    Exec = False
    
    '@Ignore VariableNotAssigned
    Dim myValue1 As Long
    '@Ignore UnassignedVariableUsage
    If Not Day07.TryGetValueFromInputOrWire(ipLogicPair.Value.GetItem(0&), iopWires, myValue1) Then Exit Function
    
    '@Ignore VariableNotAssigned
    Dim myValue2 As Long
    '@Ignore UnassignedVariableUsage
    If Not Day07.TryGetValueFromInputOrWire(ipLogicPair.Value.GetItem(2&), iopWires, myValue2) Then Exit Function
    
    '@Ignore UnassignedVariableUsage
    '@Ignore UnassignedVariableUsage
    'Debug.Print ipLogicPair.Value.GetItem(1&), "and", myValue1, myValue2, myValue1 And myValue2, ipLogicPair.Key
    iopWires.AddByKey ipLogicPair.Key, myValue1 And myValue2
        
    Exec = True

End Function


