VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

' Input
'@Ignore ProcedureNotUsed
Public Function Exec(ByVal ipLogicPair As KVPair, ByRef iopWires As Kvp) As Boolean

    If ipLogicPair.Value.GetItem(0&) <> "input" Then
    
        Err.Raise _
            5 + vbObjectError, _
            "Operation mismatch", _
            Layout.Fmt("Expecting 'input': Found '{0}'", ipLogicPair.Value.GetItem(0&))
    
    End If

    'ipLogicpair is Output Wire:String vs Logic:Variant containing array
    Exec = False
    
    '@Ignore VariableNotAssigned
    Dim myValue1 As Long
    '@Ignore UnassignedVariableUsage
    
    If Not Day07.TryGetValueFromInputOrWire(ipLogicPair.Value.GetItem(1&), iopWires, myValue1) Then Exit Function
    
    
    'myValue1 = CLng(ipLogicPair.Value.GetItem(1&))
    
    '@Ignore UnassignedVariableUsage
    'Debug.Print ipLogicPair.Value.GetItem(0&), "input", myValue1, ipLogicPair.Key
    iopWires.AddByKey ipLogicPair.Key, myValue1
    
    Exec = True
    
End Function


