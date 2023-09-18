VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpLShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

' LShift
'@Ignore ProcedureNotUsed
Public Function Exec(ByVal ipLogicPair As KVPair, ByVal iopWires As Kvp) As Boolean

    If ipLogicPair.Value.GetItem(1&) <> "lshift" Then
    
        Err.Raise _
            5 + vbObjectError, _
            "Operation mismatch", _
            Layout.Fmt("Expecting 'input': Found '{0}'", ipLogicPair.Value.GetItem(1&))
    
    End If
    
    'ipLogicPair is Output Wire:String vs Logic:variant containing array
    Exec = False
    
    '@Ignore VariableNotAssigned
    Dim myValue1 As Long
    '@Ignore UnassignedVariableUsage
    If Not Day07.TryGetValueFromInputOrWire(ipLogicPair.Value.GetItem(0&), iopWires, myValue1) Then Exit Function
    
    Dim myShift As Long
    myShift = CLng(ipLogicPair.Value.GetItem(2&))
    
    '@Ignore UnassignedVariableUsage
    'Debug.Print ipLogicPair.Value.GetItem(1&), "lshift", myValue1, myShift, Shl(myValue1, myShift), ipLogicPair.Key
    iopWires.AddByKey ipLogicPair.Key, Shl(myValue1, myShift)
    
    Exec = True

End Function


Public Function Shl(ByVal ipValue As Long, ByVal ipShift As Long) As Long
    Shl = ipValue
    If ipShift > 0 Then
        Dim i As Long
        Dim m As Long
        For i = 1 To ipShift
            m = Shl And &H40000000
            Shl = (Shl And &H3FFFFFFF) * 2
            If m <> 0 Then
                Shl = Shl Or &H80000000
            End If
        Next i
    End If
End Function

