VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Private Const RAW_DATA_PATH                 As String = "C:\Users\slayc\source\repos\VBA\AdventOfCode\2015\Day07RawData.txt"

Private Type State

    RawData                                 As Kvp
    OWvsIL                                  As Kvp 'OutputWireVsInputLogic
    Ops                                     As Kvp
    OperatedCircuit                         As Kvp
    
End Type

Private s                                   As State

Private Sub Class_Initialize()

    Set s.RawData = Common.GetRawDataByLine(RAW_DATA_PATH)
    Set s.OWvsIL = GetOutputWireVsLogicArray(s.RawData)
    ' Output wires are unique if no errors generated above
    
    Set s.Ops = New Kvp
    
    With s.Ops
    
        .AddByKey "and", OpAnd
        .AddByKey "or", OpOr
        .AddByKey "not", OpNot
        .AddByKey "lshift", OpLShift
        .AddByKey "rshift", OpRShift
        .AddByKey "input", OpInput ' we made this one up to simplify things
    
    End With
    
End Sub


Public Sub Part1()
    
    
    Set s.OperatedCircuit = OperateCircuit(s.OWvsIL)

    Dim myPair As KVPair
    Set myPair = s.OperatedCircuit.GetLast
    
    Layout.prt "The answer for Day07 Part 1 Should be 3176.  Found is {0} ", myPair.Value

End Sub


Public Sub Part2()

    
    Dim myPair As KVPair
    Set myPair = s.OperatedCircuit.GetLast
    
    Dim myKvp As Kvp
    Set myKvp = New Kvp
    myKvp.AddByIndexFromArray Array("input", CStr(myPair.Value))
    
    Set s.OWvsIL = GetOutputWireVsLogicArray(s.RawData)
    s.OWvsIL.SetItem "b", myKvp
    
    Set s.OperatedCircuit = OperateCircuit(s.OWvsIL)
    Set myPair = s.OperatedCircuit.GetLast
    '29568 is too high
    Layout.prt "The answer for Day07 Part 2 should be xxxx. Found is {0}", myPair.Value

End Sub


Public Function GetOutputWireVsLogicArray(ByVal ipRawdata As Kvp) As Kvp

    Dim myPair As KVPair
    Dim myOWvsIL As Kvp
    Set myOWvsIL = New Kvp
    For Each myPair In ipRawdata
    
        Dim myOutputWireVsLogicStr As Kvp
        Set myOutputWireVsLogicStr = Str.TriageToKvp(myPair.Value, "->")
        ' In the raw data, inputs may be a wire code or a number
        ' which will occur in myArray(0&) thus we add the term input
        ' when this occurs
           
        Dim myPrefix As String
        myPrefix = vbNullString
        Dim myLogic As Kvp
        Set myLogic = Str.TriageToKvp(myPrefix & myOutputWireVsLogicStr.GetFirst.Value)
        
        If myLogic.Count = 1 Then
        
            myPrefix = "input "
            
        End If
        
        myOWvsIL.AddByKey myOutputWireVsLogicStr.GetLast.Value, Str.TriageToKvp(myPrefix & myLogic.GetValuesAsString(" "))
    
    Next
    
    Set GetOutputWireVsLogicArray = myOWvsIL

End Function

Public Function OperateCircuit(ByVal ipOWvsIL As Kvp) As Kvp ' returns a Wires kvp

    Dim myWires As Kvp
    Set myWires = New Kvp
    Do While myWires.Count < ipOWvsIL.Count

        DoEvents
        
        Dim myPair As KVPair
        For Each myPair In ipOWvsIL
        
            If myWires.LacksKey(myPair.Key) Then
            
                Dim OpIndex As Long
                OpIndex = GetOpIndex(myPair)
                s.Ops.GetItem(myPair.Value.GetItem(OpIndex)).Exec myPair, myWires
                
            End If
        
        Next
    
    Loop
    
    Set OperateCircuit = myWires
    
End Function


Public Function GetOpIndex(ByVal ipPair As KVPair) As Long

   ' We cannot  use just IsSingleInput as the possible
   ' values for the first part of the logic statement
   ' can be a single letter that appears in 'not,input'
    GetOpIndex = 1
    
    If FirstInputIsWire(ipPair.Value) Then

        Exit Function

    ElseIf IsSingleInput(ipPair.Value) Then

        GetOpIndex = 0

    End If

End Function


Public Function FirstInputIsWire(ByVal ipOperands As Kvp) As Boolean
    FirstInputIsWire = Len(ipOperands.GetFirst.Value) <= 2
End Function


Public Function IsSingleInput(ByVal ipOperands As Kvp) As Boolean
    IsSingleInput = InStr("not,input", ipOperands.GetFirst.Value) > 0
End Function


Public Function TryGetValueFromInputOrWire(ByVal ipValue As String, ByVal ipWires As Kvp, ByRef opResult As Long) As Boolean
    
    ' Used by Op objects
    TryGetValueFromInputOrWire = True

    If IsNumeric(ipValue) Then
    
        opResult = CLng(ipValue)
        
    ElseIf ipWires.LacksKey(ipValue) Then
    
        TryGetValueFromInputOrWire = False
        
    Else
    
        opResult = ipWires.GetItem(ipValue)
    
    End If

End Function

