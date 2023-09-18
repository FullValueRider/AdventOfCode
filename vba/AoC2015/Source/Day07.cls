VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule VariableNotAssigned
'@PredeclaredId
Option Explicit

' Solution
' Input represent logic gates with an output in the form X op Y -> Value
' All but two values are labels used in X or Y
' This is reorganised to a dictionary of Y vs Seq(x,op,Y)
' A second dictionary is used to track when value is a number rther than a label
' We do repeated enumerations over Value vs Seq(x,op,Y) checking if x and Y are numbers
' and if they are then x op y is calculated and value label vs value is added to the second dictionary
' The solution is found when we find a value for label 'a'
' Gotcha #1: in X op Y, op and y may not exist meaning
' seq(x, op, y) may have a count of 1,2 or 3.  There will never be just op.
' Gotcha #2: X Y and Value are unsigned 16 bit numbers


Private Const TODAY             As String = "\Day07.txt"

Private Type State

    OutputVsLogic                           As KvpC
    Resolved                                As KvpC
    
End Type

Private s                                   As State

Public Sub Execute()
    Part01
    Part02
End Sub

Private Sub Initialise()

    Dim myData As Variant
    myData = Filer.GetFileAsArrayOfStrings(AoCRawData & Year & TODAY)
    
    ' create a set of seq of x,op,t,value
    Dim myOps As seqC
    Set myOps = _
        seqC(myData) _
            .MapIt(mpMultiReplace(Array(Array(" -> ", chars.twSpace)))) _
            .MapIt(mpToLower.Deb) _
            .MapIt(mpsplit(chars.twSpace))

'
     Set s.OutputVsLogic = KvpC.Deb
     
    
    ' create a kvp of value vs seq(x,op,y)
    'Dim myItem As Variant
    Dim myItems As Iteritems: Set myItems = Iteritems(myOps)
    Do
    
        Dim myS As seqC
        Set myS = myItems.curitem(0)

        s.OutputVsLogic.Add myS.Pop, myS
        
    Loop While myItems.MoveNext
    
End Sub


Public Sub Part01()
    
    Initialise
    
    'Resolved holds output labels vs output values
    Set s.Resolved = KvpC.Deb
    
    Do Until s.Resolved.HoldsKey("a")
    
        Dim myOutputLabel As Variant
        Dim myOutputLabels As Iteritems: Set myOutputLabels = Iteritems(s.OutputVsLogic)
        Do
            myOutputLabel = myOutputLabels.curkey(0)
            ' Scan s.outputvsLogic looking for logic where the inputs have not been resolved to numbers
            ' i.e the output (key) will not be contained in s.Resolved.keys
        
            If s.Resolved.lackskey(myOutputLabel) Then
            
                Dim myOutputValue As Long
                If TryCalculateOutputValue(myOutputLabel, myOutputValue) Then
                    s.Resolved.Add myOutputLabel, myOutputValue
                End If
               
            End If
        Loop While myOutputLabels.MoveNext
    Loop
    
    Dim myResult As Long
    myResult = s.Resolved.Item("a")
    fmt.dbg "The answer to Day {0} part 1 is 46065.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult
    
End Sub


Public Sub Part02()

    Initialise
    
    Set s.Resolved = KvpC.Deb
    s.Resolved.Add "b", "46065"
    
    Do Until s.Resolved.HoldsKey("a")
    
        Dim myOutputLabel As Variant
        Dim myOutputLabels As Iteritems: Set myOutputLabels = Iteritems(s.OutputVsLogic)
        Do
            ' Scan s.outputvsLogic looking for logic where the inputs have not been resolved to numbers
            ' i.e the output (key) will not be contained in s.Resolved.keys
            myOutputLabel = myOutputLabels.curkey(0)
            If s.Resolved.lackskey(myOutputLabel) Then
            
                Dim myOutputValue As Long
                If TryCalculateOutputValue(myOutputLabel, myOutputValue) Then
                    s.Resolved.Add myOutputLabel, myOutputValue
                End If
               
            End If
        Loop While myOutputLabels.MoveNext
    Loop
    
    Dim myResult As Long
    myResult = s.Resolved.Item("a")

    fmt.dbg "The answer to Day {0} part 02 is 14134.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult

End Sub


Private Function TryCalculateOutputValue(ByRef ipOutputLabel As Variant, ByRef opOutputValue As Long) As Boolean

    Dim myLogic As seqC
    Set myLogic = s.OutputVsLogic.Item(ipOutputLabel)
    
    TryCalculateOutputValue = True
    Select Case myLogic.Count
    
        ' test the input wires to see if they have been resolved to a value
        ' only continue to evaluate the logic gate
        ' if the inputs are numbers
        Case 1:     TryCalculateOutputValue = TryResolveOneValue(myLogic, opOutputValue)
        Case 2:     TryCalculateOutputValue = TryResolveOneOpOneValue(myLogic, opOutputValue)
        Case 3:     TryCalculateOutputValue = TryResolveOneOpTwoValues(myLogic, opOutputValue)
        Case Else:  TryCalculateOutputValue = False
            
    End Select
    
End Function

Public Function TryResolveOneValue(ByRef ipLogic As seqC, ByRef opOutputValue As Long) As Boolean
    ' could be number or string
    Dim myResult As Boolean

    myResult = True
    Select Case True
        
        Case VBA.IsNumeric(ipLogic.Last):       opOutputValue = ipLogic.Last
        Case s.Resolved.HoldsKey(ipLogic.Last): opOutputValue = s.Resolved.Item(ipLogic.Last)
        Case Else:                              myResult = False
        
    End Select
    
    TryResolveOneValue = myResult
    
End Function

' op Value
Public Function TryResolveOneOpOneValue(ByRef ipLogic As seqC, ByRef opOutputValue As Long) As Boolean

    Dim myResult As Boolean
    myResult = True
    
    Select Case True
    
        Case VBA.IsNumeric(ipLogic.Last):            opOutputValue = EvaluateOpValue(ipLogic.First, ipLogic.Last)
        Case s.Resolved.HoldsKey(ipLogic.Last):      opOutputValue = EvaluateOpValue(ipLogic.First, s.Resolved.Item(ipLogic.Last))
        Case Else:                                   myResult = False
        
    End Select
    
    TryResolveOneOpOneValue = myResult
  
End Function

' value op value
Public Function TryResolveOneOpTwoValues(ByRef ipLogic As seqC, ByRef opOutputValue As Long) As Boolean

    Dim myResult As Boolean
    myResult = True
    
    Dim myValue1 As Long
    Select Case True
    
        Case VBA.IsNumeric(ipLogic.First):           myValue1 = ipLogic.First
        Case s.Resolved.HoldsKey(ipLogic.First):     myValue1 = s.Resolved.Item(ipLogic.First)
        Case Else:                                  Exit Function
        
    End Select
    
    If Not myResult Then
        TryResolveOneOpTwoValues = myResult
        Exit Function
    End If
    
    Dim myValue2 As Long
    Select Case True
    
        Case VBA.IsNumeric(ipLogic.Last):           myValue2 = ipLogic.Last
        Case s.Resolved.HoldsKey(ipLogic.Last):     myValue2 = s.Resolved.Item(ipLogic.Last)
        Case Else:                                  Exit Function
        
    End Select
    
    If Not myResult Then
        TryResolveOneOpTwoValues = myResult
        Exit Function
    End If
    
    opOutputValue = EvaluateTwoItemOp(ipLogic.Item(2), myValue1, myValue2)
   
    TryResolveOneOpTwoValues = myResult
    
End Function
    
Private Function EvaluateOpValue(ByRef ipOp As String, ByRef ipValue As Long) As Long
    ' as we don't yet have unsigned types we work with longlongs
    Dim myResult As Long
    
    Select Case ipOp
    
        Case "not"
        
            myResult = ((Not VBA.CLng(ipValue)) And &HFFFF&)
            
    End Select
    'Debug.Print ipOp, ipValue, CStr(myResult)
    EvaluateOpValue = myResult
    
End Function


Private Function EvaluateTwoItemOp(ByRef ipOp As String, ByRef ipValue1 As Long, ByRef ipvalue2 As Long) As Long

    Dim myResult As Long
    Select Case ipOp
    
        Case "and"
        
            myResult = ipValue1 And ipvalue2
            
            
        Case "or"
        
            myResult = ipValue1 Or ipvalue2
            
        
        Case "lshift"
        
            myResult = ipValue1 * (2 ^ ipvalue2)
            
        
        Case "rshift"
        
            myResult = ipValue1 \ (2 ^ ipvalue2)
            
    
    End Select
    
    'Debug.Print ipOp, ipValue1, ipvalue2, myResult, myResult And &HFFFF
    EvaluateTwoItemOp = myResult And &HFFFF
    
End Function


