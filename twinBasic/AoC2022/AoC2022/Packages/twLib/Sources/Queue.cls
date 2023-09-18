VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit
'@PredeclaredId
'@Exposed


Private Type state
    
    Coll                                        As collection
    
End Type

Private s                                       As state

Public Function Deb() As Queue
    With New Queue
        Set Deb = .ConstructInstance
    End With
End Function

Friend Function ConstructInstance() As Queue
    Set s.Coll = New Collection
    Set ConstructInstance = Me
End Function

Public Function Enqueue(ByRef ipItem As Variant) As Queue
    s.Coll.Add ipItem
    Set Enqueue = Me
End Function

Public Function Dequeue() As Variant
    Dim myItem As Variant
    If VBA.IsObject(s.Coll.Item(1)) Then
        Set myItem = s.Coll.Item(1)
    Else
        myItem = s.Coll.Item(1)
    End If
    s.Coll.Remove 1
    Return myItem
End Function

Public Function ToArray() As Variant
    Dim myArray As Variant
    ReDim myArray(0 To s.Coll.Count - 1)
    
    Dim myIndex As Long
    For myIndex = 1 To s.Coll.Count
        If VBA.IsObject(s.Coll.Item(myIndex)) Then
            Set myArray(myIndex - 1) = s.Coll.Item(myIndex)
        Else
            myArray(myIndex - 1) = s.Coll.Item(myIndex)
        End If
    Next
    Return myArray
End Function

Public Function Count() As Long
    Count = s.Coll.Count
End Function


' Friend Function EnQueueItems(ParamArray ipParamArray() As Variant) As Queue
    
    
'     If BailOut.When(ArrayInfo.IsNotQueryable(ipParamArray), ParamArrayIsNotAllocated) Then
'         Return Me
'     End If
        
'     Dim myParamArray As Variant = Meta(CVar(ipParamArray))
    
    
'     Dim myItem As Variant
'     Dim myParam As IterItems = IterItems.Deb(myParamArray)
'     Do
'         Me.Enqueue myParam.GetItem
'     Loop While myParam.MoveNext
    
'     Set EnQueueItems = Me
    
' End Function