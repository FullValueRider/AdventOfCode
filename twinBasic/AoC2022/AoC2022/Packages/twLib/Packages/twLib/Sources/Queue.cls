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
    
    AL                                      As Seq
    
End Type

Private s                                   As state

Public Function Deb() As Queue
    With New Queue
        Set Deb = .ConstructInstance
    End With
End Function

Friend Function ConstructInstance() As Queue
    Set s.AL = Seq.Deb
    Set ConstructInstance = Me
End Function

Public Function Enqueue(ByRef ipItem As Variant) As Queue
    s.AL.Add ipItem
    Set Enqueue = Me
End Function

Public Function Dequeue() As Variant
    Sys.Assign Dequeue, s.AL.Item(s.AL.Count)
    s.AL.Remove s.AL.Count
End Function

Public Function ToArray() As Variant
    Return s.AL.Items
   
End Function

Public Function Count() As Long
    Count = s.AL.Count
End Function

' Public Function EnqueueItems(ParamArray ipParamarray() As Variant) As Queue
'     Set EnqueueItems = EnQueueItemsByFPA(CVar(ipParamarray))
' End Function

Friend Function EnQueueItems(ParamArray ipParamArray() As Variant) As Queue
    
    
    If BailOut.When(ArrayInfo.IsNotQueryable(ipParamArray), ParamArrayIsNotAllocated) Then
        Return Me
    End If
        
    Dim myParamArray As Variant = Meta(CVar(ipParamArray))
    
    
    Dim myItem As Variant
    Dim myParam As IterItems = IterItems.Deb(myParamArray)
    Do
        Me.Enqueue myParam.Item
    Loop While myParam.MoveNext
    
    Set EnQueueItems = Me
    
End Function