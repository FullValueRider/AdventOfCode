VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ranges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

'supertceded by seq.deb.repeat
' Public Function GetNumberSeries(ByVal ipStart As Variant, ByVal ipCount As Long, Optional ByVal ipStep As Variant = 1) As Variant

'     Dim myArray As Variant
'     myArray = Empty
'     GetNumberSeries = myArray
        
'     If TypeInfo.IsNotNumber(ipStart, ipCount, ipStep) Then Exit Function
    
'     ReDim myArray(0 To ipCount - 1)
'     Dim myIndex As Long
'     For myIndex = 0 To ipCount - 1
        
'         myArray(myIndex) = ipStart
'         ipStart = ipStart + ipStep
            
'     Next
    
'     GetNumberSeries = myArray
        
' End Function

' '@Description("Return an array of length ipCount populated with ipItem")
' Public Function GetItemSeries(ByVal ipCount As Long, ByVal ipItem As Variant) As Variant
'     Return Seq.Deb.Repeat(ipItem, ipCount).ToArray
' End Function


Public Function IsInRange(ByVal ipTestItem As Variant, ByVal ipLowerBound As Variant, ByVal ipUpperBound As Variant) As Boolean
    
    If BailOut.When(TypeInfo.IsNotNumber(ipTestItem), IsNotNumber) Then Return False
    If BailOut.When(TypeInfo.IsNotNumber(ipLowerBound), IsNotNumber) Then Return False
    If BailOut.When(TypeInfo.IsNotNumber(ipUpperBound), IsNotNumber) Then Return False
    
    If (ipTestItem < ipLowerBound) Or (ipTestItem > ipUpperBound) Then Return False
    Return True
End Function


Public Function IsNotInRange(ByVal ipTestItem As Variant, ByVal ipLowerBound As Variant, ByVal ipUpperBound As Variant) As Boolean
    IsNotInRange = False
    If BailOut.When(TypeInfo.IsNotNumber(ipLowerBound, ipUpperBound), IsNotNumber) Then Exit Function
    If (ipTestItem >= ipLowerBound) And (ipTestItem <= ipUpperBound) Then Exit Function
    IsNotInRange = True
End Function


' '@Description("Takes any iterable or string. The startindex and endindex should be specified using a 1 based index. The result.items lyst returns the startindex based on the actual indexing used by the iterable")
'  Friend Function TryStartRunFromAnyStartRunEnd _
' ( _
'     ByVal ipItem As Variant, _
'     Optional ByVal ipStartIndex As Variant, _
'     Optional ByVal ipRun As Variant, _
'     Optional ByVal ipEndIndex As Variant, _
'     Optional ByVal ipRank As PresetOrd = PresetOrd.ArrayFirstRank _
' ) As Boolean
' Attribute TryStartRunFromAnyStartRunEnd.VB_Description = "Takes any iterable or string. The startindex and endindex should be specified using a 1 based index. The result.items lyst returns the startindex based on the actual indexing used by the iterable"
    
'     ' Item must be an entity that has an extent, i.e. an iterable or a string
'     ' ipStartIndex , ipEndIndex and ipRun may be negative
'     '
'     ' There are 7 scenarios
'     '   1. No params
'     '   2. Start only
'     '   3. Run only
'     '   4. End only
'     '   5. Start and Run, Start and Run and End (end is ignored when start and run are provided)
'     '   6. Start and End
'     '   7. Run and End
'     '
'     ' Where coercion is used a message is retuned about the coercion applied
'     '
'     '  To simplify calculations the first and last indexes of each rank of Item are normalised
'     ' such that the first index is 1

'     Res.Reset 'NotOkay
'     TryStartRunFromAnyStartRunEnd = False

'     Dim myItem As Box
'     Set myItem = Box(ipItem)

'     Dim myItemExtents() As Rank
'     ReDim myItemExtents(1 To myItem.Extent.RankCount)
   
'     Dim myRank As Long
'     For myRank = 1 To myItem.Extent.RankCount

'         Set myItemExtents(myRank) = Rank(1, myItem.Extent(myRank).LastIndex - myItem.Extent(myRank).FirstIndex + 1)


'     Next

'     ' Now focus on the specific rank where myRankIndex is an index into myItemExtents
'     Dim myRankIndex As Long
'     If ipRank < 1 Then
    
'         myRankIndex = enums.Preset.Value(ArrayFirstRank)
'         Res.SetOkay MsgOrd.RankOfZeroResetToOne

'     Else

'         myRankIndex = ipRank

'     End If


'     ''''**** got to here
'     Dim myInputExtent As Rank
    
    
    
'     If Not VBA.IsMissing(ipRun) Then

'         If BailOut.When(ipRun = 0, RunLengthIsZero) Then Exit Function


'     End If
        
'     ' Ensure a legal start index
'     If Not VBA.IsMissing(ipStartIndex) Then
            
'         If BailOut.When(VBA.Abs(ipStartIndex) > myItem.Extent(myRank).Count, StartIndexExceedsItemCount) Then Exit Function
                
'         ' This is an intervention because we define start runs as starting from 1 or -1
'         ' so zero is not an allowed value but is taken to mean 1
'         If ipStartIndex = 0 Then
                
'             ipStartIndex = 1
'             '@Ignore FunctionReturnValueDiscarded
'             Res.SetOkay StartIndexWasZeroResetToOne
                
            
'         End If
            
'     End If
        
'     ' Ensure a legal end index
'     If Not VBA.IsMissing(ipEndIndex) Then
            
'         If VBA.Abs(ipEndIndex) > myItem.Extent(myRank).LastIndex Then
'             ' intervention to reset end index
'             ipEndIndex = myItem.Extent(myRank).LastIndex * VBA.Sgn(ipEndIndex)
'             '@Ignore FunctionReturnValueDiscarded
'             Res.SetOkay EndIndexResetToCount
     

'         End If

'     End If
    
'     Dim mySelector As String
'     If Not VBA.IsMissing(ipStartIndex) Then mySelector = mySelector & "S"
'     If Not VBA.IsMissing(ipRun) Then mySelector = mySelector & "R"
'     If Not VBA.IsMissing(ipEndIndex) Then mySelector = mySelector & "E"

'     Select Case mySelector
        
'         Case "SRE", "SR"
            
'             DeduceStartRunFromAnyStartRun ipStartIndex, ipRun, myRank, myItem
                    
                
'         Case "SE"
                
'             DeduceStartRunFromAnyStartEnd ipStartIndex, ipEndIndex, myRank, myItem
                    
                
'         Case "S"
                
'             DeduceStartRunFromAnyStartOnly ipStartIndex, myRank, myItem
                    
                
'         Case "R"
                
'             DeduceStartRunFromAnyRunOnly ipRun, myRank, myItem
                    
                
'         Case "E"
                
'             DeduceStartRunFromAnyEndOnly ipEndIndex, myRank, myItem
            

'         Case "RE"
                    
'             DeduceStartRunFromAnyRunEnd ipRun, ipEndIndex, myRank, myItem
            
                
'         Case ""
                
'             DeduceStartRunFromNoStartRunOrEnd myRank, myItem
                    
                
'     End Select
        
'     'iopResult.Item(ResultItemsEnums.StartRun(srStartIndex)) = iopResult.Item(ResultItemsEnums.StartRun(srStartIndex)) + myFirstIndex - 1
'      TryStartRunFromAnyStartRunEnd = True
'      Res.SetOkay
' End Function

' Private Function ValidateStartRunEnd _
' ( _
'     ByRef ipIterable As Box, _
'     ByVal ipStart As Long, _
'     ByVal ipRun As Long, _
'     ByVal ipEnd As Long _
' ) As Boolean
'     ValidateStartRunEnd = False
    
'     ' Start must be within bounds
'     ValidateStartRunEnd = True
' End Function
    

' Friend Sub DeduceStartRunFromAnyStartRun _
' ( _
'     ByVal ipStartIndex As Long, _
'     ByVal ipRun As Long, _
'     ByVal ipRank As Long, _
'     ByRef ipEnumerator As Enumerator _
' )
        
'     Dim myStartIndex As Long
'     Dim myRun As Long
'     If ipStartIndex > 0 And ipRun > 0 Then
            
'         If ipStartIndex + ipRun - 1 <= ipEnumerator.Count(ipRank) Then
                
'             myRun = ipRun
'             myStartIndex = ipStartIndex + ipEnumerator.FirstIndex(ipRank) - 1
                                
'         Else
                
'             myRun = ipEnumerator.Count(ipRank) - ipStartIndex + 1
'             myStartIndex = ipStartIndex
'             Res.SetOkay RunLengthTooLong
                
                
'         End If
            
'     ElseIf ipStartIndex > 0 And ipRun < 0 Then
            
'         If ipStartIndex + ipRun + 1 > 0 Then
                
'             myStartIndex = ipEnumerator.FirstIndex
'             myRun = VBA.Abs(ipRun)
                
'         Else
                
'             myStartIndex = 1
'             myRun = ipStartIndex
'             Res.SetOkay RunLengthTooLong
                
'         End If
            
            
'     ElseIf ipStartIndex < 0 And ipRun > 0 Then
            
'         myStartIndex = ipEnumerator.LastIndex(ipRank) + ipStartIndex + 1
'         If myStartIndex + ipRun - 1 <= ipEnumerator.Count(ipRank) Then
                
'             myRun = ipRun
                
'         Else
                
'             myRun = ipEnumerator.Count(ipRank) - myStartIndex + 1
'             Res.SetOkay RunLengthTooLong
                
                
'         End If
            
'     Else                                         ' ipstartindex <0 and ipRun <0 then
            
'         myStartIndex = ipEnumerator.LastIndex(ipRank) + ipStartIndex + 1
'         If myStartIndex + Abs(ipRun) + 1 > 0 Then
                
'             myStartIndex = myStartIndex + ipRun + 1
'             myRun = VBA.Abs(ipRun)
                
'         Else
                
'             myStartIndex = 1
'             myRun = ipEnumerator.Count(ipRank) - myStartIndex + 1
                
'         End If
            
            
'     End If
        
'     ' ' Adjust startindex to align witth the true first index
'     ' myStartIndex = myStartIndex - 1 + iopResult.Item(ResultItemsEnums.Iterator(IsFirstIndex))
'      '@Ignore FunctionReturnValueDiscarded
'      Res.Reset.SetItems myStartIndex, myRun '(Okay)
        
' End Sub


' Friend Sub DeduceStartRunFromAnyStartEnd _
' ( _
'     ByVal ipStartIndex As Long, _
'     ByVal ipEndIndex As Long, _
'     ByVal ipRank As Long, _
'     ByRef ipExtent As Enumerator _
' )
            
'     Dim myStartIndex As Long
'     Dim myEndIndex As Long
'     Dim myRun As Long
            
'     '@Ignore IIfSideEffect
'     '@Ignore IIfSideEffect
'     myStartIndex = IIf(ipStartIndex > 0, ipStartIndex + ipExtent.FirstIndex(ipRank) - 1, ipExtent.LastIndex(ipRank) + ipStartIndex + 1)
'     '@Ignore IIfSideEffect
'     '@Ignore IIfSideEffect
'     myEndIndex = IIf(ipEndIndex > 0, ipEndIndex + ipExtent.FirstIndex(ipRank) - 1, ipExtent.LastIndex(ipRank) + ipEndIndex + 1)
            
'     If myStartIndex < myEndIndex Then
                
'         myRun = myEndIndex - myStartIndex + 1
                
'     Else
                
'         myRun = myStartIndex - myEndIndex + 1
'         myStartIndex = myEndIndex
                
'     End If
            
'     ' ' Adjust startindex to align witth the true first index
'     ' myStartIndex = myStartIndex - 1 + iopResult.Item(ResultItemsEnums.Iterator(IsFirstIndex))
'     '@Ignore FunctionReturnValueDiscarded
'     Res.Reset.SetItems myStartIndex, myRun '(Okay)
            
' End Sub


' Friend Sub DeduceStartRunFromAnyStartOnly _
' ( _
'     ByVal ipStartIndex As Long, _
'     ByVal ipRanks As Long, _
'     ByRef ipExtent As Enumerator _
' )

'     Dim myStartIndex As Long
'     Dim myRun  As Long
'     If ipStartIndex > 0 Then
            
'         myStartIndex = ipExtent.FirstIndex + ipStartIndex - 1
'         myRun = ipExtent.Count - ipStartIndex + 1
            
'     Else
        
'         myStartIndex = ipExtent.LastIndex + ipStartIndex + 1
'         myRun = VBA.Abs(ipStartIndex)
            
'     End If

'     ' ' Adjust startindex to align witth the true first index
'     ' myStartIndex = myStartIndex - 1 + iopResult.Item(ResultItemsEnums.Iterator(IsFirstIndex))
'     '@Ignore FunctionReturnValueDiscarded
'     Res.Reset.SetItems myStartIndex, myRun '(Okay)

' End Sub


' Friend Sub DeduceStartRunFromAnyRunOnly(ByVal ipRun As Long, ByVal ipRank As Long, ByRef ipExtent As Enumerator)
            
'     Dim myStartIndex As Long
'     Dim myRun As Long
    
'     If ipRun > 0 Then
      
'       myStartIndex = ipExtent.FirstIndex(ipRank)
'       myRun = ipRun
      
      
'     Else
            
'       myStartIndex = ipExtent.LastIndex + ipRun + 1
'       myRun = VBA.Abs(ipRun)
            
            
'     End If
        
'     ' ' Adjust startindex to align witth the true first index
'     ' myStartIndex = myStartIndex - 1 + iopResult.Item(ResultItemsEnums.Iterator(IsFirstIndex))
'     '@Ignore FunctionReturnValueDiscarded
'     Res.Reset.SetItems myStartIndex, myRun 'Okay)
        
' End Sub


' Friend Sub DeduceStartRunFromAnyEndOnly(ByVal ipEndIndex As Long, ByVal ipRank As Long, ByRef ipEnumerator As Enumerator)
        
'     Dim myStartIndex As Long
'     Dim myRun As Long
            
'     If ipEndIndex > 0 Then
        
'       myStartIndex = ipEnumerator.FirstIndex(ipRank)
'       myRun = ipEndIndex
        
            
'     Else
      
'       myStartIndex = ipEnumerator.FirstIndex(ipRank)
'       myRun = ipEnumerator.Count(ipRank) + ipEndIndex + 1
            
            
'     End If
        
'     ' Adjust startindex to align witth the true first index
        
'     '@Ignore FunctionReturnValueDiscarded
'     Res.Reset.SetItems myStartIndex, myRun '(Okay)
        
' End Sub


' '@Ignore EmptyMethod
' 'need to add code ignore for now
' Friend Sub DeduceStartRunFromAnyRunEnd(ByVal ipRun As Long, ByVal ipEndIndex As Long, ByVal ipRamk As Long, ByVal myExtent As Enumerator)
'     'To add code
' End Sub


' Friend Sub DeduceStartRunFromNoStartRunOrEnd(ByVal ipRank As Long, ByRef ipExtent As Enumerator)
        
'     ' Dim myFirstIndex As Long
'     ' myFirstIndex = ipEnumerator.FirstIndex
        
'     ' Dim myRun As Long
'     ' myRun = iopExtent.Item(ResultItemsEnums.Iterator(ieCount))
        
'     ' Adjust startindex to align witth the true first index
'     'myStartIndex = myStartIndex - 1 + iopResult.Item(ResultItemsEnums.Iterator(IsFirstIndex))
'     '@Ignore FunctionReturnValueDiscarded
'     Res.SetOkay.SetItems ipExtent.FirstIndex(ipRank), ipExtent.Count(ipRank)

' End Sub

' '@Description("Returns the specified subrange in Item(0) of the result")
' Public Function TryRange _
' ( _
'     ByVal ipIterable As Variant, _
'     Optional ByVal ipStartIndex As Variant, _
'     Optional ByVal ipRun As Variant, _
'     Optional ByVal ipEndIndex As Variant, _
'     Optional ByVal ipRank As Variant _
' ) As Result
' Attribute TryRange.VB_Description = "Returns the specified subrange in Item(0) of the result"
   
'     '@Ignore FunctionReturnValueDiscarded
'     Res.SetOkay
'     Set TryRange = Res
'     Dim myRank As Long
'     If VBA.IsMissing(ipRank) Then myRank = enums.Preset.Value(ArrayFirstRank) Else myRank = ipRank
        
'     If Me.TryStartRunFromAnyStartRunEnd(ipIterable, ipStartIndex, ipRun, ipEndIndex, myRank) Then Exit Function
 
        
'     'get startrun
'     Dim myRangeStartRun As ResultStartRun
'     Set myRangeStartRun = Res.Items
    
'    Dim myColl As Collection
'    Set myColl = New Collection
   
'     'ToDo: Check if this loop adequatly deals with stacks and queues
'     Dim myIndex As Long
'     For myIndex = myRangeStartRun.StartIndex To myRangeStartRun.LastIndex
          
'        myColl.Add ipIterable(myIndex)
            
'     Next
    
'     Set Res.Items.Item(1) = myColl
'     Set TryRange = Res.SetOkay
    
' End Function


' Public Function GetIterantFromAnyStartRunEnd _
' ( _
'     ByVal ipFirst As Long, _
'     ByVal ipLast As Long, _
'     Optional ByVal ipStart As Variant = Empty, _
'     Optional ByVal ipRun As Variant = Empty, _
'     Optional ByVal ipStop As Variant = Empty _
' ) As Iterant

'     If VBA.IsEmpty(ipStart) And VBA.IsEmpty(ipRun) And VBA.IsEmpty(ipStop) Then
        
'         GetIterantFromAnyStartRunEnd = DecodeIterantFromFirstLast(ipFirst, ipLast)
        
        
'     ElseIf VBA.IsEmpty(ipStart) And VBA.IsEmpty(ipRun) And (Not VBA.IsEmpty(ipStop)) Then
       
'         GetIterantFromAnyStartRunEnd = DecodeIterantFromStop(ipFirst, ipLast, ipStop)
        
        
'     ElseIf VBA.IsEmpty(ipStart) And VBA.IsEmpty(ipRun) And VBA.IsEmpty(ipStop) Then
        
'     '     GetIterantFromAnyStartRunEnd = DecodeIterantFrom
        
        
'     ' ElseIf VBA.IsEmpty(ipStart) And VBA.IsEmpty(ipRun) And VBA.IsEmpty(ipStop) Then
        
'     '     GetIterantFromAnyStartRunEnd = DecodeIterantFrom
        
        
'     ' ElseIf VBA.IsEmpty(ipStart) And VBA.IsEmpty(ipRun) And VBA.IsEmpty(ipStop) Then
        
'     '     GetIterantFromAnyStartRunEnd = DecodeIterantFrom
        
    
'     ' ElseIf VBA.IsEmpty(ipStart) And VBA.IsEmpty(ipRun) And VBA.IsEmpty(ipStop) Then
        
'     '     GetIterantFromAnyStartRunEnd = DecodeIterantFrom
    
    
'     ' ElseIf VBA.IsEmpty(ipStart) And VBA.IsEmpty(ipRun) And VBA.IsEmpty(ipStop) Then
        
'     '     GetIterantFromAnyStartRunEnd = DecodeIterantFrom
    
    
'     ' ElseIf VBA.IsEmpty(ipStart) And VBA.IsEmpty(ipRun) And VBA.IsEmpty(ipStop) Then
        
'     '     GetIterantFromAnyStartRunEnd = DecodeIterantFrom
    
    
   
' End If

' End Function

' Friend Function DecodeIterantFromFirstLast(ByVal ipFirst As Long, ByVal ipLast As Long) As Iterant
    
'     Dim myIterant As Iterant
'     Set myIterant = Iterant.Deb
'     With myIterant
    
'         .Start = ipFirst
'         .Stop = ipLast
'         If .Start <.Stop Then
            
'             .Count = Abs(.Stop -.Start + 1)
'             .Step = 1
'         Else
            
'             .Count = Abs(.Start - .Stop + 1)
'             .Step = -1
'         End If
        
'     End With
    
'     Set DecodeIterantFromFirstLast = myIterant
    
' End Function

' Friend Function DecodeIterantFromStop(ByVal ipFirst As Long, ByVal iplast As Long, ByVal ipStop As Long) As Iterant
    
'     Dim myIterant As Iterant
'     Set myIterant = Iterant.Deb
    
'     Dim myFirst As Long
'     ' Dim myLast As Long
'     ' Dim myStop As Long
    
'     ' 'rebase so First = 1
'     ' myFirst = 0
'     ' myLast = iplast - ipFirst
'     ' myStop = ipStop - ipFirst
    
'     ' If myStop < myFirst Then myStop = myFirst
'     ' If myStop > myLast Then myStop = myLast
    


'     With myIterant
    
       
'         .Stop = ipStop
'         .Start = ipFirst
'         If ipFirst < iplast Then
            
            
'             .Count = Abs(ipStop - ipFirst) + 1
'             .Step = 1
'         Else
            
           
'             .Count = Abs(ipFirst - ipStop) + 1
'             .Step = -1
'         End If
        
'     End With
    
    
    
'     Set DecodeIterantFromStop = myIterant
' End Function