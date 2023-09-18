VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Extent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const MY_TYPENAME                       As String = "Extent"

Private Type Properties
   
    GroupType                               As GroupType
   ' Host                                    As variant
    Ranks                                   As seq
   ' Keys                                    As Variant
    
End Type

Private p                                   As Properties

'@DefaultMember
Public Function Deb(ByRef ipEnumerable As Variant) As Extent
    With New Extent
        Return .ConstructInstance(ipEnumerable)
    End With
End Function

Friend Function ConstructInstance(ByRef ipEnumerable As Variant) As Extent
    
    Dim myBaseType As String = TypeInfo.BaseType(ipEnumerable)
    Set p.Ranks = Seq.Deb
    
    Select Case True
        
        Case VBA.IsArray(ipEnumerable)
        
            p.GroupType = OfArray
            Dim myRankCount As Long = ArrayInfo.Ranks(ipEnumerable)
           
            Dim myRank As Long
            For myRank = 1 To myRankCount
                
                p.Ranks.Add Rank.Deb(LBound(ipEnumerable, myRank), UBound(ipEnumerable, myRank))
                
            Next
         
        '    p.Host = ipEnumerable

            
        Case TypeInfo.Group(OfItemByKeysForeach).HoldsKey(myBaseType)
        
            p.GroupType = OfItemByKeysForeach
            p.Ranks.Add Rank.Deb(0, CLng(ipEnumerable.Count - 1))
       
            
        Case TypeInfo.Group(OfItemByToArrayForEach).HoldsKey(TypeInfo.BaseType(ipEnumerable))
        
            p.GroupType = OfItemByToArrayForEach
            
            If TypeInfo.Group(OfOneBasedEnumerables).HoldsKey(myBaseType) Then
            	p.Ranks.Add Rank.Deb(1, CLng(ipEnumerable.count))
            Else
                p.Ranks.Add Rank.Deb(0, CLng(ipEnumerable.count - 1))
           End If
           
           
        Case TypeInfo.Group(OfItemByForEach).HoldsKey(TypeInfo.BaseType(ipEnumerable))
        
        '    Set p.Host = ipEnumerable
            If TypeInfo.Group(OfOneBasedEnumerables).HoldsKey(myBaseType) Then
                
                'p.GroupType = OfOneBasedEnumerables
                p.Ranks.Add Rank.Deb(1, CLng(ipEnumerable.count))
              

            Else
                
               ' p.GroupType = OfZeroBasedEnumerables
                p.Ranks.Add Rank.Deb(0, CLng(ipEnumerable.count - 1))
                

            End If
            
        Case TypeInfo.IsAString(ipEnumerable)
        
      '      p.Host = Seq.Deb(ipEnumerable).ToArray
            p.GroupType = OfStrings
            p.Ranks.Add Rank.Deb(1, CLng(VBA.Len(ipEnumerable)))
           
            
        Case Else
        
     '       p.Host = Array(ipEnumerable)
            p.Ranks.Add Rank.Deb(0, 0)
       '     p.Host = Array(ipEnumerable)
       ' 
        
    End Select
    
    Return Me
    
End Function

Public Function FirstRank() As Long
    Return 1
End Function

Public Function LastRank() As Long
    Return p.Ranks.Count
End Function

Public Function IsQueryable() As Boolean
    Return p.Ranks.Count > 0
End Function

Public Function IsNotQueryable() As Boolean
    Return p.Ranks.Count < 1
End Function

Public Function HasOneItem() As Boolean
    Return p.Ranks.Count = 1
End Function

Public Function HasAnyItems() As Boolean
    Return p.Ranks.Count > 0
End Function


Public Function HasItems() As Boolean
    Return p.Ranks.Count > 1
End Function

Public Property Get First() As Rank
    Return p.Ranks(1)
End Property


Public Property Get FirstIndex(ByRef ipIndex As Long) As Long
    Return p.Ranks.Item(ipIndex).firstindex
End Property

Public Property Let FirstIndex(ByRef ipIndex As Long, ByRef ipValue As Long)
    p.Ranks.Item(ipIndex).firstindex = ipValue
End Property


Public Property Get LastIndex(ByRef ipIndex As Long) As Long
    Return p.Ranks.Item(ipIndex).Lastindex
End Property

Public Property Let LastIndex(ByRef ipIndex As Long, ByRef ipValue As Long)
    p.Ranks.Item(ipIndex).Lastindex = ipValue
End Property


' Public Property Get StartIndex(ByRef ipIndex As Long) As Long
'     Return p.Ranks.Item(ipIndex).Startindex
' End Property

' Public Property Let StartIndex(ByRef ipIndex As Long, ByRef ipValue As Long)
'     p.Ranks.Item(ipIndex).Startindex = ipValue
' End Property


' Public Property Get StopIndex(ByRef ipIndex As Long) As Long
'     Return p.Ranks.Item(ipIndex).Stopindex
' End Property

' Public Property Let StopIndex(ByRef ipIndex As Long, ByRef ipValue As Long)
'     p.Ranks.Item(ipIndex).Stopindex = ipValue
' End Property


Public Property Get Count(ByRef ipIndex As Long) As Long
    Return p.Ranks.Item(ipIndex).Count
End Property

Public Property Let Count(ByRef ipIndex As Long, ByRef ipValue As Long)
    p.Ranks.Item(ipIndex).Count = ipValue
End Property


' Public Property Get Span(ByRef ipIndex As Long) As Long
'     Return p.Ranks.Item(ipIndex).Span
' End Property

' Public Property Let Span(ByRef ipIndex As Long, ByRef ipValue As Long)
'     p.Ranks.Item(ipIndex).Span = ipValue
' End Property


Public Function RanksCount() As Long
    Return p.Ranks.Count
End Function
    

Public Function Ranks() As Seq
    Return p.Ranks
End Function


Public Function HasRank(ByVal ipRank As Long) As Boolean
    Return (ipRank >= 1) And (ipRank <= p.Ranks.Count)
End Function

Public Function LacksRank(ByVal ipRank As Long) As Boolean
    Return (ipRank < 1) And (ipRank < p.Ranks.Count)
End Function


Public Function GroupType() As GroupType
    Return p.GroupType
End Function


' Public Function Keys() As Variant
'     If p.GroupType = OfItemByKeysForeach Then
'         Return p.Host.keys
'     Else
'         Return Null
'     End If
' End Function

' Public Function Key(ByVal ipindex As Long) As Variant
'     Return p.Keys(ipindex)
' End Function


' Public Function Item(ByRef ipIndex As Long) As Variant
'     If p.GroupType = OfItemByKeysForeach Then
'         Return p.Host(p.Keys(ipIndex))
        
'     Else
'         Return p.Host(p.FirstIndex + ipIndex)
'     End If
' End Function

Public Function TYPEName() As String
    TYPEName = MY_TYPENAME
End Function

Public Function ToArray() As Variant
   Return p.Ranks.ToArray
End Function

' '@Enumerator
' Public Property Get NewEnum() As IEnumVARIANT
'     Return p.Ranks.[_NewEnum]
' End Property