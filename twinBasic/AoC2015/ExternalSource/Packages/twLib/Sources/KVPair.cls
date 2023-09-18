VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KVPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A simple single item dictionary"
'@IgnoreModule AssignedByValParameter
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========COption Explicit
'@Moduledescription("A simple single item dictionary")
' Originally this class was implemented to emulate a tuple.  However it was an
' oddity when trying to triage a paramarray and consequently it was transformed
' into a single item dictionary so that it could be treated a member of the
' IterableKeysByEnum.
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Kvp")

Private Type Properties

    Key                                 As Variant
    Item                                As Variant
        
End Type

Private p               As Properties

'@DefaultMember
Public Function Deb(Optional ByVal ipKVPairOrKey As Variant = Empty, Optional ByVal ipItem As Variant = Empty) As KVPair
Attribute Deb.VB_UserMemId = 0
        
        
    With New KVPair
    
        Set Deb = .ConstructInstance(ipKVPairOrKey, ipItem)
    
    
    End With
        
End Function
 
Friend Function ConstructInstance(ByVal ipKVPairOrKey As Variant, ByVal ipItem As Variant) As KVPair

    ReDim p.Key(0 To 0)
    ReDim p.Item(0 To 0)
    
    If TypeName(ipKVPairOrKey) = enums.Types.OrdToTypeName.Item(idKVPair) Then

        If VBA.IsObject(ipKVPairOrKey.Key) Then
            Set p.Key(0) = ipKVPairOrKey.Key
        Else
            p.Key(0) = ipKVPairOrKey.Key
        End If
        If VBA.IsObject(ipKVPairOrKey.Item) Then
            Set p.Item(0) = ipKVPairOrKey.Item
        Else
            p.Item(0) = ipKVPairOrKey.Item
        End If
        
     
        
        
    ElseIf Not VBA.IsEmpty(ipItem) And TypeName(ipKVPairOrKey) <> enums.Types.OrdToTypeName.Item(idKVPair) Then
        
        Sys.Assign p.Key(0), ipKVPairOrKey
        Sys.Assign p.Item(0), ipItem
        
        
    End If
        
   Return Me
    
End Function

Public Property Get Key() As Variant
    Sys.Assign Key, p.Key(0)
End Property

Public Property Let Key(ByVal ipKey As Variant)
    Sys.Assign p.Key(0), ipKey
End Property

' Item takes a key but ignores it as there is only one item in a KVPair
' But the key parameter is required to act as a dictionary
'@Ignore ParameterNotUsed
Public Property Get Item(ByVal ipKey As Variant) As Variant
    Sys.Assign Item, p.Item(0)
End Property

'@Ignore ParameterNotUsed
Public Property Let Item(ByVal ipKey As Variant, ByVal ipItem As Variant)
    Sys.Assign p.Item(0), ipItem
End Property

Public Function ToString(Optional ByVal ipSeparator As Variant) As String
    If VBA.IsMissing(ipSeparator) Then ipSeparator = Char.twComma
        'Debug.Print ipSeparator, p.Key.Item(0), p.Item.Item(0)
    ToString = Stringifier.ToString(p.Key.Item(0)) & ipSeparator & Stringifier.ToString(p.Item.Item(0))
End Function

'@DefaultMember
Public Function Name() As String
Attribute Name.VB_UserMemId = 0
    Name = TypeName(Me)
End Function

'This may seem stupid but a KVPair can be treated as a single item dictionary
' So to conform with other dictionary type we implement Keys and Items functions

Public Function Keys() As Variant
    Keys = p.Key
End Function

Public Function Items() As Variant
    Items = p.Item
End Function

' '@Enumerator
' Public Function NewEnum() As IUnknown
' Attribute NewEnum.VB_UserMemId = -4

'     Dim enumerator As IUnknown
'     '@Ignore DefaultMemberRequired
'     Set enumerator = p.Key.GetEnumerator(0, p.Key.Count)
'     Set NewEnum = enumerator

' End Function

Public Function Count() As Long
    Count = 1
End Function
